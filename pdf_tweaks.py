"""PDF Tweaks 1.1.0 - combine selected documents and images into one PDF.

Requires Python 3.11+ with Tkinter.
Install: py -m pip install -r requirements.txt
GUI: py pdf_tweaks.py
Preview a folder: py pdf_tweaks.py --folder "C:/Example/Files" --list
Combine selected files: py pdf_tweaks.py --files "report.pdf" "photo.png"
Use --help for output filenames, folder selection, ordering and conversion options.

Office formats require Windows and the matching installed desktop Office app.
TXT/CSV/TSV must use UTF-8 or BOM-marked UTF-16; CSV uses commas and TSV tabs.
Images become A4 pages, including every frame of multi-frame images. Excel
uses visible sheets and existing print areas; imported data is one page wide.
Word tracked-change markup is excluded. Review page appearance and links.
Merged copies do not retain the cryptographic validity of source signatures.
Source files are retained; new results are written to separate dated folders.

Created by Talal Faiz: https://github.com/talalondata
Suggestions, bug reports and contributions are welcome; see README.md.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import multiprocessing
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import time
import traceback
from contextlib import contextmanager
from datetime import datetime
from uuid import uuid4


OUTPUT_DIR = "_PDF_Combiner_Output"
VERSION = "1.1.0"
GITHUB_PROFILE_URL = "https://github.com/talalondata"
WORD_CACHE_FOLDER = "_Word_Cache"
# Change this profile when Word export settings change, so earlier PDFs are not reused.
WORD_EXPORT_PROFILE = "native-word-print-all-pages-no-markup-v1"


class WordWorkerError(RuntimeError):
    """Carry the worker's retry decision and retained-export status to the parent."""

    def __init__(self, message: str, retryable: bool = False, export_retained: bool = False,
                 reason: str = "disconnected"):
        super().__init__(message)
        self.retryable = retryable
        self.export_retained = export_retained
        self.reason = reason


class WordExportMissing(RuntimeError):
    """Word returned from export without producing a nonempty PDF."""


class WordCloseDisconnected(RuntimeError):
    """Export returned and its PDF was validated before the close failed."""


class SelectionCancelled(Exception):
    """The user requested cancellation."""


def _exception_chain(error: Exception):
    seen = set()
    while error is not None and id(error) not in seen:
        seen.add(id(error))
        yield error
        error = error.__cause__ or error.__context__


def _is_word_disconnection(error: Exception) -> bool:
    """Recognize RPC_E_DISCONNECTED in nested COM exceptions."""
    for error in _exception_chain(error):
        hresult = getattr(error, "hresult", None)
        if hresult is None and error.args and isinstance(error.args[0], int):
            hresult = error.args[0]
        if isinstance(hresult, int) and hresult & 0xFFFFFFFF == 0x80010108:
            return True
    return False


def _raise_scan_error(error: OSError) -> None:
    raise error


def discover(folder: Path, recursive: bool = False) -> list[Path]:
    """Ignore Word lock files and all previous output folders."""
    if not folder.is_dir():
        raise ValueError(f"Folder does not exist: {folder}")
    entries = []
    # Prune output folders instead of walking every previous result recursively.
    for root, dirs, names in os.walk(folder, followlinks=False, onerror=_raise_scan_error):
        dirs[:] = [d for d in dirs if d.casefold() not in {OUTPUT_DIR.casefold(), WORD_CACHE_FOLDER.casefold()}]
        entries.extend(Path(root) / name for name in names)
        if not recursive:
            break
    files = [
        p for p in entries
        if p.is_file()
        and p.suffix.casefold() in GUI_FORMATS
        and not p.name.startswith("~$")
        and OUTPUT_DIR.casefold() not in {
            part.casefold() for part in p.relative_to(folder).parts[:-1]
        }
    ]
    # Filename first; relative path breaks ties when including subfolders.
    files.sort(key=lambda p: (
        p.name.casefold(), str(p.relative_to(folder)).casefold(), p.name
    ))
    if not files:
        raise ValueError("No supported documents or images were found.")
    return files


def fingerprint(path: Path) -> tuple[int, int]:
    """Return a cheap metadata snapshot; content checks use SHA-256 separately."""
    stat = path.stat()
    return stat.st_size, stat.st_mtime_ns


@contextmanager
def word_application(on_started=None):
    """Use a separate Word instance and close only that instance."""
    if os.name != "nt":
        raise RuntimeError("Word conversion requires Windows and desktop Word.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Install dependencies: py -m pip install pywin32 pypdf") from exc
    pythoncom.CoInitialize()
    word = None
    options = None
    application_settings = {}
    option_settings = {}
    try:
        word = win32com.client.DispatchEx("Word.Application")
        if on_started is not None and on_started(word) is False:
            # Do not change settings or close a pre-existing Word process.
            word = None
            raise RuntimeError("Word reused an existing process; close Word and rerun.")
        application_settings = {
            name: getattr(word, name)
            for name in ("Visible", "DisplayAlerts", "AutomationSecurity")
        }
        options = word.Options
        option_settings = {
            name: getattr(options, name)
            for name in ("UpdateLinksAtOpen", "UpdateFieldsAtPrint", "UpdateLinksAtPrint")
        }
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # Disable document macros.
        options.UpdateLinksAtOpen = False
        options.UpdateFieldsAtPrint = False
        options.UpdateLinksAtPrint = False
        yield word
    finally:
        had_error = sys.exc_info()[0] is not None
        cleanup_errors = []
        try:
            if word is not None:
                for target, settings in ((options, option_settings), (word, application_settings)):
                    for name, value in settings.items():
                        try:
                            setattr(target, name, value)
                        except Exception as exc:
                            cleanup_errors.append(f"Restoring {name}: {exc}")
                try:
                    word.Quit(SaveChanges=0)
                except Exception as exc:
                    cleanup_errors.append(f"Closing Word: {exc}")
        finally:
            pythoncom.CoUninitialize()
        if cleanup_errors:
            message = "; ".join(cleanup_errors)
            if had_error:
                logging.warning("Additional Word cleanup error: %s", message)
            else:
                raise RuntimeError(message)


def _validate_exported_pdf(path: Path) -> None:
    """Read the page tree and content streams before retaining an export."""
    from pypdf import PdfReader
    with path.open("rb") as stream:
        reader = PdfReader(stream, strict=True)
        if reader.is_encrypted:
            raise ValueError("Exported PDF is encrypted.")
        if len(reader.pages) == 0:
            raise ValueError("Exported PDF contains no pages.")
        for page in reader.pages:
            contents = page.get_contents()
            if contents is not None:
                contents.get_data()


def export_word(word, source: Path, destination: Path) -> None:
    """Export all pages using Word's own PDF engine; do not save the source."""
    document = None
    try:
        # A random non-empty password makes password-protected files fail
        # instead of asking for a password in an invisible Word window.
        dummy_password = uuid4().hex
        text_options = {"Encoding": text_encoding(source)[1]} if source.suffix.casefold() == ".txt" else {}
        document = word.Documents.Open(
            FileName=str(source), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False, OpenAndRepair=False,
            PasswordDocument=dummy_password, WritePasswordDocument=dummy_password,
            PasswordTemplate=dummy_password, NoEncodingDialog=True, **text_options,
        )
        document.ExportAsFixedFormat(
            OutputFileName=str(destination), ExportFormat=17,
            OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
            IncludeDocProps=True, KeepIRM=True, CreateBookmarks=1,
            DocStructureTags=True, BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        if not destination.is_file() or destination.stat().st_size == 0:
            raise WordExportMissing(f"Word did not produce a nonempty PDF: {source.name}")
    finally:
        if document is not None:
            had_error = sys.exc_info()[0] is not None
            try:
                document.Close(SaveChanges=0)
            except Exception as exc:
                if had_error:
                    logging.warning("Additional close error for %s: %s", source.name, exc)
                elif _is_word_disconnection(exc):
                    # Reaching Close without an earlier exception means that
                    # ExportAsFixedFormat returned and produced a nonempty file.
                    # Never accept an export that raised, or an unreadable PDF.
                    try:
                        _validate_exported_pdf(destination)
                    except Exception as validation_error:
                        raise RuntimeError(f"Cannot validate PDF exported from {source.name} "
                                           f"after Word disconnected: {validation_error}") from validation_error
                    raise WordCloseDisconnected(
                        f"Retained the validated PDF for {source.name}; "
                        "Word disconnected while closing the document."
                    ) from exc
                else:
                    raise


def _word_process_id(word) -> int:
    """Identify the conversion process through a temporary Word document window.

    Create a disposable blank document to obtain a document-window handle
    before opening any input. Close the blank document without saving.
    """
    import win32process
    probe = None
    try:
        probe = word.Documents.Add(Visible=True)
        hwnd = int(probe.ActiveWindow.Hwnd) & 0xFFFFFFFF
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if not pid:
            raise RuntimeError("Could not identify the conversion Word instance.")
        return pid
    finally:
        if probe is not None:
            had_error = sys.exc_info()[0] is not None
            try:
                probe.Close(SaveChanges=0)
            except Exception as exc:
                if had_error:
                    logging.warning("Additional blank-document close error: %s", exc)
                else:
                    raise


class _WorkerWarningHandler(logging.Handler):
    def __init__(self, messages):
        super().__init__(logging.WARNING)
        self.messages = messages

    def emit(self, record):
        try:
            self.messages.send(("warning", self.format(record)))
        except (OSError, EOFError):
            pass  # The parent may already have cancelled and closed its pipe.


def _word_worker(sources: list[str], destination: str, existing_pids: set[int], messages) -> None:
    """Run COM calls in a child process so the parent can enforce timeouts."""
    export_retained = False
    log_handler = _WorkerWarningHandler(messages)
    logging.getLogger().addHandler(log_handler)
    try:
        def ready(word):
            pid = _word_process_id(word)
            if pid in existing_pids:
                return False
            messages.send(("ready", pid))
            return True

        with word_application(on_started=ready) as word:
            for index, source_text in enumerate(sources, 1):
                source = Path(source_text)
                messages.send(("working", source.name))
                pdf_path = Path(destination) / f"word_{index:06d}.pdf"
                try:
                    export_word(word, source, pdf_path)
                except WordCloseDisconnected as exc:
                    messages.send(("converted", source_text, str(pdf_path)))
                    export_retained = True
                    messages.send(("warning", str(exc)))
                    raise
                except Exception as exc:
                    raise RuntimeError(f"Cannot convert {source.name}: {exc}") from exc
                messages.send(("converted", source_text, str(pdf_path)))
        messages.send(("success",))
    except Exception as exc:
        missing = any(isinstance(error, WordExportMissing) for error in _exception_chain(exc))
        messages.send(("error", traceback.format_exc(), missing or _is_word_disconnection(exc),
                       export_retained, "missing_pdf" if missing else "disconnected"))
    finally:
        logging.getLogger().removeHandler(log_handler)
        log_handler.close()
        messages.close()


def _check_cancel(cancel_event=None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise SelectionCancelled("Cancelled by user.")


def _wait_for_word(process, messages, timeout: float, on_message, cancel_event=None,
                   worker_label="Word") -> None:
    """Bound each startup/conversion/shutdown stage, including hidden dialogs."""
    deadline = time.monotonic() + timeout
    current = f"{worker_label} startup"
    while True:
        _check_cancel(cancel_event)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise RuntimeError(f"Timed out after {timeout:g} seconds: {current}.")
        if messages.poll(min(remaining, 0.2)):
            try:
                event = messages.recv()
            except EOFError as exc:
                raise RuntimeError(f"{worker_label} worker exited unexpectedly: {current}.") from exc
            kind = event[0]
            if kind == "error":
                raise WordWorkerError(event[1], retryable=len(event) > 2 and event[2],
                                      export_retained=len(event) > 3 and event[3],
                                      reason=event[4] if len(event) > 4 else "disconnected")
            if kind == "success":
                return
            on_message(event)
            current = event[1] if kind == "working" else f"{worker_label} conversion/shutdown"
            deadline = time.monotonic() + timeout
        elif not process.is_alive():
            raise RuntimeError(f"{worker_label} worker exited unexpectedly: {current}.")


def _run_word_batch(sources: list[Path], destination: Path, timeout: float,
                    converted: dict[Path, Path], total: int, cancel_event=None) -> None:
    if os.name != "nt":
        raise RuntimeError("Word conversion requires Windows and desktop Word.")
    try:
        import win32api
        import win32event
        import win32process
    except ImportError as exc:
        raise RuntimeError("Install dependencies: py -m pip install pywin32 pypdf") from exc
    # Never terminate a process that was already running before this job.
    existing_pids = set(win32process.EnumProcesses())
    context = multiprocessing.get_context("spawn")
    receive, send = context.Pipe(duplex=False)
    process = context.Process(target=_word_worker, args=(
        [str(p) for p in sources], str(destination), existing_pids, send,
    ))
    word_handle = None
    success = False

    def on_message(event):
        nonlocal word_handle
        if event[0] == "ready":
            pid = event[1]
            if pid in existing_pids:
                raise RuntimeError("Word reused an existing process; close Word and rerun.")
            # Hold a handle to this exact process, avoiding later PID reuse.
            # PROCESS_TERMINATE | SYNCHRONIZE: also check if Word has exited.
            word_handle = win32api.OpenProcess(0x100001, False, pid)
        elif event[0] == "working":
            logging.info("Converting [%d/%d] %s", len(converted) + 1, total, event[1])
        elif event[0] == "converted":
            converted[Path(event[1])] = Path(event[2])
        elif event[0] == "warning":
            logging.warning("%s", event[1])
        else:
            raise RuntimeError(f"Unexpected Word worker event: {event[0]}")

    try:
        process.start()
        send.close()
        _wait_for_word(process, receive, timeout, on_message, cancel_event)
        process.join(timeout=5)
        if process.is_alive() or process.exitcode != 0:
            raise RuntimeError("Word worker did not shut down cleanly.")
        if any(source not in converted for source in sources):
            raise RuntimeError("Some Word documents were not converted.")
        success = True
    finally:
        if not success and word_handle is not None:
            try:
                if win32event.WaitForSingleObject(word_handle, 0) == 258:  # WAIT_TIMEOUT
                    try:
                        win32api.TerminateProcess(word_handle, 1)
                    except Exception as exc:
                        code = getattr(exc, "winerror", None)
                        if code is None and exc.args:
                            code = exc.args[0]
                        # Word can exit between the status check and termination.
                        # Suppress error 5 only if this exact process has exited.
                        if code != 5 or win32event.WaitForSingleObject(word_handle, 5000) != 0:
                            raise
                    win32event.WaitForSingleObject(word_handle, 5000)
            except Exception as exc:
                logging.warning("Word process cleanup: %s", exc)
        if process.pid is not None and process.is_alive():
            process.terminate()
            process.join(timeout=5)
            if process.is_alive():
                process.kill()
                process.join(timeout=5)
        if word_handle is not None:
            word_handle.Close()
        receive.close()
        send.close()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _word_cache_descriptor(source: Path) -> tuple[str, dict]:
    # Include the content hash and export profile, not just the source filename.
    before = fingerprint(source)
    identity = {"path": os.path.normcase(str(source.resolve())), "size": before[0],
                "mtime_ns": before[1], "sha256": _sha256_file(source)}
    if fingerprint(source) != before:
        raise RuntimeError(f"Source changed while checking cache: {source.name}")
    key = hashlib.sha256(json.dumps(
        {"source": identity, "profile": WORD_EXPORT_PROFILE}, sort_keys=True,
    ).encode("utf-8")).hexdigest()
    return key, identity


def _read_word_cache(cache_dir: Path, key: str, identity: dict, source: Path):
    manifest = cache_dir / f"{key}.json"
    if not manifest.is_file():
        return None
    try:
        data = json.loads(manifest.read_text(encoding="utf-8"))
        if data.get("schema") != 1 or data.get("source") != identity or data.get("profile") != WORD_EXPORT_PROFILE:
            raise ValueError("Cache identity does not match.")
        pdf_hash = data.get("pdf_sha256", "")
        if not isinstance(pdf_hash, str) or re.fullmatch(r"[0-9a-f]{64}", pdf_hash) is None:
            raise ValueError("Invalid cache checksum.")
        pdf = cache_dir / f"{key}_{pdf_hash}.pdf"
        if _sha256_file(pdf) != pdf_hash:
            raise ValueError("Cached PDF checksum does not match.")
        _validate_exported_pdf(pdf)
        return pdf
    except Exception as exc:
        logging.warning("Cannot reuse saved PDF for %s; converting again: %s", source.name, exc)
        return None


def _save_word_cache(cache_dir: Path, descriptors: dict, converted: dict, saved: set) -> None:
    for source, pdf in list(converted.items()):
        if source in saved:
            continue
        key, identity = descriptors[source]
        staging = cache_dir / f"{key}_{uuid4().hex}.pending"
        staging_manifest = cache_dir / f"{key}_{uuid4().hex}.json.pending"
        try:
            shutil.copyfile(pdf, staging)
            _validate_exported_pdf(staging)
            if (fingerprint(source) != (identity["size"], identity["mtime_ns"])
                    or _sha256_file(source) != identity["sha256"]):
                raise RuntimeError(f"Source changed before saving its PDF: {source.name}")
            pdf_hash = _sha256_file(staging)
            target = cache_dir / f"{key}_{pdf_hash}.pdf"
            data = {"schema": 1, "profile": WORD_EXPORT_PROFILE, "source": identity,
                    "pdf_sha256": pdf_hash}
            staging_manifest.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            # Publish the validated PDF before its manifest can advertise it for reuse.
            os.replace(staging, target)
            os.replace(staging_manifest, cache_dir / f"{key}.json")
            converted[source] = target
            saved.add(source)
        except Exception as exc:
            raise RuntimeError(f"Cannot save validated Word PDF for {source.name}: {exc}") from exc
        finally:
            had_error = sys.exc_info()[0] is not None
            for temporary in (staging, staging_manifest):
                try:
                    temporary.unlink(missing_ok=True)
                except OSError as exc:
                    if had_error:
                        logging.warning("Additional cache cleanup error: %s", exc)
                    else:
                        raise


def convert_word_files(sources: list[Path], destination: Path, timeout: float,
                       cache_dir: Path | None = None, reuse_cache: bool = True,
                       refresh_files: set[Path] | None = None,
                       conversion_status: dict[Path, str] | None = None,
                       cancel_event=None) -> dict[Path, Path]:
    """Restart failed exports; preserve validated conversions across runs."""
    converted = {}
    descriptors = {}
    saved = set()
    refresh_files = set(refresh_files or ())
    if refresh_files.difference(sources):
        raise ValueError("Refresh files must belong to the Word input list.")
    if cache_dir is not None:
        cache_dir.mkdir(parents=True, exist_ok=True)
        for source in sources:
            _check_cancel(cancel_event)
            key, identity = descriptors[source] = _word_cache_descriptor(source)
            force_refresh = source in refresh_files
            if force_refresh:
                logging.info("Requested Word refresh: %s", source.name)
            pdf = _read_word_cache(cache_dir, key, identity, source) if reuse_cache and not force_refresh else None
            if pdf is not None:
                converted[source] = pdf
                saved.add(source)
                if conversion_status is not None:
                    conversion_status[source] = "Reused validated Word PDF"
                logging.info("Reusing validated Word PDF [%d/%d] %s", len(converted), len(sources), source.name)
    logging.info("Word conversion plan: %d reused; %d require conversion.",
                 len(converted), len(sources) - len(converted))
    retry_counts = {}
    batch = 0
    while len(converted) < len(sources):
        _check_cancel(cancel_event)
        pending = [source for source in sources if source not in converted]
        batch += 1
        batch_dir = destination / f"word_batch_{batch:04d}"
        batch_dir.mkdir()
        converted_before = len(converted)
        try:
            _run_word_batch(pending, batch_dir, timeout, converted, len(sources), cancel_event)
        except WordWorkerError as exc:
            pending = [source for source in sources if source not in converted]
            if not exc.retryable:
                raise
            if not pending:
                if exc.reason == "missing_pdf":
                    raise
                logging.warning("Word disconnected during cleanup; all %d exported PDFs "
                                "were retained and will be checked during merging.", len(converted))
                return converted
            current = pending[0]
            if exc.export_retained and len(converted) > converted_before:
                logging.warning("Restarting Word at %s; retaining %d completed conversions.",
                                current.name, len(converted))
                continue
            retry_counts[current] = retry_counts.get(current, 0) + 1
            if retry_counts[current] > 2:
                if exc.reason == "missing_pdf":
                    raise RuntimeError(f"No PDF was produced for {current.name} after 3 attempts; "
                                       "open it in Word and export it to PDF manually for diagnosis.") from exc
                raise RuntimeError(f"Word disconnected repeatedly at {current.name}; "
                                   "open this document in Word and inspect it.") from exc
            message = "Word returned without a PDF" if exc.reason == "missing_pdf" else "Word disconnected"
            logging.warning(
                "%s. Retaining %d completed conversions; restarting at %s (retry %d/2).",
                message, len(converted), current.name, retry_counts[current])
        finally:
            # Retain completed conversions even if a later document fails or is cancelled.
            if conversion_status is not None:
                for source in converted:
                    if source not in conversion_status:
                        conversion_status[source] = (
                            "Converted from Word (requested refresh)" if source in refresh_files
                            else "Converted from Word"
                        )
            if cache_dir is not None:
                had_error = sys.exc_info()[0] is not None
                try:
                    _save_word_cache(cache_dir, descriptors, converted, saved)
                except Exception:
                    if had_error:
                        logging.exception("Could not save some completed Word conversions for reuse.")
                    else:
                        raise
    return converted


def _is_duplicate_pdf_info(message: str) -> bool:
    return re.fullmatch(
        r"Multiple definitions in dictionary at byte 0x[0-9a-fA-F]+ for key /Info",
        message,
    ) is not None


class _PDFRecoveryWarnings(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.problems = []

    def emit(self, record):
        message = record.getMessage()
        if not _is_duplicate_pdf_info(message):
            self.problems.append(message)

    def check(self):
        if self.problems:
            raise ValueError("Additional PDF problems during metadata recovery: "
                             + "; ".join(self.problems[:3]))


def _is_pdf_whitespace_warning(message: str) -> bool:
    return re.fullmatch(
        r"Superfluous whitespace found in object header b'[0-9]+' b'[0-9]+'", message,
    ) is not None


class _ConsolePdfWhitespaceFilter(logging.Filter):
    def filter(self, record):
        return not (record.name.startswith("pypdf.") and _is_pdf_whitespace_warning(record.getMessage()))


class _PDFWhitespaceCounter(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.count = 0

    def emit(self, record):
        if _is_pdf_whitespace_warning(record.getMessage()):
            self.count += 1


@contextmanager
def _pdf_reader_for_merge(stream, source: Path):
    counter = _PDFWhitespaceCounter()
    logger = logging.getLogger("pypdf")
    logger.addHandler(counter)
    try:
        with _pdf_reader_for_merge_impl(stream, source) as result:
            yield result
    finally:
        logger.removeHandler(counter)
        counter.close()
        if counter.count:
            logging.warning("PDF whitespace warnings in %s: %d object headers; "
                            "details are in Processing_Log.txt.", source.name, counter.count)


@contextmanager
def _pdf_reader_for_merge_impl(stream, source: Path):
    """Allow only the observed duplicate /Info error, and record recovery."""
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError
    try:
        reader = PdfReader(stream, strict=True)
    except PdfReadError as exc:
        if not _is_duplicate_pdf_info(str(exc)):
            raise
    else:
        yield reader, False
        return

    # Retry only this known metadata error; other parser problems still stop the merge.
    logging.warning("Recovering duplicate /Info metadata in %s; validating all pages.",
                    source.name)
    collector = _PDFRecoveryWarnings()
    pdf_logger = logging.getLogger("pypdf")
    pdf_logger.addHandler(collector)
    try:
        stream.seek(0)
        reader = PdfReader(stream, strict=False)
        if not reader.is_encrypted:
            for page in reader.pages:
                contents = page.get_contents()
                if contents is not None:
                    contents.get_data()
        collector.check()
        yield reader, True
        # Some objects (including resources/forms) are read only when appended.
        collector.check()
    finally:
        pdf_logger.removeHandler(collector)
        collector.close()


def check_input_pdfs(files: list[Path], cancel_event=None) -> None:
    """Check existing PDFs before starting document conversion."""
    count = 0
    for source in files:
        _check_cancel(cancel_event)
        if source.suffix.casefold() != ".pdf":
            continue
        logging.info("Checking PDF before Word conversion: %s", source.name)
        try:
            with source.open("rb") as stream, _pdf_reader_for_merge(stream, source) as (reader, recovered):
                if reader.is_encrypted:
                    raise ValueError("Encrypted PDF: supply an unencrypted working copy.")
                if len(reader.pages) == 0:
                    raise ValueError("PDF contains no pages.")
                for page in reader.pages:
                    _check_cancel(cancel_event)
                    contents = page.get_contents()
                    if contents is not None:
                        contents.get_data()
                reader.get_fields()
        except SelectionCancelled:
            raise
        except Exception as exc:
            raise RuntimeError(f"Cannot check {source.name}: {exc}") from exc
        count += 1
    if count:
        logging.info("PDF precheck passed: %d files.", count)


def merge_pdfs(items: list[tuple[Path, Path]], temporary_output: Path, cancel_event=None) -> list[dict]:
    """Merge complete documents; verify page count and record page ranges."""
    if not items:
        raise ValueError("No input documents supplied.")
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError as exc:
        raise RuntimeError("Install dependencies: py -m pip install pywin32 pypdf") from exc
    writer = PdfWriter()
    report = []
    expected = 0
    try:
        for sequence, (source, pdf_path) in enumerate(items, 1):
            _check_cancel(cancel_event)
            logging.info("Merging %s", source.name)
            try:
                with pdf_path.open("rb") as stream, _pdf_reader_for_merge(stream, source) as (reader, recovered):
                    if reader.is_encrypted:
                        raise ValueError("Encrypted PDF: supply an unencrypted working copy.")
                    pages = len(reader.pages)
                    if pages == 0:
                        raise ValueError("PDF contains no pages.")
                    # PDFs may reuse names such as "Name"; keep each form's fields separate.
                    if reader.get_fields():
                        reader.add_form_topname(f"document_{sequence:04d}")
                    writer.append(reader, outline_item=source.name)
                    report.append({
                        "Order": sequence, "Filename": source.name,
                        "Source": str(source), "Pages": pages,
                        "First page": expected + 1, "Last page": expected + pages,
                        "PDF handling": "Duplicate /Info recovery" if recovered else "Strict",
                    })
                    expected += pages
            except Exception as exc:
                raise RuntimeError(f"Cannot merge {source.name}: {exc}") from exc
        _check_cancel(cancel_event)
        writer.write(str(temporary_output))
    finally:
        writer.close()
    with temporary_output.open("rb") as stream:
        actual = len(PdfReader(stream, strict=True).pages)
    if actual != expected:
        raise RuntimeError(f"Page-count mismatch: expected {expected}, found {actual}.")
    return report


WORD_FORMATS = {".doc", ".docx", ".docm", ".rtf", ".odt", ".txt"}
EXCEL_FORMATS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv"}
POWERPOINT_FORMATS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".odp"}
IMAGE_FORMATS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
GUI_FORMATS = {".pdf"} | WORD_FORMATS | EXCEL_FORMATS | POWERPOINT_FORMATS | IMAGE_FORMATS


def ordered_files(paths, mode="filename") -> list[Path]:
    """Return a deterministic order; manual queue order changes only on request."""
    def natural_key(path):
        pieces = re.split(r"([0-9]+)", path.name.casefold())
        return tuple((1, int(piece)) if piece.isdecimal() else (0, piece) for piece in pieces)

    if mode == "filename":
        return sorted(paths, key=lambda p: (p.name.casefold(), str(p).casefold()))
    if mode == "natural":
        return sorted(paths, key=lambda p: (natural_key(p), str(p).casefold()))
    if mode == "selected":
        return list(paths)
    raise ValueError(f"Unknown sorting mode: {mode}")


def validate_output_name(value: str) -> str:
    """Accept a Windows-safe filename, never a path or reserved device name."""
    name = value.strip()
    if not name or name.endswith((".", " ")):
        raise ValueError("Enter a PDF filename, such as My_Document_Pack.pdf.")
    if any(ord(char) < 32 or char in '<>:"/\\|?*' for char in name):
        raise ValueError("PDF filename cannot contain paths or these characters: < > : \" / \\ | ? *")
    if not name.casefold().endswith(".pdf"):
        name += ".pdf"
    if not name[:-4].strip(" ."):
        raise ValueError("Enter a name before .pdf.")
    reserved = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}
    reserved.update(f"{prefix}{number}" for prefix in ("COM", "LPT") for number in range(1, 10))
    if name.split(".", 1)[0].strip().upper() in reserved:
        raise ValueError("That filename is reserved by Windows; choose another name.")
    if len(name) > 180:
        raise ValueError("Keep the PDF filename within 180 characters.")
    return name


def default_result_root() -> Path:
    for candidate in (Path.home() / "Downloads", Path.home() / "Documents", Path.home()):
        if candidate.is_dir():
            return candidate / OUTPUT_DIR
    return Path.home() / OUTPUT_DIR


class OutputSetupError(RuntimeError):
    """The chosen result location could not be prepared."""


def create_run_directory(result_root: Path) -> Path:
    try:
        result_root.mkdir(parents=True, exist_ok=True)
        run_dir = result_root / (datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid4().hex[:8])
        run_dir.mkdir()
        return run_dir
    except OSError as exc:
        raise OutputSetupError(
            f"Cannot create the result folder inside:\n{result_root}\n\n"
            "Use Browse to choose another existing folder, such as Downloads, then try again.\n\n"
            f"System message: {exc}"
        ) from exc


def write_startup_diagnostic(details: str) -> str:
    """Save a real diagnostic when the selected output folder cannot hold a log."""
    try:
        with tempfile.NamedTemporaryFile(mode="w", prefix="pdf_combiner_error_", suffix=".txt",
                                         encoding="utf-8", delete=False) as stream:
            stream.write(datetime.now().isoformat() + "\n" + details)
            return stream.name
    except OSError:
        return ""


def file_kind(path: Path) -> str:
    extension = path.suffix.casefold()
    for formats, label in ((WORD_FORMATS, "Word / text"), (EXCEL_FORMATS, "Excel / data"),
                           (POWERPOINT_FORMATS, "PowerPoint"), (IMAGE_FORMATS, "Image")):
        if extension in formats:
            return label
    return "PDF" if extension == ".pdf" else "Unsupported"


def validate_selected(paths) -> list[Path]:
    """Validate the exact queue without scanning or adding other documents."""
    files, seen = [], set()
    for value in paths:
        source = Path(value).expanduser().resolve(strict=True)
        if not source.is_file():
            raise ValueError(f"Not a regular file: {source}")
        if source.name.startswith("~$"):
            raise ValueError(f"Select the document, rather than its Office lock file: {source.name}")
        if source.suffix.casefold() not in GUI_FORMATS:
            raise ValueError(f"Unsupported file type: {source.name}. Export this file to PDF first.")
        if not source.stat().st_size:
            raise ValueError(f"File is empty: {source.name}")
        key = os.path.normcase(str(source))
        if key in seen:
            raise ValueError(f"This file is already included: {source}")
        seen.add(key)
        files.append(source)
    if not files:
        raise ValueError("Choose at least one file using Add files.")
    return files


class FileSelection:
    """Store the GUI queue; removal and replacement change this list only."""
    def __init__(self):
        self.files: list[Path] = []

    def same_document(self, source: Path) -> list[int]:
        # Treat another extension as a possible replacement only within the same folder.
        return [i for i, old in enumerate(self.files)
                if old.parent == source.parent and old.stem.casefold() == source.stem.casefold()
                and old != source]

    def add(self, source: Path, replace=False) -> None:
        source = validate_selected([source])[0]
        if source in self.files:
            return
        matches = self.same_document(source) if replace else []
        if matches:
            position = matches[0]
            self.files = [p for i, p in enumerate(self.files) if i not in matches]
            self.files.insert(position, source)
        else:
            self.files.append(source)

    def replace(self, index: int, source: Path) -> None:
        source = validate_selected([source])[0]
        if source in self.files and self.files[index] != source:
            raise ValueError("The replacement file is already included elsewhere in the list.")
        self.files[index] = source

    def move(self, indices, direction: int) -> list[int]:
        selected = set(indices)
        if direction not in {-1, 1}:
            raise ValueError("Move direction must be -1 or 1.")
        for index in sorted(selected, reverse=direction > 0):
            other = index + direction
            if 0 <= other < len(self.files) and other not in selected:
                self.files[index], self.files[other] = self.files[other], self.files[index]
                selected.remove(index)
                selected.add(other)
        return sorted(selected)


def text_encoding(source: Path) -> tuple[str, int]:
    """Accept explicit UTF BOMs or valid UTF-8; fail rather than guess an encoding."""
    with source.open("rb") as stream:
        prefix = stream.read(4)
    if prefix.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        raise ValueError(f"Save {source.name} as UTF-8 or UTF-16 text first.")
    encoding, codepage = ("utf-16", 1200 if prefix.startswith(b"\xff\xfe") else 1201) \
        if prefix.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", 65001)
    try:
        with source.open("r", encoding=encoding) as stream:
            for chunk in iter(lambda: stream.read(65536), ""):
                if "\x00" in chunk:
                    raise ValueError(f"Unexpected binary content in text file: {source.name}")
    except UnicodeError as exc:
        raise ValueError(f"Cannot decode {source.name}. Save it as UTF-8 text and try again.") from exc
    return encoding, codepage


def convert_image(source: Path, destination: Path, cancel_event=None) -> int:
    """Place every image frame on an A4 page, without resampling image pixels."""
    from PIL import Image, ImageOps
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.pagesizes import A4, landscape
    canvas = Canvas(str(destination), pagesize=A4, pageCompression=1)
    count = 0
    try:
        with Image.open(source) as image:
            for index in range(getattr(image, "n_frames", 1)):
                _check_cancel(cancel_event)
                image.seek(index)
                oriented = ImageOps.exif_transpose(image)
                rgba = oriented.convert("RGBA")
                rgb = Image.new("RGB", rgba.size, "white")
                try:
                    # Flatten transparency onto white; scale page placement, not image pixels.
                    rgb.paste(rgba, mask=rgba.getchannel("A"))
                    page_width, page_height = landscape(A4) if rgb.width > rgb.height else A4
                    scale = min((page_width - 36) / rgb.width, (page_height - 36) / rgb.height)
                    width, height = rgb.width * scale, rgb.height * scale
                    canvas.setPageSize((page_width, page_height))
                    canvas.drawImage(ImageReader(rgb), (page_width - width) / 2,
                                     (page_height - height) / 2, width=width, height=height)
                    canvas.showPage()
                    count += 1
                finally:
                    rgb.close()
                    rgba.close()
                    oriented.close()
        canvas.save()
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    _validate_exported_pdf(destination)
    return count


def _office_export(app, source: Path, destination: Path, kind: str) -> None:
    """Run on a verified, dedicated Office application inside the worker."""
    document = None
    try:
        if kind == "Excel":
            if source.suffix.casefold() in {".csv", ".tsv"}:
                encoding, codepage = text_encoding(source)
                delimiter = "\t" if source.suffix.casefold() == ".tsv" else ","
                columns, rows = 0, 0
                with source.open("r", encoding=encoding, newline="") as stream:
                    for row in csv.reader(stream, delimiter=delimiter, strict=True):
                        rows += 1
                        columns = max(columns, len(row))
                        if columns > 16384 or rows > 1048576:
                            raise ValueError("CSV/TSV exceeds Excel worksheet limits.")
                if columns == 0:
                    raise ValueError("CSV/TSV is empty or exceeds Excel worksheet limits.")
                # Import identifiers as text so values such as "00123" retain leading zeros.
                app.Workbooks.OpenText(
                    FileName=str(source), Origin=codepage, DataType=1, TextQualifier=1,
                    ConsecutiveDelimiter=False, Tab=delimiter == "\t", Comma=delimiter == ",",
                    Semicolon=False, Space=False, Other=False,
                    FieldInfo=tuple((i, 2) for i in range(1, columns + 1)), Local=False,
                )
                document = app.ActiveWorkbook
                sheet = document.Worksheets(1)
                sheet.PageSetup.Orientation = 2
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
            else:
                document = app.Workbooks.Open(
                    FileName=str(source), UpdateLinks=0, ReadOnly=True,
                    Password=uuid4().hex, WriteResPassword=uuid4().hex,
                    IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False,
                )
            document.ExportAsFixedFormat(Type=0, Filename=str(destination), Quality=0,
                                        IncludeDocProperties=True, IgnorePrintAreas=False,
                                        OpenAfterPublish=False)
        else:
            document = app.Presentations.Open(str(source), -1, 0, 0)
            # ppSaveAsPDF = 32; positional arguments work with late-bound COM.
            document.SaveAs(str(destination), 32)
        if not destination.is_file() or not destination.stat().st_size:
            raise WordExportMissing(f"{kind} did not produce a PDF: {source.name}")
        _validate_exported_pdf(destination)
    finally:
        if document is not None:
            had_error = sys.exc_info()[0] is not None
            try:
                if kind == "Excel":
                    document.Close(SaveChanges=False)
                else:
                    # Close a PDF-exported presentation without saving its source.
                    document.Saved = True
                    document.Close()
            except Exception as exc:
                if had_error:
                    logging.warning("Additional %s close error: %s", kind, exc)
                elif _is_word_disconnection(exc):
                    _validate_exported_pdf(destination)
                    logging.warning("Retaining validated PDF after %s disconnected: %s", kind, source.name)
                else:
                    raise


def _office_worker(source_text: str, destination: str, kind: str,
                   existing_pids: set[int], messages) -> None:
    import pythoncom
    import win32com.client
    import win32process
    pythoncom.CoInitialize()
    log_handler = _WorkerWarningHandler(messages)
    logging.getLogger().addHandler(log_handler)
    app, owned = None, False
    saved = {}
    error = None
    try:
        app = win32com.client.DispatchEx(kind + ".Application")
        hwnd = int(app.Hwnd if kind == "Excel" else app.HWND) & 0xFFFFFFFF
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if not pid or pid in existing_pids:
            raise RuntimeError(f"{kind} reused an existing process. Close {kind} and try again.")
        owned = True
        messages.send(("ready", pid))
        for name, value in (("AutomationSecurity", 3), ("DisplayAlerts", False if kind == "Excel" else 1)):
            saved[name] = getattr(app, name)
            setattr(app, name, value)
        if kind == "Excel":
            for name in ("EnableEvents", "AskToUpdateLinks"):
                saved[name] = getattr(app, name)
                setattr(app, name, False)
        messages.send(("working", Path(source_text).name))
        _office_export(app, Path(source_text), Path(destination), kind)
    except Exception as exc:
        retryable = _is_word_disconnection(exc) or any(
            isinstance(e, WordExportMissing) for e in _exception_chain(exc))
        error = ("error", traceback.format_exc(), retryable)
    finally:
        if owned and app is not None:
            for name, value in saved.items():
                try:
                    setattr(app, name, value)
                except Exception as exc:
                    messages.send(("warning", f"{kind} setting restoration: {exc}"))
            try:
                app.Quit()
            except Exception as exc:
                messages.send(("warning", f"{kind} shutdown: {exc}"))
        app = None
        pythoncom.CoUninitialize()
        logging.getLogger().removeHandler(log_handler)
        log_handler.close()
        messages.send(error or ("success",))
        messages.close()


def _run_office_once(source: Path, destination: Path, kind: str, timeout: float,
                     cancel_event=None) -> None:
    if os.name != "nt":
        raise RuntimeError(f"{kind} conversion requires Windows with desktop {kind} installed.")
    try:
        import win32api
        import win32event
        import win32process
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Install dependencies: py -m pip install pywin32 pypdf Pillow reportlab") from exc
    existing_pids = set(win32process.EnumProcesses())
    context = multiprocessing.get_context("spawn")
    receive, send = context.Pipe(duplex=False)
    process = context.Process(target=_office_worker, args=(
        str(source), str(destination), kind, existing_pids, send))
    handle = None

    def on_message(event):
        nonlocal handle
        if event[0] == "ready":
            if event[1] in existing_pids:
                raise RuntimeError(f"Refusing to control an existing {kind} process.")
            handle = win32api.OpenProcess(0x100001, False, event[1])
        elif event[0] == "warning":
            logging.warning("%s", event[1])
        elif event[0] != "working":
            raise RuntimeError(f"Unexpected {kind} worker event: {event[0]}")

    try:
        _check_cancel(cancel_event)
        process.start()
        send.close()
        _wait_for_word(process, receive, timeout, on_message, cancel_event, kind)
        process.join(timeout=5)
        if process.is_alive() or process.exitcode != 0:
            raise RuntimeError(f"{kind} worker did not exit cleanly.")
        _validate_exported_pdf(destination)
    finally:
        # A held process handle refers to the verified conversion instance only.
        if handle is not None:
            try:
                if win32event.WaitForSingleObject(handle, 0) == 258:
                    win32api.TerminateProcess(handle, 1)
                    win32event.WaitForSingleObject(handle, 5000)
            except Exception as exc:
                logging.warning("%s conversion process cleanup: %s", kind, exc)
            finally:
                handle.Close()
        if process.pid is not None and process.is_alive():
            process.terminate()
            process.join(timeout=5)
            if process.is_alive():
                process.kill()
                process.join(timeout=5)
        receive.close()
        send.close()


def convert_office(source: Path, destination: Path, timeout: float, cancel_event=None) -> None:
    kind = "Excel" if source.suffix.casefold() in EXCEL_FORMATS else "PowerPoint"
    for attempt in range(3):
        _check_cancel(cancel_event)
        destination.unlink(missing_ok=True)
        try:
            _run_office_once(source, destination, kind, timeout, cancel_event)
            return
        except WordWorkerError as exc:
            if not exc.retryable or attempt == 2:
                raise RuntimeError(f"Cannot convert {source.name} with {kind}: {exc}") from exc
            logging.warning("Restarting %s for %s (retry %d/2).", kind, source.name, attempt + 1)


def check_gui_dependencies(files: list[Path]) -> None:
    """Check only the packages and desktop applications needed by this selection."""
    import importlib.util
    modules = {"pypdf": "pypdf"}
    if any(p.suffix.casefold() in IMAGE_FORMATS for p in files):
        modules.update({"PIL": "Pillow", "reportlab": "reportlab"})
    if any(p.suffix.casefold() in WORD_FORMATS | EXCEL_FORMATS | POWERPOINT_FORMATS for p in files):
        if os.name != "nt":
            raise RuntimeError("Office/text conversion needs Windows and the corresponding desktop Office app.")
        modules["win32com"] = "pywin32"
    missing = [package for module, package in modules.items() if importlib.util.find_spec(module) is None]
    if missing:
        raise RuntimeError("Install missing packages first: py -m pip install " + " ".join(missing))
    if os.name == "nt":
        import winreg
        for formats, app in ((WORD_FORMATS, "Word"), (EXCEL_FORMATS, "Excel"),
                             (POWERPOINT_FORMATS, "PowerPoint")):
            if any(p.suffix.casefold() in formats for p in files):
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, app + ".Application\\CLSID"):
                        pass
                except FileNotFoundError as exc:
                    raise RuntimeError(f"Desktop {app} is not registered on this PC. "
                                       "Export those files to PDF first, or install the application.") from exc


def combine_selected(files, run_dir: Path, *, timeout=300, fresh=False,
                     cancel_event=None, output_name="Combined_Files.pdf") -> Path:
    """Build the exact reviewed queue; publish only after checks and cleanup."""
    files = validate_selected(files)
    check_gui_dependencies(files)
    _check_cancel(cancel_event)
    if timeout <= 0:
        raise ValueError("Conversion timeout must be greater than zero.")
    if not run_dir.is_dir():
        raise ValueError("The result directory does not exist.")
    output = run_dir / validate_output_name(output_name)
    pending = output.with_suffix(".pending")
    if output.exists() or pending.exists():
        raise ValueError("This run folder already contains output; choose a new run folder.")
    # Compare source metadata and bytes again after processing, before publishing output.
    before = {}
    for source in files:
        _check_cancel(cancel_event)
        stat = fingerprint(source)
        checksum = _sha256_file(source)
        if fingerprint(source) != stat:
            raise RuntimeError(f"Source changed while checking it: {source.name}")
        before[source] = stat, checksum
    try:
        check_input_pdfs(files, cancel_event)
        with tempfile.TemporaryDirectory(prefix="combine_", dir=run_dir) as temp:
            work = Path(temp)
            word_files = [p for p in files if p.suffix.casefold() in WORD_FORMATS]
            status = {}
            converted = convert_word_files(
                word_files, work, timeout, cache_dir=run_dir.parent / WORD_CACHE_FOLDER,
                reuse_cache=not fresh, conversion_status=status, cancel_event=cancel_event,
            ) if word_files else {}
            for index, source in enumerate(files):
                _check_cancel(cancel_event)
                suffix = source.suffix.casefold()
                if suffix == ".pdf" or source in converted:
                    continue
                destination = work / f"converted_{index:06d}.pdf"
                logging.info("Converting %s: %s", file_kind(source), source.name)
                if suffix in IMAGE_FORMATS:
                    frames = convert_image(source, destination, cancel_event)
                    status[source] = f"Image to A4 PDF; {frames} frame(s)"
                else:
                    convert_office(source, destination, timeout, cancel_event)
                    status[source] = "Converted with " + ("Excel" if suffix in EXCEL_FORMATS else "PowerPoint")
                converted[source] = destination
            report = merge_pdfs([(p, converted.get(p, p)) for p in files], pending, cancel_event)
            for source, row in zip(files, report):
                _check_cancel(cancel_event)
                row["Conversion"] = status.get(source, "Original PDF (reread)")
                row["Source SHA256"] = before[source][1]
                if fingerprint(source) != before[source][0] or _sha256_file(source) != before[source][1]:
                    raise RuntimeError(f"Source changed during processing: {source.name}. Please run again.")
        _check_cancel(cancel_event)
        with (run_dir / "Merge_Report.csv").open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(report[0]))
            writer.writeheader()
            writer.writerows(report)
        _check_cancel(cancel_event)
        # Rename only the completed, page-count-checked PDF to the user's chosen name.
        os.replace(pending, output)
    finally:
        had_error = sys.exc_info()[0] is not None
        try:
            pending.unlink(missing_ok=True)
        except OSError as exc:
            if had_error:
                logging.warning("Additional temporary PDF cleanup error: %s", exc)
            else:
                raise
    logging.info("SUCCESS: %d files; %d pages; %s", len(files), sum(r["Pages"] for r in report), output)
    logging.info("Review page appearance and links. A PDF merge does not preserve cryptographic signature validity.")
    return output


class _QueueLogHandler(logging.Handler):
    def __init__(self, events):
        super().__init__(logging.INFO)
        self.events = events
        self.addFilter(_ConsolePdfWhitespaceFilter())

    def emit(self, record):
        self.events.put(("log", self.format(record)))


def gui_merge_job(files, result_root, timeout, fresh, cancel_event, events,
                  output_name="Combined_Files.pdf") -> None:
    """Run conversion without touching Tk widgets; return messages through the queue."""
    run_dir = None
    completion = None
    root_logger = logging.getLogger()
    old_level = root_logger.level
    handler = _QueueLogHandler(events)
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    file_handler = None
    log_path = None
    root_logger.addHandler(handler)
    root_logger.setLevel(logging.INFO)
    try:
        files = validate_selected(files)
        output_name = validate_output_name(output_name)
        check_gui_dependencies(files)
        _check_cancel(cancel_event)
        run_dir = create_run_directory(result_root)
        log_path = run_dir / "Processing_Log.txt"
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s: %(message)s"))
        root_logger.addHandler(file_handler)
        logging.info("Reviewed selection: %d files in the displayed order.", len(files))
        output = combine_selected(files, run_dir, timeout=timeout, fresh=fresh,
                                  cancel_event=cancel_event, output_name=output_name)
        completion = ("done", str(output))
    except SelectionCancelled:
        logging.info("Cancelled. No new combined PDF was published.")
        completion = ("cancelled", str(run_dir or ""))
    except Exception as exc:
        details = traceback.format_exc()
        logging.exception("STOPPED. No new combined PDF was published.")
        # Report a log path only after it was opened, or a fallback diagnostic was saved.
        actual_log = str(log_path) if file_handler is not None else write_startup_diagnostic(details)
        actual_run = str(run_dir) if run_dir is not None and run_dir.is_dir() else ""
        completion = ("error", str(exc), actual_run, actual_log)
    finally:
        root_logger.removeHandler(handler)
        handler.close()
        if file_handler is not None:
            root_logger.removeHandler(file_handler)
            file_handler.close()
        root_logger.setLevel(old_level)
    events.put(completion)


class CombinerWindow:
    """Manage the reviewed file queue and apply worker events on Tk's main thread."""

    def __init__(self, root):
        import queue
        import threading
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        self.tk, self.ttk = tk, ttk
        self.dialog, self.message = filedialog, messagebox
        self.root = root
        self.model = FileSelection()
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.running = False
        self.close_when_stopped = False
        self.output = None
        self.run_folder = None
        self.controls = []
        root.title(f"PDF Tweaks | {VERSION}")
        window_width = min(1080, max(920, root.winfo_screenwidth() - 80))
        window_height = min(800, max(680, root.winfo_screenheight() - 100))
        root.geometry(f"{window_width}x{window_height}")
        root.minsize(920, 680)
        root.configure(bg="#f3f5f8")
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#f3f5f8")
        style.configure("TLabel", background="#f3f5f8", foreground="#243247", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=(11, 7))
        style.configure("Accent.TButton", background="#185caa", foreground="white")
        style.map("Accent.TButton", background=[("active", "#124a8b"), ("disabled", "#bbc6d2")])
        style.configure("Treeview", rowheight=29, font=("Segoe UI", 10), background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=(7, 7))
        style.configure("TCheckbutton", background="#f3f5f8", font=("Segoe UI", 10))
        style.configure("Link.TButton", font=("Segoe UI", 9, "underline"), padding=(5, 4),
                        foreground="#185caa", background="#f3f5f8", borderwidth=0, relief="flat")
        style.map("Link.TButton", foreground=[("active", "#124a8b")],
                  background=[("pressed", "#dce6f2"), ("active", "#e7edf5")])

        header = tk.Frame(root, bg="#14283f", padx=24, pady=18)
        header.pack(fill="x")
        tk.Label(header, text="Combine your files into one PDF", bg="#14283f", fg="white",
                 font=("Segoe UI", 21, "bold")).pack(anchor="w")
        tk.Label(header, text="Choose files  /  Review the order  /  Create PDF", bg="#14283f",
                 fg="#c5d7eb", font=("Segoe UI", 11)).pack(anchor="w", pady=(6, 0))
        body = ttk.Frame(root, padding=(22, 15))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)
        self.status = tk.StringVar(value="Add the files you want to include. Nothing is combined automatically.")
        self.count = tk.StringVar(value="0 files selected")
        self.folder = tk.StringVar(value=str(default_result_root()))
        self.output_name = tk.StringVar(value="Combined_Files.pdf")
        self.sort_mode = tk.StringVar(value="Filename (A-Z)")
        self.timeout = tk.StringVar(value="300")
        self.fresh = tk.BooleanVar(value=False)

        bar = ttk.Frame(body)
        bar.grid(row=0, column=0, sticky="ew")
        self.button(bar, "Add files...", self.add_files, accent=True).pack(side="left")
        self.button(bar, "Replace selected...", self.replace_file).pack(side="left", padx=(7, 0))
        self.button(bar, "Remove", self.remove).pack(side="left", padx=(7, 0))
        self.button(bar, "Clear", self.clear).pack(side="left", padx=(7, 0))
        ttk.Label(bar, textvariable=self.count).pack(side="right")
        ttk.Label(body, text="PDF | Word / TXT | Excel / CSV / TSV | PowerPoint | JPG / PNG / TIFF / BMP / GIF / WebP",
                  wraplength=880).grid(row=1, column=0, sticky="w", pady=(12, 6))
        table = ttk.Frame(body)
        table.grid(row=2, column=0, sticky="nsew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("order", "name", "type", "folder"),
                                 show="headings", selectmode="extended", height=6)
        for column, label, width in (("order", "#", 45), ("name", "File", 370),
                                     ("type", "Type", 120), ("folder", "Folder", 410)):
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, minwidth=45, anchor="w", stretch=column in {"name", "folder"})
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.bind("<Delete>", lambda event: self.remove())

        order_bar = ttk.Frame(body)
        order_bar.grid(row=3, column=0, sticky="ew", pady=(8, 12))
        self.button(order_bar, "Move up", lambda: self.move(-1)).pack(side="left")
        self.button(order_bar, "Move down", lambda: self.move(1)).pack(side="left", padx=6)
        sort_box = ttk.Combobox(order_bar, textvariable=self.sort_mode, width=22, state="readonly",
                                values=("Filename (A-Z)", "Natural (2 before 10)"))
        sort_box.pack(side="left", padx=(0, 6))
        self.controls.append(sort_box)
        self.button(order_bar, "Sort", self.sort_files).pack(side="left")
        ttk.Label(order_bar, text="Displayed order = merge order.").pack(side="right")

        filename_bar = ttk.Frame(body)
        filename_bar.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(filename_bar, text="PDF filename:").pack(side="left")
        self.entry(filename_bar, self.output_name, 36).pack(side="left", padx=(8, 0))
        self.button(filename_bar, "Supported formats / notes", self.show_help).pack(side="right")

        destination_bar = ttk.Frame(body)
        destination_bar.grid(row=5, column=0, sticky="ew")
        ttk.Label(destination_bar, text="Results folder:").pack(side="left")
        self.entry(destination_bar, self.folder).pack(side="left", fill="x", expand=True, padx=8)
        self.button(destination_bar, "Browse...", self.choose_output).pack(side="left")
        ttk.Label(body, text="Each run creates a new dated folder with the PDF, report and processing log.").grid(
            row=6, column=0, sticky="w", pady=(5, 8))
        options = ttk.Frame(body)
        options.grid(row=7, column=0, sticky="ew")
        check = ttk.Checkbutton(options, text="Reconvert Word files (ignore saved conversions)", variable=self.fresh)
        check.pack(side="left")
        self.controls.append(check)
        self.entry(options, self.timeout, 6).pack(side="right")
        ttk.Label(options, text="Office timeout, seconds: ").pack(side="right")

        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.grid(row=8, column=0, sticky="ew", pady=(12, 7))
        ttk.Label(body, textvariable=self.status, wraplength=870).grid(row=9, column=0, sticky="w")
        log_frame = ttk.Frame(body)
        log_frame.grid(row=10, column=0, sticky="ew", pady=(8, 12))
        self.log = tk.Text(log_frame, height=4, wrap="word", font=("Consolas", 9),
                           state="disabled", bg="white", fg="#34445a", relief="solid", borderwidth=1)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_scroll.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=log_scroll.set)
        footer = ttk.Frame(body)
        footer.grid(row=11, column=0, sticky="ew")
        self.button(footer, "Create combined PDF", self.start, accent=True).pack(side="left")
        self.cancel_button = ttk.Button(footer, text="Cancel run", command=self.cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=8)
        # Draw the external-link icon with Tk; retain it on self so Tk can keep displaying it.
        self.github_icon = tk.PhotoImage(master=root, width=10, height=10)
        for rectangle in ((2, 5, 4, 14), (2, 12, 11, 14), (9, 9, 11, 14), (2, 5, 7, 7),
                          (8, 2, 14, 4), (12, 2, 14, 8), (6, 8, 8, 10), (8, 6, 10, 8),
                          (10, 4, 12, 6)):
            self.github_icon.put("#185caa", to=rectangle)
        self.github_button = ttk.Button(footer, text="Reports Bugs", style="Link.TButton",
                                        image=self.github_icon, compound="right", command=self.open_github,
                                        cursor="hand2", takefocus=True)
        self.github_button.pack(side="left", padx=(4, 0))
        self.github_button.bind("<Return>", lambda event: self.open_github())
        self.open_button = ttk.Button(footer, text="Open PDF", command=self.open_pdf, state="disabled")
        self.open_button.pack(side="right")
        self.folder_button = ttk.Button(footer, text="Open results", command=self.open_folder, state="disabled")
        self.folder_button.pack(side="right", padx=8)
        self._poll_id = root.after(100, self.poll)

    @staticmethod
    def default_folder() -> Path:
        documents = Path.home() / "Documents"
        return documents if documents.is_dir() else Path.home()

    def button(self, parent, text, command, accent=False):
        button = self.ttk.Button(parent, text=text, command=command,
                                  style="Accent.TButton" if accent else "TButton")
        self.controls.append(button)
        return button

    def entry(self, parent, variable, width=None):
        entry = self.ttk.Entry(parent, textvariable=variable, **({"width": width} if width else {}))
        self.controls.append(entry)
        return entry

    def selected_indices(self):
        return sorted(int(item) for item in self.tree.selection())

    def redraw(self, selected=()):
        self.tree.delete(*self.tree.get_children())
        for i, path in enumerate(self.model.files):
            self.tree.insert("", "end", iid=str(i), values=(i + 1, path.name, file_kind(path), str(path.parent)))
        for i in selected:
            if i < len(self.model.files):
                self.tree.selection_add(str(i))
        self.count.set(f"{len(self.model.files)} files selected")

    def pick_files(self, single=False):
        initial = self.model.files[-1].parent if self.model.files else self.default_folder()
        types = [("Supported documents and images", " ".join("*" + x for x in sorted(GUI_FORMATS))),
                 ("PDF", "*.pdf"), ("All files (unsupported types are reported)", "*")]
        options = dict(parent=self.root, title="Choose files to include" if not single else "Choose replacement file",
                       initialdir=str(initial), filetypes=types)
        return self.dialog.askopenfilename(**options) if single else self.dialog.askopenfilenames(**options)

    def add_files(self):
        if self.running:
            return
        chosen = self.pick_files()
        if not chosen:
            return
        try:
            paths = validate_selected(chosen)
        except Exception as exc:
            self.message.showerror("Files not added", str(exc), parent=self.root)
            return
        mode = "natural" if self.sort_mode.get().startswith("Natural") else "filename"
        for source in ordered_files(paths, mode):
            if source in self.model.files:
                continue
            matches = self.model.same_document(source)
            replace = False
            if matches:
                names = ", ".join(self.model.files[i].name for i in matches)
                answer = self.message.askyesnocancel(
                    "Review another format of this document",
                    f"Already selected: {names}\nNew file: {source.name}\n\n"
                    "Yes = replace the earlier item in this list.\nNo = include both formats.\n"
                    "Cancel = skip this new file.\n\nSource files are kept.", parent=self.root)
                if answer is None:
                    continue
                replace = answer
            self.model.add(source, replace=replace)
        self.redraw()
        self.status.set("Review the list, adjust the order, then click Create combined PDF.")

    def replace_file(self):
        if self.running:
            return
        indices = self.selected_indices()
        if len(indices) != 1:
            self.message.showinfo("Select one item", "Select exactly one row to replace in the merge list.", parent=self.root)
            return
        chosen = self.pick_files(single=True)
        if not chosen:
            return
        try:
            self.model.replace(indices[0], Path(chosen))
            self.redraw(indices)
            self.status.set("Selected item replaced in the list. The source files were kept.")
        except Exception as exc:
            self.message.showerror("Cannot replace item", str(exc), parent=self.root)

    def remove(self):
        if not self.running:
            for index in reversed(self.selected_indices()):
                del self.model.files[index]
            self.redraw()

    def clear(self):
        if not self.running:
            self.model.files.clear()
            self.redraw()

    def move(self, direction):
        if not self.running:
            self.redraw(self.model.move(self.selected_indices(), direction))

    def sort_files(self):
        if not self.running:
            mode = "natural" if self.sort_mode.get().startswith("Natural") else "filename"
            self.model.files = ordered_files(self.model.files, mode)
            self.redraw()


    def choose_output(self):
        chosen = self.dialog.askdirectory(parent=self.root, title="Choose where result folders are created",
                                          initialdir=str(self.default_folder()), mustexist=True)
        if chosen:
            self.folder.set(chosen)

    def show_help(self):
        self.message.showinfo("Supported formats and conversion notes",
            "PDF: .pdf\nWord: .doc, .docx, .docm, .rtf, .odt, .txt\n"
            "Excel: .xls, .xlsx, .csv, .tsv (.xlsm, .xlsb, are not supported yet, .ods is not tested yet)\n"
            "PowerPoint: .ppt, .pptx, .pptm, .pps, .ppsx, .odp\n"
            "Images: .jpg, .jpeg, .png, .tif, .tiff, .bmp, .gif, .webp\n\n"
            "Windows desktop Word, Excel or PowerPoint is required for its respective formats. "
            "Use UTF-8 or BOM-marked UTF-16 for TXT/CSV/TSV. CSV uses commas; TSV uses tabs. "
            "Data imports use text columns to retain leading zeros.\n\n"
            "Excel uses visible sheets and existing print areas; check Print Preview first. "
            "CSV/TSV is fitted to one page wide. Images use A4 pages; all image frames are included. "
            "Word tracked-change markup is excluded.\n\n"
            "Unsupported types (such as ZIP, EXE, audio and video) need a separate representation "
            "or export before they can become document pages. Encrypted inputs require an "
            "unencrypted working copy. Merging does not retain digital-signature validity. "
            "Review the final page appearance, links and report.", parent=self.root)

    def start(self):
        if self.running:
            return
        try:
            files = validate_selected(self.model.files)
            check_gui_dependencies(files)
            timeout = int(self.timeout.get())
            if timeout <= 0:
                raise ValueError("Timeout must be a positive number of seconds.")
            if not self.folder.get().strip():
                raise ValueError("Choose a results folder.")
            destination = Path(self.folder.get()).expanduser().resolve()
            output_name = validate_output_name(self.output_name.get())
            self.output_name.set(output_name)
        except Exception as exc:
            self.message.showerror("Check your selection", str(exc), parent=self.root)
            return
        if not self.message.askyesno("Create combined PDF?",
            f"Combine exactly {len(files)} files in the displayed order?\n\n"
            f"First: {files[0].name}\nLast: {files[-1].name}\n\n"
            f"PDF filename: {output_name}\n\nA new result folder will be created inside:\n{destination}",
            parent=self.root):
            return
        self.running = True
        self.output, self.run_folder = None, None
        self.cancel_event.clear()
        for control in self.controls:
            control.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.status.set("Checking and converting the selected files...")
        self.progress.start(12)
        import threading
        # The worker handles files and Office calls; poll() alone updates Tk widgets.
        threading.Thread(target=gui_merge_job, args=(files, destination, timeout, self.fresh.get(),
                         self.cancel_event, self.events, output_name), daemon=False).start()

    def cancel(self):
        if self.running:
            self.cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status.set("Cancellation requested. Waiting for the current PDF operation to stop...")

    def poll(self):
        """Drain worker messages on the main thread and restore controls after completion."""
        import queue
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    self.log.configure(state="normal")
                    self.log.insert("end", event[1] + "\n")
                    if int(self.log.index("end-1c").split(".")[0]) > 2000:
                        self.log.delete("1.0", "200.0")
                    self.log.see("end")
                    self.log.configure(state="disabled")
                    if not self.cancel_event.is_set():
                        self.status.set(event[1].split("\n")[0][:180])
                    continue
                self.running = False
                self.progress.stop()
                self.cancel_button.configure(state="disabled")
                for control in self.controls:
                    control.configure(state="readonly" if isinstance(control, self.ttk.Combobox) else "normal")
                if event[0] == "done":
                    self.output = Path(event[1])
                    self.run_folder = self.output.parent
                    self.open_button.configure(state="normal")
                    self.status.set(f"PDF created: {self.output.name}. Review it using Open PDF and Open results.")
                    if not self.close_when_stopped:
                        self.message.showinfo("Combined PDF created", f"Saved:\n{self.output}\n\n"
                            "The same folder contains Merge_Report.csv and Processing_Log.txt. "
                            "Review page appearance and any warnings.", parent=self.root)
                else:
                    result_folder = event[2] if event[0] == "error" else event[1]
                    self.run_folder = Path(result_folder) if result_folder else None
                    if event[0] == "cancelled":
                        self.status.set("Cancelled. No new combined PDF was published.")
                    else:
                        self.status.set("Processing stopped. No new combined PDF was published.")
                        if not self.close_when_stopped:
                            detail = event[1][:1400]
                            if len(event) > 3 and event[3]:
                                detail += f"\n\nDiagnostic log: {event[3]}"
                            else:
                                detail += "\n\nNo diagnostic file could be saved. Details are shown in the window log."
                            self.message.showerror("Could not create PDF", detail, parent=self.root)
                if self.run_folder:
                    self.folder_button.configure(state="normal")
                if self.close_when_stopped:
                    self.destroy()
                    return
        except queue.Empty:
            self._poll_id = self.root.after(100, self.poll)

    def _open_path(self, path):
        try:
            if os.name == "nt":
                os.startfile(str(path))
            else:
                import subprocess
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
        except Exception as exc:
            self.message.showerror("Cannot open result", str(exc), parent=self.root)

    def open_github(self):
        """Open the public profile only on request, without sending document paths."""
        import webbrowser
        try:
            if not webbrowser.open_new_tab(GITHUB_PROFILE_URL):
                self.message.showerror("Could not open GitHub",
                    f"Open this address in your browser:\n{GITHUB_PROFILE_URL}", parent=self.root)
        except Exception as exc:
            self.message.showerror("Could not open GitHub",
                f"{exc}\n\nOpen this address in your browser:\n{GITHUB_PROFILE_URL}", parent=self.root)

    def open_pdf(self):
        if self.output:
            self._open_path(self.output)

    def open_folder(self):
        if self.run_folder:
            self._open_path(self.run_folder)

    def close(self):
        if self.running:
            if self.message.askyesno("Cancel and close?", "Cancel the current run and close after it stops?",
                                      parent=self.root):
                self.close_when_stopped = True
                self.cancel()
        else:
            self.destroy()

    def destroy(self):
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        self.root.destroy()


def launch_gui() -> int:
    try:
        import tkinter as tk
        root = tk.Tk()
    except Exception as exc:
        print(f"Cannot open the graphical interface: {exc}\n"
              "Use a Python installation with Tcl/Tk, or use --files / --folder from the command line.",
              file=sys.stderr)
        return 1
    CombinerWindow(root)
    root.mainloop()
    return 0


def main() -> int:
    if not sys.argv[1:] or sys.argv[1:] == ["--gui"]:
        return launch_gui()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    inputs = parser.add_mutually_exclusive_group()
    inputs.add_argument("--files", nargs="+", metavar="FILE", help="Exact input files, in the given order")
    inputs.add_argument("--folder", metavar="FOLDER", help="Include supported files from this folder")
    parser.add_argument("--gui", action="store_true", help="Open the graphical interface without other options")
    parser.add_argument("--recursive", action="store_true", help="Include subfolders with --folder")
    parser.add_argument("--list", action="store_true", help="Preview the selected file order without processing")
    parser.add_argument("--sort", choices=("filename", "natural"), help="Optional sorting of the chosen files")
    parser.add_argument("--output-dir", metavar="FOLDER", help="Parent directory for new dated result folders")
    parser.add_argument("--output-name", default="Combined_Files.pdf", metavar="NAME", help="PDF filename")
    parser.add_argument("--fresh", action="store_true", help="Reconvert Word files instead of reusing saved PDFs")
    parser.add_argument("--timeout", type=int, default=300, help="Seconds allowed for each Office conversion stage")
    parser.add_argument("--yes", action="store_true", help="Process the reviewed selection without a confirmation prompt")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args()
    if args.gui:
        parser.error("Use --gui without other options; make the selection in the window.")
    if not args.files and args.folder is None:
        parser.error("Choose inputs with --files or --folder, or run without options to open the GUI.")
    if args.recursive and args.folder is None:
        parser.error("--recursive requires --folder.")
    if args.timeout <= 0:
        parser.error("--timeout must be a positive number of seconds.")
    try:
        files = validate_selected(args.files) if args.files else discover(Path(args.folder).expanduser().resolve(), args.recursive)
        if args.sort:
            files = ordered_files(files, args.sort)
        output_name = validate_output_name(args.output_name)
        result_root = Path(args.output_dir).expanduser().resolve() if args.output_dir else default_result_root()
        print(f"Selected files: {len(files)}", flush=True)
        for index, source in enumerate(files, 1):
            print(f"{index:04d}. {source}", flush=True)
        if args.list:
            return 0
        print(f"Output: {output_name}; new dated folder under {result_root}", flush=True)
        if not args.yes:
            try:
                answer = input("Combine exactly these files? Type YES to continue: ").strip()
            except EOFError as exc:
                raise ValueError("No confirmation input is available. Use --list to review, then --yes to process.") from exc
            if answer.casefold() != "yes":
                print("Cancelled. No conversion or merge was started.")
                return 0
        import queue
        import threading
        events = queue.Queue()
        console = logging.StreamHandler(sys.stdout)
        console.addFilter(_ConsolePdfWhitespaceFilter())
        console.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        root_logger = logging.getLogger()
        root_logger.addHandler(console)
        try:
            gui_merge_job(files, result_root, args.timeout, args.fresh, threading.Event(), events, output_name)
        finally:
            root_logger.removeHandler(console)
            console.close()
        completion = None
        while not events.empty():
            event = events.get_nowait()
            if event[0] != "log":
                completion = event
        if completion is not None and completion[0] == "done":
            return 0
        if completion is not None and completion[0] == "error":
            print(f"ERROR: {completion[1]}", file=sys.stderr)
            if len(completion) > 3 and completion[3]:
                print(f"Diagnostic log: {completion[3]}", file=sys.stderr)
            return 1
        print("Cancelled. No new combined PDF was published.")
        return 0
    except KeyboardInterrupt:
        print("\nCancelled by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    # Keep startup behind this guard: Windows spawns Office workers by importing the module.
    multiprocessing.freeze_support()
    raise SystemExit(main())
