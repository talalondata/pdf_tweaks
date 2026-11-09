# PDF Tweaks — Release Notes

## Version 1.1.0

A general-purpose desktop application for combining selected documents and images into one PDF, with a graphical interface and command-line mode.

### Highlights

- General-purpose file selection, replacement and manual ordering.
- Custom output PDF filename and a choice of alphabetical or natural filename sorting.
- Separate dated result folders containing the combined PDF, page-range report and processing log.
- Clearer output-folder errors, with diagnostic-log paths shown when a log is saved.
- A **Report Bugs ↗** footer link to the author's profile, available by mouse or keyboard.
- Consistent `pdf_tweaks.py` commands and clearer contribution guidance.
- Useful code comments covering Office worker ownership, cache validation, PDF recovery and GUI event handling.

### Getting started on Windows

Install Python 3.11 or later with Tcl/Tk support, then open PowerShell in the project folder. Create an isolated environment, install dependencies and launch the application:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe pdf_tweaks.py
```

Office inputs require Windows and the corresponding installed desktop application. PDF/image processing uses Python packages directly. See [README.md](README.md) for supported formats and command-line examples.

### Validation and limitations

The automated suite covers real PDF/image operations and Tkinter interface behavior. Office export paths use simulated applications; actual Word, Excel and PowerPoint conversion requires a separate check on the target Windows PC.

Review page appearance, links, the page-range report and any warnings before sharing a combined PDF. Encrypted PDFs need an unencrypted working copy, and merging does not retain the cryptographic validity of source digital signatures.

### Feedback

Suggestions, bug reports and contributions are welcome through the [GitHub repository](https://github.com/talalondata/pdf_tweaks).

Created with care by [Talal Faiz](https://github.com/talalondata). Licensed under [MIT](LICENSE).
