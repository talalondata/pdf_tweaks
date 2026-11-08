import importlib.util
import json
import logging
import os
from pathlib import Path
import queue
import subprocess
import sys
import tempfile
import threading
import unittest
from unittest.mock import MagicMock, patch

SCRIPT = Path(__file__).resolve().parents[1] / 'pdf_tweaks.py'
spec = importlib.util.spec_from_file_location('combiner', SCRIPT)
c = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = c
spec.loader.exec_module(c)

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas


def pdf(path, labels=('page',)):
    out = Canvas(str(path))
    for label in labels:
        out.drawString(50, 780, label)
        out.showPage()
    out.save()
    return path


class Tests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name)
        self.a = pdf(self.base / '02 - Report.pdf', ('Alpha',))
        self.b = pdf(self.base / '10 - Appendix.pdf', ('Beta', 'Gamma'))
        self.run = self.base / 'results' / 'run1'
        self.run.mkdir(parents=True)

    def tearDown(self):
        self.temp.cleanup()

    def file(self, name, content=b'example'):
        p = self.base / name
        p.write_bytes(content)
        return p

    def test_selected_order_is_exact(self):
        self.assertEqual(c.validate_selected([self.b, self.a]), [self.b, self.a])

    def test_missing_rejected(self):
        with self.assertRaises(FileNotFoundError): c.validate_selected([self.base / 'gone.pdf'])

    def test_unsupported_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Unsupported'): c.validate_selected([self.file('archive.zip')])

    def test_lock_rejected(self):
        with self.assertRaisesRegex(ValueError, 'lock'): c.validate_selected([self.file('~$Test.docx')])

    def test_empty_rejected(self):
        with self.assertRaisesRegex(ValueError, 'empty'): c.validate_selected([self.file('empty.pdf', b'')])

    def test_directory_rejected(self):
        with self.assertRaisesRegex(ValueError, 'regular file'): c.validate_selected([self.base])

    def test_duplicate_rejected(self):
        with self.assertRaisesRegex(ValueError, 'already'): c.validate_selected([self.a, self.a])

    def test_no_selection_rejected(self):
        with self.assertRaisesRegex(ValueError, 'at least'): c.validate_selected([])

    def test_uppercase_extension(self):
        self.assertEqual(c.file_kind(self.file('CAPS.PNG')), 'Image')

    def test_same_stem_replacement_keeps_position_and_sources(self):
        word = self.file('02 - Report.docx')
        model = c.FileSelection()
        model.add(word); model.add(self.b); model.add(self.a, replace=True)
        self.assertEqual(model.files, [self.a, self.b])
        self.assertTrue(word.exists())

    def test_same_stem_both_allowed(self):
        word = self.file('02 - Report.docx')
        model = c.FileSelection(); model.add(word); model.add(self.a)
        self.assertEqual(model.files, [word, self.a])

    def test_duplicate_add_is_ignored(self):
        model = c.FileSelection(); model.add(self.a); model.add(self.a)
        self.assertEqual(model.files, [self.a])

    def test_explicit_replace_different_name(self):
        model = c.FileSelection(); model.add(self.a); model.replace(0, self.b)
        self.assertEqual(model.files, [self.b]); self.assertTrue(self.a.exists())

    def test_explicit_replace_duplicate_rejected(self):
        model = c.FileSelection(); model.add(self.a); model.add(self.b)
        with self.assertRaisesRegex(ValueError, 'already'): model.replace(0, self.b)

    def test_multiselection_move_keeps_relative_order(self):
        model = c.FileSelection(); model.files = [self.a, self.b, self.file('321.pdf'), self.file('322.pdf')]
        indices = model.move([1, 2], -1)
        self.assertEqual(indices, [0, 1]); self.assertEqual(model.files[-2:], [self.a, self.base / '322.pdf'])
        self.assertEqual(model.move(indices, 1), [1, 2])
        self.assertEqual(model.files[0], self.a)

    def test_move_boundary(self):
        model = c.FileSelection(); model.files = [self.a, self.b]
        self.assertEqual(model.move([0], -1), [0]); self.assertEqual(model.move([1], 1), [1])




    def test_utf8_text(self):
        p = self.file('Urdu.txt', 'السلام علیکم'.encode())
        self.assertEqual(c.text_encoding(p), ('utf-8-sig', 65001))

    def test_utf16_text(self):
        p = self.file('text.txt', 'sample'.encode('utf-16'))
        self.assertEqual(c.text_encoding(p), ('utf-16', 1200))

    def test_invalid_encoding_not_guessed(self):
        with self.assertRaisesRegex(ValueError, 'UTF-8'): c.text_encoding(self.file('latin.txt', b'\xe9'))

    def test_utf32_not_misread_as_utf16(self):
        with self.assertRaises(ValueError): c.text_encoding(self.file('text.txt', 'a'.encode('utf-32')))

    def test_binary_not_treated_as_text(self):
        with self.assertRaisesRegex(ValueError, 'binary'): c.text_encoding(self.file('binary.txt', b'a\0b'))

    def test_pdf_merge_order_and_page_ranges(self):
        output = c.combine_selected([self.b, self.a], self.run)
        reader = PdfReader(output)
        self.assertEqual([p.extract_text().strip() for p in reader.pages], ['Beta', 'Gamma', 'Alpha'])
        import csv
        with (self.run / 'Merge_Report.csv').open(encoding='utf-8-sig') as f:
            report = list(csv.DictReader(f))
        self.assertEqual((report[0]['First page'], report[0]['Last page']), ('1', '2'))
        self.assertEqual(report[1]['First page'], '3')
        self.assertEqual(report[0]['Source SHA256'], c._sha256_file(self.b))
        self.assertEqual(len(reader.outline), 2)

    def test_pdf_selection_does_not_auto_include_other_files(self):
        output = c.combine_selected([self.a], self.run)
        self.assertEqual(len(PdfReader(output).pages), 1)

    def test_unselected_folder_changes_do_not_abort(self):
        real_merge = c.merge_pdfs
        def merger(*args):
            pdf(self.base / '299 - added.pdf')
            return real_merge(*args)
        with patch.object(c, 'merge_pdfs', side_effect=merger):
            output = c.combine_selected([self.a], self.run)
        self.assertTrue(output.is_file())

    def test_existing_output_is_kept(self):
        output = c.combine_selected([self.a], self.run)
        original = output.read_bytes()
        with self.assertRaisesRegex(ValueError, 'already'): c.combine_selected([self.b], self.run)
        self.assertEqual(output.read_bytes(), original)

    def test_source_change_aborts_publication(self):
        real_merge = c.merge_pdfs
        def merger(*args):
            report = real_merge(*args)
            with self.a.open('ab') as f: f.write(b'\n% modified')
            return report
        with patch.object(c, 'merge_pdfs', side_effect=merger):
            with self.assertRaisesRegex(RuntimeError, 'changed'): c.combine_selected([self.a], self.run)
        self.assertFalse((self.run / 'Combined_Files.pdf').exists())
        self.assertFalse((self.run / 'Combined_Files.pending').exists())

    def test_encrypted_input_stops(self):
        encrypted = self.base / 'encrypted.pdf'
        w = PdfWriter(); w.add_blank_page(100,100); w.encrypt('password'); w.write(encrypted); w.close()
        with self.assertRaisesRegex(RuntimeError, 'Encrypted'): c.combine_selected([encrypted], self.run)
        self.assertFalse((self.run / 'Combined_Files.pdf').exists())

    def test_invalid_pdf_stops(self):
        invalid = self.file('invalid.pdf')
        with self.assertRaises(RuntimeError): c.combine_selected([invalid], self.run)
        self.assertFalse((self.run / 'Combined_Files.pdf').exists())

    def test_duplicate_info_recovery_is_retained(self):
        import re
        data=self.a.read_bytes()
        match=re.search(rb'/Info\s+\d+\s+0\s+R',data)
        self.assertIsNotNone(match)
        self.a.write_bytes(data[:match.end()]+b' '+match.group()+data[match.end():])
        output=c.combine_selected([self.a,self.b],self.run)
        self.assertEqual(len(PdfReader(output).pages),3)
        self.assertIn('Duplicate /Info recovery',(self.run/'Merge_Report.csv').read_text())

    def test_same_size_same_timestamp_source_change_detected(self):
        stat=self.a.stat();real_merge=c.merge_pdfs
        def merger(*args):
            report=real_merge(*args)
            data=self.a.read_bytes();self.a.write_bytes(data.replace(b'PDF',b'pdf',1))
            os.utime(self.a,ns=(stat.st_atime_ns,stat.st_mtime_ns))
            return report
        with patch.object(c,'merge_pdfs',side_effect=merger):
            with self.assertRaisesRegex(RuntimeError,'changed'):c.combine_selected([self.a],self.run)
        self.assertFalse((self.run/'Combined_Files.pdf').exists())

    def test_cleanup_error_keeps_original_failure(self):
        original=Path.unlink
        def unlink(path,*args,**kwargs):
            if path.name=='Combined_Files.pending':raise PermissionError('cleanup failed')
            return original(path,*args,**kwargs)
        with patch.object(c,'merge_pdfs',side_effect=RuntimeError('original merge failure')):
            with patch.object(Path,'unlink',unlink):
                with self.assertRaisesRegex(RuntimeError,'original merge failure'):
                    c.combine_selected([self.a],self.run)

    def test_pre_cancelled_does_no_merge(self):
        event = threading.Event(); event.set()
        with self.assertRaises(c.SelectionCancelled): c.combine_selected([self.a], self.run, cancel_event=event)
        self.assertEqual(list(self.run.iterdir()), [])

    def test_cancel_during_merge_no_final(self):
        event = threading.Event()
        real_merge = c.merge_pdfs
        def merger(*args):
            report = real_merge(*args)
            event.set()
            return report
        with patch.object(c, 'merge_pdfs', side_effect=merger):
            with self.assertRaises(c.SelectionCancelled): c.combine_selected([self.a], self.run, cancel_event=event)
        self.assertFalse((self.run / 'Combined_Files.pdf').exists())
        self.assertFalse((self.run / 'Combined_Files.pending').exists())

    def test_transparent_png_mixed_with_pdf(self):
        image = self.base / '300.png'
        Image.new('RGBA', (100,200), (255,0,0,128)).save(image)
        output = c.combine_selected([self.a, image, self.b], self.run)
        reader = PdfReader(output)
        self.assertEqual(len(reader.pages), 4)
        xobj = next(iter(reader.pages[1]['/Resources']['/XObject'].values())).get_object()
        self.assertEqual(xobj['/Width'], 100)
        self.assertEqual(xobj.get_data()[:3], bytes([255,127,127]))
        self.assertEqual(reader.pages[2].extract_text().strip(), 'Beta')

    def test_multiframe_tiff_keeps_all_frames(self):
        image = self.base / 'scan.tiff'
        a = Image.new('RGB',(30,50),'red'); b = Image.new('RGB',(30,50),'blue')
        a.save(image, save_all=True, append_images=[b]); a.close(); b.close()
        output = self.run / 'tiff.pdf'
        self.assertEqual(c.convert_image(image, output), 2)
        reader = PdfReader(output)
        colors = [next(iter(p['/Resources']['/XObject'].values())).get_object().get_data()[:3] for p in reader.pages]
        self.assertEqual(colors, [bytes([255,0,0]), bytes([0,0,255])])

    def test_animated_gif_keeps_all_frames(self):
        image = self.base / 'animation.gif'
        a = Image.new('RGB',(40,30),'red'); b = Image.new('RGB',(40,30),'green')
        a.save(image, save_all=True, append_images=[b], duration=100, loop=0); a.close(); b.close()
        self.assertEqual(c.convert_image(image, self.run / 'gif.pdf'), 2)

    def test_image_exif_orientation(self):
        image = self.base / 'rotated.jpg'
        a = Image.new('RGB',(80,40),'red'); exif = a.getexif(); exif[274] = 6; a.save(image, exif=exif); a.close()
        output = self.run / 'rotated.pdf'; c.convert_image(image, output)
        p = PdfReader(output).pages[0]
        obj = next(iter(p['/Resources']['/XObject'].values())).get_object()
        self.assertEqual((obj['/Width'], obj['/Height']), (40,80))
        self.assertLess(float(p.mediabox.width), float(p.mediabox.height))

    def test_bad_image_has_no_published_pdf(self):
        image = self.file('bad.png')
        with self.assertRaises(Exception): c.combine_selected([image], self.run)
        self.assertFalse((self.run / 'Combined_Files.pdf').exists())

    def test_image_cancelled_before_frame(self):
        image = self.base / 'x.png'; Image.new('RGB',(4,4)).save(image)
        event = threading.Event(); event.set()
        output = self.run / 'image.pdf'
        with self.assertRaises(c.SelectionCancelled): c.convert_image(image, output, event)
        self.assertFalse(output.exists())

    def test_pdf_form_names_are_separate(self):
        forms=[]
        for i in range(2):
            p=self.base / f'form{i}.pdf'; out=Canvas(str(p)); out.acroForm.textfield(name='Name', value=f'value{i}',x=40,y=700,width=200,height=20)
            out.showPage();out.save();forms.append(p)
        output = c.combine_selected(forms, self.run)
        fields=PdfReader(output).get_fields()
        self.assertEqual(fields['document_0001.Name']['/V'], 'value0')
        self.assertEqual(fields['document_0002.Name']['/V'], 'value1')

    def test_word_cache_reuses_valid_conversions(self):
        source=self.file('word.docx'); cache=self.base / 'cache'
        def batch(sources,destination,timeout,converted,total,cancel_event=None):
            for i,p in enumerate(sources): converted[p]=pdf(destination / f'{i}.pdf', (p.read_text(),))
        work=self.base / 'work1';work.mkdir()
        with patch.object(c,'_run_word_batch', side_effect=batch) as mock:
            first=c.convert_word_files([source],work,300,cache_dir=cache)
            self.assertEqual(mock.call_count,1)
            work2=self.base / 'work2';work2.mkdir();status={}
            second=c.convert_word_files([source],work2,300,cache_dir=cache,conversion_status=status)
            self.assertEqual(mock.call_count,1)
            self.assertEqual(first,second);self.assertIn('Reused',status[source])

    def test_word_cache_detects_saved_changes(self):
        source=self.file('word.docx');cache=self.base / 'cache'
        def batch(sources,destination,timeout,converted,total,cancel_event=None):
            for i,p in enumerate(sources): converted[p]=pdf(destination / f'{i}.pdf', (p.read_text(),))
        with patch.object(c,'_run_word_batch',side_effect=batch) as mock:
            for i in range(2):
                work=self.base / f'work{i}';work.mkdir();c.convert_word_files([source],work,300,cache_dir=cache)
                source.write_text('changed')
            self.assertEqual(mock.call_count,2)

    def test_excel_export_readonly_and_print_areas(self):
        source=self.file('work.xlsx'); output=self.run / 'excel.pdf';app=MagicMock()
        book=app.Workbooks.Open.return_value
        book.ExportAsFixedFormat.side_effect=lambda **kw: pdf(Path(kw['Filename']))
        c._office_export(app,source,output,'Excel')
        self.assertTrue(app.Workbooks.Open.call_args.kwargs['ReadOnly'])
        self.assertEqual(app.Workbooks.Open.call_args.kwargs['UpdateLinks'],0)
        self.assertFalse(book.ExportAsFixedFormat.call_args.kwargs['IgnorePrintAreas'])
        book.Close.assert_called_once_with(SaveChanges=False)

    def test_csv_import_retains_text_columns(self):
        source=self.file('table.csv',b'Code,Description\n00123,"Honey, dates"\n'); output=self.run / 'csv.pdf';app=MagicMock()
        app.ActiveWorkbook.ExportAsFixedFormat.side_effect=lambda **kw: pdf(Path(kw['Filename']))
        c._office_export(app,source,output,'Excel')
        options=app.Workbooks.OpenText.call_args.kwargs
        self.assertEqual(options['FieldInfo'],((1,2),(2,2)))
        self.assertEqual(options['Origin'],65001);self.assertTrue(options['Comma']);self.assertFalse(options['Tab'])

    def test_tsv_import_uses_tabs(self):
        source=self.file('table.tsv',b'Code\tName\n001\tHoney\n');output=self.run / 'tsv.pdf';app=MagicMock()
        app.ActiveWorkbook.ExportAsFixedFormat.side_effect=lambda **kw:pdf(Path(kw['Filename']))
        c._office_export(app,source,output,'Excel')
        self.assertTrue(app.Workbooks.OpenText.call_args.kwargs['Tab'])
        self.assertFalse(app.Workbooks.OpenText.call_args.kwargs['Comma'])

    def test_ppt_export_is_pdf_and_open_readonly(self):
        source=self.file('slides.pptx');output=self.run / 'ppt.pdf';app=MagicMock();presentation=app.Presentations.Open.return_value
        presentation.SaveAs.side_effect=lambda name,format:pdf(Path(name))
        c._office_export(app,source,output,'PowerPoint')
        app.Presentations.Open.assert_called_once_with(str(source),-1,0,0)
        presentation.SaveAs.assert_called_once_with(str(output),32)
        self.assertTrue(presentation.Saved);presentation.Close.assert_called_once()

    def test_word_utf8_options_are_explicit(self):
        source=self.file('text.txt','سلام'.encode()); output=self.run / 'txt.pdf'; app=MagicMock()
        doc=app.Documents.Open.return_value
        doc.ExportAsFixedFormat.side_effect=lambda **kw:pdf(Path(kw['OutputFileName']))
        c.export_word(app,source,output)
        self.assertEqual(app.Documents.Open.call_args.kwargs['Encoding'],65001)
        self.assertTrue(app.Documents.Open.call_args.kwargs['ReadOnly'])

    def test_office_retry_is_bounded(self):
        with patch.object(c,'_run_office_once',side_effect=c.WordWorkerError('disconnect',retryable=True)) as mock:
            with self.assertRaises(RuntimeError):c.convert_office(self.file('x.xlsx'),self.run/'x.pdf',2)
        self.assertEqual(mock.call_count,3)

    def test_office_nonretryable_error_stops(self):
        with patch.object(c,'_run_office_once',side_effect=c.WordWorkerError('bad file')) as mock:
            with self.assertRaises(RuntimeError):c.convert_office(self.file('x.xlsx'),self.run/'x.pdf',2)
        self.assertEqual(mock.call_count,1)

    def office_modules(self, app, pid):
        import types
        win32com=types.ModuleType('win32com');client=types.ModuleType('win32com.client')
        client.DispatchEx=MagicMock(return_value=app);win32com.client=client
        process=types.ModuleType('win32process');process.GetWindowThreadProcessId=lambda hwnd:(1,pid)
        return {'pythoncom':MagicMock(),'win32com':win32com,'win32com.client':client,'win32process':process}

    def test_office_worker_refuses_existing_process(self):
        app=MagicMock();app.Hwnd=100;pipe=MagicMock();modules=self.office_modules(app,42)
        with patch.dict(sys.modules,modules):
            c._office_worker(str(self.a),str(self.run/'x.pdf'),'Excel',{42},pipe)
        app.Quit.assert_not_called();app.Workbooks.Open.assert_not_called()
        self.assertTrue(any(x.args[0][0]=='error' for x in pipe.send.call_args_list))

    def test_office_worker_restores_settings_and_exits(self):
        app=MagicMock();app.Hwnd=100;app.AutomationSecurity=1;app.DisplayAlerts=True
        pipe=MagicMock();modules=self.office_modules(app,43)
        with patch.dict(sys.modules,modules),patch.object(c,'_office_export') as export:
            c._office_worker(str(self.a),str(self.run/'x.pdf'),'Excel',{42},pipe)
        export.assert_called_once();app.Quit.assert_called_once()
        self.assertEqual(app.AutomationSecurity,1);self.assertTrue(app.DisplayAlerts)
        self.assertEqual(pipe.send.call_args_list[-1].args[0],('success',))

    def test_office_disconnect_during_close_retains_valid_pdf(self):
        app=MagicMock();book=app.Workbooks.Open.return_value
        source=self.file('x.xlsx');output=self.run/'x.pdf'
        book.ExportAsFixedFormat.side_effect=lambda **kw:pdf(Path(kw['Filename']))
        error=RuntimeError('disconnected');error.hresult=0x80010108;book.Close.side_effect=error
        c._office_export(app,source,output,'Excel')
        self.assertEqual(len(PdfReader(output).pages),1)

    def test_office_close_error_keeps_export_failure(self):
        app=MagicMock();book=app.Workbooks.Open.return_value
        book.ExportAsFixedFormat.side_effect=ValueError('export failed');book.Close.side_effect=RuntimeError('close failed')
        with self.assertRaisesRegex(ValueError,'export failed'):
            c._office_export(app,self.file('x.xlsx'),self.run/'x.pdf','Excel')

    def test_worker_warnings_forwarded_to_parent(self):
        pipe=MagicMock();handler=c._WorkerWarningHandler(pipe)
        record=logging.LogRecord('root',logging.WARNING,'',0,'Retained %s',('file.pdf',),None)
        handler.emit(record);pipe.send.assert_called_once_with(('warning','Retained file.pdf'))

    def test_job_success_event_and_log(self):
        events=queue.Queue();before=list(logging.getLogger().handlers)
        c.gui_merge_job([self.a],self.base/'job',300,False,threading.Event(),events)
        result=list(events.queue);done=[e for e in result if e[0]=='done']
        self.assertEqual(len(done),1); self.assertTrue(Path(done[0][1]).exists())
        self.assertTrue((Path(done[0][1]).parent/'Processing_Log.txt').is_file())
        self.assertEqual(logging.getLogger().handlers,before)

    def test_job_error_does_not_report_success(self):
        events=queue.Queue();c.gui_merge_job([self.file('bad.pdf')],self.base/'job',300,False,threading.Event(),events)
        self.assertEqual([e[0] for e in events.queue if e[0] in {'done','error'}],['error'])

    def test_job_cancel_event(self):
        event=threading.Event();event.set();events=queue.Queue()
        c.gui_merge_job([self.a],self.base/'job',300,False,event,events)
        self.assertEqual([item for item in events.queue if item[0]!='log'],[('cancelled','')])
        self.assertFalse((self.base/'job').exists())

    def test_cli_version(self):
        result=subprocess.run([sys.executable,str(SCRIPT),'--version'],capture_output=True,text=True)
        self.assertEqual(result.returncode,0);self.assertEqual(result.stdout.strip(),'1.1.0')

    def test_cli_preview_includes_supported_images(self):
        self.file('image.png')
        result=subprocess.run([sys.executable,str(SCRIPT),'--folder',str(self.base),'--list'],capture_output=True,text=True)
        self.assertEqual(result.returncode,0);self.assertIn('image.png',result.stdout)
        self.assertIn(self.a.name,result.stdout)

    def test_natural_sort_has_numeric_order(self):
        paths=[Path('Report10.pdf'),Path('Report2.pdf'),Path('Report1.pdf')]
        self.assertEqual([p.name for p in c.ordered_files(paths,'natural')],['Report1.pdf','Report2.pdf','Report10.pdf'])
        self.assertEqual([p.name for p in c.ordered_files(paths,'filename')],['Report1.pdf','Report10.pdf','Report2.pdf'])

    def test_selected_sort_keeps_manual_order(self):
        self.assertEqual(c.ordered_files([self.b,self.a],'selected'),[self.b,self.a])

    def test_custom_output_filename(self):
        output=c.combine_selected([self.a,self.b],self.run,output_name='My_Document_Pack')
        self.assertEqual(output.name,'My_Document_Pack.pdf')
        self.assertEqual(len(PdfReader(output).pages),3)
        self.assertFalse(list(self.run.glob('*.pending')))

    def test_output_name_rejects_paths_and_reserved_names(self):
        for name in ['../outside.pdf','C:\\outside.pdf','NUL.pdf','con.pdf','COM1.pdf','a?.pdf','.pdf','']:
            with self.subTest(name=name),self.assertRaises(ValueError):c.validate_output_name(name)

    def test_output_name_preserves_unicode(self):
        self.assertEqual(c.validate_output_name('رپورٹ.pdf'),'رپورٹ.pdf')

    def test_invalid_output_name_writes_no_pdf(self):
        with self.assertRaises(ValueError):c.combine_selected([self.a],self.run,output_name='../outside.pdf')
        self.assertFalse((self.run.parent/'outside.pdf').exists())
        self.assertEqual(list(self.run.iterdir()),[])

    def test_folder_setup_error_has_actionable_context(self):
        with patch.object(Path,'mkdir',side_effect=FileNotFoundError(2,'Cannot find path','example')):
            with self.assertRaisesRegex(c.OutputSetupError,'Use Browse'):
                c.create_run_directory(self.base/'output')

    def test_startup_error_has_real_fallback_log(self):
        events=queue.Queue()
        with patch.object(c,'create_run_directory',side_effect=c.OutputSetupError('Cannot create result folder')):
            c.gui_merge_job([self.a],self.base/'missing',300,False,threading.Event(),events)
        error=[e for e in events.queue if e[0]=='error'][0]
        self.assertEqual(error[2],'')
        log=Path(error[3]);self.addCleanup(log.unlink,missing_ok=True)
        self.assertTrue(log.is_file());self.assertIn('Cannot create result folder',log.read_text())

    def test_log_open_failure_has_real_fallback_log(self):
        events=queue.Queue()
        with patch.object(logging,'FileHandler',side_effect=PermissionError('Cannot write log')):
            c.gui_merge_job([self.a],self.base/'output',300,False,threading.Event(),events)
        error=[e for e in events.queue if e[0]=='error'][0]
        self.assertTrue(Path(error[2]).is_dir())
        log=Path(error[3]);self.addCleanup(log.unlink,missing_ok=True)
        self.assertTrue(log.is_file());self.assertIn('Cannot write log',log.read_text())

    def test_diagnostic_failure_does_not_hide_original_error(self):
        events=queue.Queue()
        with patch.object(c,'create_run_directory',side_effect=c.OutputSetupError('Original folder error')):
            with patch.object(c.tempfile,'NamedTemporaryFile',side_effect=OSError('Temp unavailable')):
                c.gui_merge_job([self.a],self.base/'output',300,False,threading.Event(),events)
        error=[e for e in events.queue if e[0]=='error'][0]
        self.assertEqual(error[1],'Original folder error');self.assertEqual(error[3],'')

    def test_folder_discovery_includes_images_and_skips_cache(self):
        image=self.file('scan.png');cache=self.base/c.WORD_CACHE_FOLDER;cache.mkdir()
        pdf(cache/'cached.pdf')
        discovered=c.discover(self.base,recursive=True)
        self.assertIn(image,discovered);self.assertNotIn(cache/'cached.pdf',discovered)

    def test_cli_explicit_files_with_custom_name(self):
        output=self.base/'cli_output'
        result=subprocess.run([sys.executable,str(SCRIPT),'--files',str(self.b),str(self.a),
                               '--output-dir',str(output),'--output-name','Selected.pdf','--yes'],capture_output=True,text=True)
        self.assertEqual(result.returncode,0,result.stdout+result.stderr)
        outputs=list(output.glob('*/Selected.pdf'));self.assertEqual(len(outputs),1)
        self.assertEqual([p.extract_text().strip() for p in PdfReader(outputs[0]).pages],['Beta','Gamma','Alpha'])

    def test_cli_confirmation_can_cancel(self):
        output=self.base/'not_created'
        result=subprocess.run([sys.executable,str(SCRIPT),'--files',str(self.a),'--output-dir',str(output)],
                               input='NO\n',capture_output=True,text=True)
        self.assertEqual(result.returncode,0);self.assertIn('Cancelled',result.stdout);self.assertFalse(output.exists())

    def test_cli_natural_folder_preview(self):
        self.file('file10.png');self.file('file2.png')
        result=subprocess.run([sys.executable,str(SCRIPT),'--folder',str(self.base),'--list','--sort','natural'],capture_output=True,text=True)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertLess(result.stdout.index('file2.png'),result.stdout.index('file10.png'))

    def test_general_purpose_help(self):
        result=subprocess.run([sys.executable,str(SCRIPT),'--help'],capture_output=True,text=True)
        self.assertEqual(result.returncode,0)
        self.assertIn('--files',result.stdout);self.assertIn('--output-name',result.stdout)



if __name__ == '__main__':
    unittest.main(verbosity=1)
