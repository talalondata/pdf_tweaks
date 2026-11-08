import importlib.util
import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
import unittest
import random
from unittest.mock import MagicMock, patch

BASE=Path(__file__).resolve().parents[1]

import tkinter as tk
from PIL import Image, ImageGrab
from pypdf import PdfReader
from reportlab.pdfgen.canvas import Canvas

spec=importlib.util.spec_from_file_location('combiner_gui',BASE/'pdf_tweaks.py')
c=importlib.util.module_from_spec(spec);sys.modules[spec.name]=c;spec.loader.exec_module(c)

@unittest.skipUnless(os.environ.get('RUN_GUI_TESTS') == '1', 'Set RUN_GUI_TESTS=1 with a desktop or virtual display')
class GuiTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.base=Path(self.temp.name)
        self.a=self.base/'02 - Report.pdf';self.b=self.base/'10 - Appendix.pdf'
        for p,text in [(self.a,'Report sample'),(self.b,'Appendix sample')]:
            canvas=Canvas(str(p));canvas.drawString(50,700,text);canvas.showPage();canvas.save()
        self.root=tk.Tk();self.app=c.CombinerWindow(self.root)
        self.app.message=MagicMock();self.app.message.askyesno.return_value=True
        self.app.folder.set(str(self.base/'results'))
        self.root.update()

    def tearDown(self):
        self.app.close();self.temp.cleanup()
        self.app=None;self.root=None
        import gc
        gc.collect()

    def add(self):
        self.app.pick_files=MagicMock(return_value=[str(self.b),str(self.a)])
        self.app.add_files();self.root.update()

    def test_add_and_sort_display_order(self):
        self.add();self.assertEqual(self.app.model.files,[self.a,self.b])
        self.assertEqual(self.app.tree.item('0')['values'][1],self.a.name)

    def test_move_and_remove(self):
        self.add();self.app.tree.selection_set('1');self.app.move(-1)
        self.assertEqual(self.app.model.files,[self.b,self.a]);self.app.remove()
        self.assertEqual(self.app.model.files,[self.a]);self.assertTrue(self.b.exists())

    def test_replace_selected(self):
        self.add();self.app.tree.selection_set('0');p=self.base/'updated.png';Image.new('RGB',(20,20)).save(p)
        self.app.pick_files=MagicMock(return_value=str(p));self.app.replace_file()
        self.assertEqual(self.app.model.files,[p,self.b]);self.assertTrue(self.a.exists())

    def test_same_stem_asks_replace(self):
        old=self.base/'10 - Appendix.docx';old.write_text('example')
        self.app.model.add(old);self.app.redraw();self.app.pick_files=MagicMock(return_value=[str(self.b)])
        self.app.message.askyesnocancel.return_value=True;self.app.add_files()
        self.app.message.askyesnocancel.assert_called_once();self.assertEqual(self.app.model.files,[self.b])
        self.assertTrue(old.exists())


    def test_no_merge_before_confirmation(self):
        self.add();self.app.message.askyesno.return_value=False;self.app.start()
        self.assertFalse(self.app.running);self.assertFalse((self.base/'results').exists())

    def test_complete_merge_via_window(self):
        self.add();self.app.start()
        deadline=time.monotonic()+5
        while self.app.running and time.monotonic()<deadline:
            self.root.update();time.sleep(.01)
        self.assertFalse(self.app.running);self.assertIsNotNone(self.app.output)
        self.assertEqual(len(PdfReader(self.app.output).pages),2)
        self.assertEqual(str(self.app.open_button.cget('state')),'normal')
        self.app.message.showerror.assert_not_called()

    def test_empty_selection_cannot_start(self):
        self.app.start();self.app.message.showerror.assert_called_once();self.assertFalse(self.app.running)

    def test_custom_filename_through_window(self):
        self.add();self.app.output_name.set('General_Document_Pack.pdf');self.app.start()
        deadline=time.monotonic()+5
        while self.app.running and time.monotonic()<deadline:
            self.root.update();time.sleep(.01)
        self.assertFalse(self.app.running);self.assertEqual(self.app.output.name,'General_Document_Pack.pdf')

    def test_natural_sort_through_window(self):
        a=self.base/'file2.pdf';b=self.base/'file10.pdf'
        self.app.model.files=[b,a];self.app.redraw();self.app.sort_mode.set('Natural (2 before 10)');self.app.sort_files()
        self.assertEqual(self.app.model.files,[a,b])

    def test_startup_failure_shows_actual_diagnostic(self):
        self.add()
        with patch.object(c,'create_run_directory',side_effect=c.OutputSetupError('Choose another result folder')):
            self.app.start();deadline=time.monotonic()+5
            while self.app.running and time.monotonic()<deadline:
                self.root.update();time.sleep(.01)
        self.assertFalse(self.app.running);self.assertIsNone(self.app.output);self.assertIsNone(self.app.run_folder)
        detail=self.app.message.showerror.call_args.args[1]
        self.assertIn('Diagnostic log:',detail)
        diagnostic=Path(detail.split('Diagnostic log: ',1)[1]);self.assertTrue(diagnostic.is_file());diagnostic.unlink()

    def test_invalid_filename_stops_before_processing(self):
        self.add();self.app.output_name.set('../outside.pdf');self.app.start()
        self.assertFalse(self.app.running);self.app.message.showerror.assert_called_once()

    def test_cancel_button(self):
        def cancellable(files,folder,timeout,fresh,cancel,events,output_name="Combined_Files.pdf"):
            cancel.wait(2);events.put(('cancelled',''))
        self.add()
        with patch.object(c,'gui_merge_job',side_effect=cancellable):
            self.app.start();self.app.cancel();deadline=time.monotonic()+3
            while self.app.running and time.monotonic()<deadline:
                self.root.update();time.sleep(.01)
        self.assertFalse(self.app.running);self.assertIn('Cancelled',self.app.status.get())

    def test_actions_fit_at_smallest_window_size(self):
        self.root.geometry('920x680');self.root.update()
        create=[x for x in self.app.controls if x.winfo_class()=='TButton' and x.cget('text')=='Create combined PDF'][0]
        for widget in (create, self.app.cancel_button, self.app.github_button,
                       self.app.folder_button, self.app.open_button):
            with self.subTest(control=widget.cget('text')):
                self.assertTrue(widget.winfo_viewable())
                left=widget.winfo_rootx()-self.root.winfo_rootx()
                top=widget.winfo_rooty()-self.root.winfo_rooty()
                self.assertGreaterEqual(left, 0)
                self.assertGreaterEqual(top, 0)
                self.assertLessEqual(left+widget.winfo_width(), self.root.winfo_width())
                self.assertLessEqual(top+widget.winfo_height(), self.root.winfo_height())

    def test_screenshot_populated_window(self):
        self.add();image=self.base/'300 - Site inspection photographs.png';Image.new('RGB',(120,80),'#9bc9e8').save(image)
        self.app.model.add(image);self.app.sort_files();self.root.update()
        screenshot=self.base/'interface.png'
        ImageGrab.grab().save(screenshot)
        self.assertTrue(screenshot.is_file())

if __name__ == '__main__':
    unittest.main()
