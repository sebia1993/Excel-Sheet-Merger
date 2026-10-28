"""Fresh-process GUI checks also fail on Tcl errors outside unittest callbacks."""
from __future__ import annotations

import os
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUI_AVAILABLE = os.name == 'nt' or bool(os.environ.get('DISPLAY'))


@unittest.skipUnless(GUI_AVAILABLE, 'requires a Tk display')
class GuiThemeLifecycleTests(unittest.TestCase):
    def run_gui(self, script: str) -> None:
        prefix = '''
import sys
sys.path.insert(0, 'src')
from excel_sheet_merger.gui import MergeApp
errors = []
def observe(app):
    app.tk.createcommand('esm_capture_error', lambda message, details: errors.append((message, details)) or '')
    app.tk.eval('proc bgerror {message} {esm_capture_error $message $::errorInfo}')
    app.report_callback_exception = lambda kind, value, tb: errors.append((str(value), kind.__name__))
'''
        completed = subprocess.run(
            [sys.executable, '-c', prefix + textwrap.dedent(script)],
            cwd=ROOT, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=30,
        )
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)

    def test_first_render_and_scrolling_do_not_raise_background_color_errors(self):
        self.run_gui('''
            app = MergeApp()
            observe(app)
            try:
                for i in range(200):
                    app.tree.insert('', 'end', text=str(i), values=('11시', 'test.csv'))
                app._show_log()
                app._write_log('\\n'.join(str(i) for i in range(100)))
                app.update()
                app.tree.yview_moveto(0.7)
                app.log.yview_moveto(0.5)
                app.update()
                assert app.tree.yview()[0] > 0
                assert not errors, errors
            finally:
                app.destroy()
        ''')

    def test_reopened_window_styles_belong_to_current_interpreter(self):
        self.run_gui('''
            from excel_sheet_merger.ui_theme import COLORS
            for i in range(3):
                app = MergeApp()
                observe(app)
                try:
                    assert app.theme_style.tk is app.tk, 'theme still belongs to a destroyed interpreter'
                    app.update()
                    color = app.tk.call('ttk::style', 'lookup', 'Office.Primary.TButton', '-background')
                    assert color == COLORS['primary'], color
                finally:
                    app.destroy()
            assert not errors, errors
        ''')

    def test_destroy_releases_only_owned_theme_instance(self):
        self.run_gui('''
            import importlib.util
            if importlib.util.find_spec('ttkbootstrap') is None:
                sys.exit(0)
            from ttkbootstrap import Style
            app = MergeApp()
            observe(app)
            app.update_idletasks()
            owned = app.theme_style
            assert Style.instance is owned
            app.destroy()
            assert Style.instance is None, 'destroyed root remains referenced by global theme'
        ''')

    def test_destroying_child_does_not_clear_live_root_theme(self):
        self.run_gui('''
            import importlib.util
            if importlib.util.find_spec('ttkbootstrap') is None:
                sys.exit(0)
            from tkinter import ttk
            from ttkbootstrap import Style
            app = MergeApp()
            observe(app)
            try:
                owned = app.theme_style
                child = ttk.Label(app, text='transient')
                child.destroy()
                app.update()
                assert Style.instance is owned, 'child teardown cleared the live root theme'
                assert not errors, errors
            finally:
                app.destroy()
        ''')

    def test_small_failure_layout_preserves_text_and_restores_original_spacing(self):
        self.run_gui('''
            from excel_sheet_merger.failure_guidance import diagnose_failure
            from excel_sheet_merger.merge import SourceFileError
            from pathlib import Path
            app = MergeApp()
            observe(app)
            try:
                app.update()
                root_frame = app.winfo_children()[0]
                original_padding = tuple(root_frame.cget('padding'))
                # Real decoding diagnosis, not a fabricated short placeholder.
                error = UnicodeDecodeError('cp949', b'\\xff', 0, 1, 'invalid start byte')
                advice = diagnose_failure(SourceFileError(Path('A사업장_20261028_14시.csv'), error))
                app._register_failure(advice)
                app._write_log(advice.reason + advice.action)
                message = app.failure_text.get()
                log_before = app.log.get('1.0', 'end')
                app.geometry('900x640')
                app.update()
                assert app.tree.winfo_height() >= 70, app.tree.winfo_height()
                assert advice.reason in app.failure_text.get()
                assert advice.action in app.failure_text.get()
                app.geometry('1100x760')
                app.update()
                assert tuple(root_frame.cget('padding')) == original_padding
                assert app.failure_text.get() == message
                assert app.log.get('1.0', 'end') == log_before
                app.geometry('900x640')
                app.update()
                app._reset_failures()
                app.update()
                assert tuple(root_frame.cget('padding')) == original_padding
                assert not errors, errors
            finally:
                app.destroy()
        ''')
