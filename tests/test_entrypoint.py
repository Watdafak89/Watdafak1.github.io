"""Check Cloud entrypoint reruns without requiring a running Streamlit server."""
import os
from pathlib import Path
import runpy
import shutil
import sys
import tempfile
import types
import unittest
from unittest.mock import patch


class EntrypointTests(unittest.TestCase):
    def test_streamlit_reexecutes_page_after_widget_change(self):
        flask = types.ModuleType('flask')
        flask.Flask = lambda name: types.SimpleNamespace(get=lambda route: lambda fn: fn)
        flask.render_template = lambda name: name
        streamlit = types.ModuleType('streamlit')
        streamlit.render_count = 0
        with tempfile.TemporaryDirectory() as folder:
            entrypoint = Path(folder) / 'app.py'
            shutil.copyfile(Path(__file__).resolve().parents[1] / 'app.py', entrypoint)
            (Path(folder) / 'streamlit_app.py').write_text(
                'import streamlit\nstreamlit.render_count += 1\n', encoding='utf-8'
            )
            with patch.dict(sys.modules, {'flask': flask, 'streamlit': streamlit}), patch.dict(
                os.environ, {'STREAMLIT_SERVER_PORT': '8501'}
            ), patch.object(sys, 'path', [folder, *sys.path]):
                runpy.run_path(str(entrypoint), run_name='__main__')
                self.assertEqual(streamlit.render_count, 1)
                runpy.run_path(str(entrypoint), run_name='__main__')
                self.assertEqual(streamlit.render_count, 2)


if __name__ == '__main__':
    unittest.main()
