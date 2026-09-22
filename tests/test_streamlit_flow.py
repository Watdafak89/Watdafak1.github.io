from pathlib import Path
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from teaching_plan import validate_plan
from test_teaching_plan import pdf_bytes, plan_data, template_bytes


class StreamlitFlowTests(unittest.TestCase):
    def test_upload_generate_rerun_and_invalidate_download(self):
        app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'app.py')).run()
        self.assertFalse(app.exception)
        uploaders = app.get('file_uploader')
        if not hasattr(uploaders[0], 'upload'):
            self.skipTest('File upload testing requires a newer Streamlit test runtime')
        uploaders[0].upload('template.docx', template_bytes()).run()
        app.get('file_uploader')[1].upload('plan.pdf', pdf_bytes(), 'application/pdf').run()
        self.assertFalse(app.exception)
        app.text_input[0].set_value('test-key').run()
        with patch('teaching_plan.generate_plan', return_value=validate_plan(plan_data(18))) as generate:
            app.button[-1].click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.get('download_button')), 1)
        self.assertEqual(generate.call_count, 1)
        app.run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.get('download_button')), 1)
        # A changed course must never silently download the previous course's plan.
        app.radio[0].set_value('กรอกเอง').run()
        next(x for x in app.text_input if x.label == 'รหัสวิชา').set_value('NEW-CODE').run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.get('download_button')), 0)

    def test_missing_uploads_shows_actionable_error(self):
        app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'app.py')).run()
        app.button[-1].click().run()
        self.assertFalse(app.exception)
        self.assertIn('template DOCX', app.error[0].value)


if __name__ == '__main__':
    unittest.main()
