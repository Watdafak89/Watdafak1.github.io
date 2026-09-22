from io import BytesIO
import json
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
from zipfile import ZipFile

from lxml import etree
from pypdf import PdfWriter

from teaching_plan import (
    END, NS, START, PlanError, generate_plan, inspect_template,
    render_template, text_of, validate_pdf, validate_plan, gemini_schema, provider_error,
)


def template_bytes(extra=''):
    def row(text):
        return '<w:tr><w:tc><w:p><w:r><w:t>' + text + '</w:t></w:r></w:p></w:tc></w:tr>'
    xml = ('<w:document xmlns:w="' + NS['w'] + '"><w:body><w:p>'
           '<w:r><w:rPr><w:b/></w:rPr><w:t>วิชา {{ sub</w:t></w:r>'
           '<w:r><w:t>ject }} / {{ code }}</w:t></w:r></w:p><w:tbl>'
           + row('หัวตาราง') + row(START)
           + row(' '.join('{{ r.' + key + ' }}' for key in ['w', 't', 'p', 'a', 'm', 'e']))
           + row(END) + row('') + '</w:tbl>' + extra + '<w:sectPr/></w:body></w:document>')
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        archive.writestr('word/document.xml', xml.encode())
        archive.writestr('word/styles.xml', b'<unchanged/>')
        archive.writestr('word/media/image1.png', b'unchanged image')
        archive.writestr('customXml/item1.xml', b'<opaque/>')
    return output.getvalue()


def plan_data(n=2):
    return dict(curriculum='ปวช. 2567', code='21901-2015', subject='ระบบเครือข่าย',
                level='ปวช.', year_level='', h=4, n=n, term='', notes=[],
                weeks=[dict(w=i, t=f'หัวข้อ {i}', p='อุปกรณ์ & สาย <LAN>',
                            a='ฝึกปฏิบัติ', m='สาย UTP', e='ตรวจชิ้นงาน') for i in range(1, n + 1)])


def pdf_bytes():
    writer = PdfWriter()
    writer.add_blank_page(width=595, height=842)
    output = BytesIO()
    writer.write(output)
    return output.getvalue()


class TemplateTests(unittest.TestCase):
    def test_fills_form_numbers_and_keeps_existing_numbers(self):
        extra = '<w:p><w:r><w:t>แผ่นที่ : </w:t></w:r></w:p><w:p><w:r><w:t>หน้า</w:t></w:r><w:r><w:t>ที่</w:t></w:r></w:p><w:p><w:r><w:t>หน้าที่ 99</w:t></w:r></w:p>'
        output = render_template(template_bytes(extra), plan_data(), sheet_number=2, page_number=7)
        with ZipFile(BytesIO(output)) as archive:
            root = etree.fromstring(archive.read('word/document.xml'))
            paragraphs = [text_of(p) for p in root.xpath('.//w:p', namespaces=NS)]
        self.assertIn('แผ่นที่ :  2', paragraphs)
        self.assertIn('หน้าที่ 7', paragraphs)
        self.assertIn('หน้าที่ 99', paragraphs)

    def test_split_tokens_rows_and_package_preservation(self):
        source = template_bytes()
        result = render_template(source, plan_data())
        with ZipFile(BytesIO(source)) as original, ZipFile(BytesIO(result)) as output:
            self.assertEqual(original.namelist(), output.namelist())
            for name in original.namelist():
                if name != 'word/document.xml':
                    self.assertEqual(original.read(name), output.read(name))
            root = etree.fromstring(output.read('word/document.xml'))
            content = text_of(root)
            self.assertIn('วิชา ระบบเครือข่าย / 21901-2015', content)
            self.assertIn('อุปกรณ์ & สาย <LAN>', content)
            self.assertNotIn('{{', content)
            self.assertNotIn('{%', content)
            self.assertEqual(len(root.xpath('.//w:tr', namespaces=NS)), 3)
            self.assertTrue(root.xpath('.//w:rPr/w:b', namespaces=NS))

    def test_rejects_template_expressions_without_executing(self):
        for text in ['{{ secret }}', '{{ code.__class__ }}', '{% include "file" %}', '{{ r.w }}']:
            with self.subTest(text=text), self.assertRaises(PlanError):
                inspect_template(template_bytes('<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>'))

    def test_rejects_missing_and_duplicate_weeks(self):
        data = plan_data()
        data['weeks'][1]['w'] = 1
        with self.assertRaises(PlanError):
            validate_plan(data)
        with self.assertRaises(PlanError):
            validate_plan(plan_data(), {'n': 3})

    def test_user_metadata_wins_and_blank_keeps_pdf_values(self):
        plan = validate_plan(plan_data(), {'code': 'CUSTOM', 'subject': '', 'h': 0})
        self.assertEqual(plan.code, 'CUSTOM')
        self.assertEqual(plan.subject, 'ระบบเครือข่าย')
        self.assertEqual(plan.h, 4)

    def test_pdf_validation(self):
        self.assertEqual(validate_pdf(pdf_bytes()), 1)
        with self.assertRaises(PlanError):
            validate_pdf(b'not pdf')
        writer = PdfWriter()
        writer.add_blank_page(width=595, height=842)
        writer.encrypt('password')
        output = BytesIO()
        writer.write(output)
        with self.assertRaises(PlanError):
            validate_pdf(output.getvalue())


class GeminiTests(unittest.TestCase):
    def test_sdk_accepts_schema_before_network_call(self):
        from google import genai
        from google.genai import types
        from teaching_plan import TeachingPlan
        response = {'candidates': [{'content': {'role': 'model', 'parts': [
            {'text': json.dumps(plan_data())}]}, 'finishReason': 'STOP'}]}
        with genai.Client(api_key='test-key') as client, patch.object(
            client._api_client, 'request', return_value=SimpleNamespace(
                headers={}, body=json.dumps(response))
        ) as request:
            output = client.models.generate_content(model='gemini-2.5-flash', contents='test',
                config=types.GenerateContentConfig(response_mime_type='application/json', response_json_schema=gemini_schema()))
            self.assertEqual(json.loads(output.text)['n'], 2)
            request.assert_called_once()
            wire_config = request.call_args.args[2]['generationConfig']
            self.assertNotIn('responseSchema', wire_config)
            self.assertEqual(wire_config['responseJsonSchema'], gemini_schema())
            serialized = json.dumps(wire_config['responseJsonSchema'])
            for unsupported in ('additional_properties', 'max_length', '$ref', '$defs', 'maxItems'):
                self.assertNotIn(unsupported, serialized)

    def test_error_details_redact_keys_and_identify_stage(self):
        error = SimpleNamespace(code=400, message='API key not valid. test-secret https://example.test/?key=test-secret')
        message = provider_error(error, 'test-secret', 'ส่งไฟล์ PDF')
        self.assertIn('ปฏิเสธ API Key', message)
        self.assertIn('ส่งไฟล์ PDF', message)
        self.assertIn('HTTP 400', message)
        self.assertNotIn('test-secret', message)
        self.assertNotIn('https://', message)

    def test_schema_error_preserves_diagnostic(self):
        error = SimpleNamespace(code=400, message='response schema has too many states')
        message = provider_error(error, 'test-key', 'วิเคราะห์แผนการสอน')
        self.assertIn('รูปแบบคำตอบ', message)
        self.assertIn('too many states', message)

    def client(self):
        client = Mock()
        client.files.upload.return_value = SimpleNamespace(name='files/test', state=SimpleNamespace(name='ACTIVE'))
        client.models.generate_content.return_value = SimpleNamespace(text=json.dumps(plan_data()))
        return client

    def test_generation_sends_pdf_schema_and_cleans_up(self):
        client = self.client()
        messages = []
        result = generate_plan('test-key', pdf_bytes(), template_bytes(), {'n': 2},
                               client_factory=Mock(return_value=client), progress=messages.append)
        self.assertEqual(result.n, 2)
        config = client.models.generate_content.call_args.kwargs['config']
        self.assertEqual(config.response_mime_type, 'application/json')
        self.assertIsNone(config.response_schema)
        self.assertEqual(config.response_json_schema, gemini_schema())
        self.assertIn('เอกสารอ้างอิงเท่านั้น', config.system_instruction)
        self.assertEqual(client.files.upload.call_args.kwargs['file'].getvalue(), pdf_bytes())
        client.files.delete.assert_called_once_with(name='files/test')
        client.close.assert_called_once()

    def test_cleanup_and_safe_error_on_quota_failure(self):
        client = self.client()
        error = RuntimeError('private key or provider data must not be displayed')
        error.code = 429
        client.models.generate_content.side_effect = error
        with self.assertRaisesRegex(PlanError, 'โควตา') as caught:
            generate_plan('test-key', pdf_bytes(), template_bytes(), {}, client_factory=Mock(return_value=client))
        self.assertNotIn('private', str(caught.exception))
        client.files.delete.assert_called_once()

    def test_cleanup_on_invalid_response(self):
        client = self.client()
        client.models.generate_content.return_value.text = '{unfinished'
        with self.assertRaises(PlanError):
            generate_plan('test-key', pdf_bytes(), template_bytes(), {}, client_factory=Mock(return_value=client))
        client.files.delete.assert_called_once()

    def test_invalid_input_never_calls_provider(self):
        factory = Mock()
        with self.assertRaises(PlanError):
            generate_plan('test-key', b'bad pdf', template_bytes(), {}, client_factory=factory)
        factory.assert_not_called()


if __name__ == '__main__':
    unittest.main()
