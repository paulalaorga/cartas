import unittest
import tempfile
import os
from lxml import etree
from export_letters import process_paragraph

class TestDocumentProcessing(unittest.TestCase):
    def setUp(self):
        self.namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        self.temp_dir = tempfile.mkdtemp()
        self.test_docx = os.path.join(self.temp_dir, "test.docx")
        self.output_docx = os.path.join(self.temp_dir, "output.docx")

    def test_process_paragraph(self):
        # Create test paragraph
        paragraph = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
        
        # First run with replacement text
        run1 = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        rPr1 = etree.SubElement(run1, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        etree.SubElement(rPr1, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b")  # Bold
        text1 = etree.SubElement(run1, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        text1.text = "Hello #NAME#!"
        
        # Second run with following text
        run2 = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        text2 = etree.SubElement(run2, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        text2.set(f'{{{self.namespaces["w"]}}}space', 'preserve')
        text2.text = " How are you?"

        replacements = {"#NAME#": "John"}
        process_paragraph(paragraph, replacements, self.namespaces)
        
        # Verify results
        result_runs = paragraph.findall('.//w:r', namespaces=self.namespaces)
        self.assertEqual(len(result_runs), 2, "Expected exactly 2 runs")
        
        # Check first run
        first_text = result_runs[0].find('.//w:t', namespaces=self.namespaces)
        self.assertEqual(first_text.text, "Hello John!")
        self.assertTrue(result_runs[0].find('.//w:b', namespaces=self.namespaces) is not None)
        
        # Check second run
        second_text = result_runs[1].find('.//w:t', namespaces=self.namespaces)
        self.assertEqual(second_text.text, " How are you?")
        self.assertEqual(second_text.get(f'{{{self.namespaces["w"]}}}space'), 'preserve')

    def test_whitespace_attributes(self):
        # Test whitespace preservation
        paragraph = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
        run = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        text = etree.SubElement(run, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        text.set(f'{{{self.namespaces["w"]}}}space', 'preserve')
        text.text = "  #TEXT#  "
        
        replacements = {"#TEXT#": "replacement"}
        process_paragraph(paragraph, replacements, self.namespaces)
        
        # Verify results
        result = paragraph.find('.//w:t', namespaces=self.namespaces)
        self.assertEqual(result.get(f'{{{self.namespaces["w"]}}}space'), 'preserve')
        self.assertEqual(result.text, "  replacement  ")

    def tearDown(self):
        for file in [self.test_docx, self.output_docx]:
            if os.path.exists(file):
                os.remove(file)
        os.rmdir(self.temp_dir)

if __name__ == '__main__':
    unittest.main()