# test_export_letters.py
import unittest
from unittest.mock import patch, MagicMock, mock_open
import os
from export_letters import process_document, main

class TestExportLetters(unittest.TestCase):
    def setUp(self):
        self.test_docx = 'test.docx'
        self.output_docx = 'output.docx'
        self.replacements = {'#TEST#': 'replacement'}

    @patch('export_letters.ZipFile')
    def test_process_document(self, mock_zip_class):
        # Setup mock ZIP file
        mock_zip = MagicMock()
        mock_zip_class.return_value.__enter__.return_value = mock_zip
        
        # Mock file content
        mock_content = b'''<?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:p><w:r><w:t>Test #TEST# content</w:t></w:r></w:p>
            </w:document>'''
        
        mock_zip.read.return_value = mock_content
        mock_zip.filelist = [MagicMock(filename='word/document.xml')]

        # Call process_document
        process_document(self.test_docx, self.output_docx, self.replacements)

        # Verify ZIP operations
        mock_zip.read.assert_called_with('word/document.xml')
        self.assertTrue(mock_zip_class.called)

    @patch('export_letters.create_engine')
    @patch('export_letters.sessionmaker')
    def test_main(self, mock_sessionmaker, mock_create_engine):
        # Setup mock database session
        mock_session = MagicMock()
        mock_sessionmaker.return_value.return_value = mock_session
        
        # Mock database query result
        mock_result = {
            'Interviniente_Nombre': 'Test Name',
            'Agencia': 'Test Agency',
            'Telefono': '123456789'
        }
        mock_session.execute.return_value.first.return_value = mock_result

        # Call main function
        main()

        # Verify session operations
        mock_session.execute.assert_called_once()
        mock_session.close.assert_called_once()
