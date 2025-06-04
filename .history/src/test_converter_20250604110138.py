import os
import unittest
from converter import PDFConverter

class TestPDFConverter(unittest.TestCase):
    def setUp(self):
        self.converter = PDFConverter()
        self.test_output_dir = "test_output"
        os.makedirs(self.test_output_dir, exist_ok=True)

    def tearDown(self):
        # Clean up test output directory
        for file in os.listdir(self.test_output_dir):
            os.remove(os.path.join(self.test_output_dir, file))
        os.rmdir(self.test_output_dir)

    def test_language_detection(self):
        # Test English text
        self.assertEqual(self.converter.detect_language("Hello World"), "en")
        
        # Test Arabic text
        self.assertEqual(self.converter.detect_language("مرحبا بالعالم"), "ar")
        
        # Test Urdu text
        self.assertEqual(self.converter.detect_language("ہیلو ورلڈ"), "ur")
        
        # Test mixed text
        self.assertEqual(self.converter.detect_language("Hello مرحبا"), "ar")

    def test_empty_text(self):
        self.assertEqual(self.converter.detect_language(""), "en")
        self.assertEqual(self.converter.detect_language(None), "en")

if __name__ == '__main__':
    unittest.main() 