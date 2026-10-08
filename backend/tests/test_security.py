import unittest

from backend.app.main import has_expected_file_signature, safe_download_name


class SecurityHelpersTests(unittest.TestCase):
    def test_file_signatures_must_match_the_extension(self) -> None:
        self.assertTrue(has_expected_file_signature(b"%PDF-1.7\n", ".pdf"))
        self.assertTrue(has_expected_file_signature(b"PK\x03\x04docx content", ".docx"))
        self.assertFalse(has_expected_file_signature(b"not a pdf", ".pdf"))
        self.assertFalse(has_expected_file_signature(b"%PDF-1.7\n", ".docx"))

    def test_download_name_cannot_inject_headers_or_paths(self) -> None:
        name = safe_download_name('../../Product "A"\r\nX-Header: unsafe')

        self.assertEqual(name, "Product-A-X-Header-unsafe")
        self.assertNotIn("\r", name)
        self.assertNotIn("\n", name)


if __name__ == "__main__":
    unittest.main()
