import os
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from backend.app.storage import VOLUME_PATH_ENV, VolumeArchive


class VolumeArchiveTests(unittest.TestCase):
    def test_disabled_archive_does_not_write(self) -> None:
        archive = VolumeArchive()

        result = archive.write("source", "sample.docx", b"content", "run-1")

        self.assertIsNone(result)

    def test_archive_uses_category_date_and_safe_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive = VolumeArchive(Path(temp_dir))
            created_at = datetime(2026, 9, 17, tzinfo=timezone.utc)

            result = archive.write(
                "source",
                "../Product specification - smor.docx",
                b"document content",
                "../../unsafe processing id",
                created_at,
            )

            expected = (
                Path(temp_dir)
                / "source"
                / "2026"
                / "09"
                / "17"
                / "unsafe_processing_id_Product_specification_-_smor.docx"
            )
            self.assertEqual(result, expected)
            self.assertEqual(expected.read_bytes(), b"document content")

    def test_archive_rejects_unknown_category(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive = VolumeArchive(Path(temp_dir))

            with self.assertRaises(ValueError):
                archive.write("other", "sample.docx", b"content", "run-1")

    def test_archive_reads_volume_path_from_environment(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.dict(os.environ, {VOLUME_PATH_ENV: temp_dir}):
                archive = VolumeArchive.from_environment()

            self.assertTrue(archive.enabled)
            self.assertEqual(archive.root, Path(temp_dir))


if __name__ == "__main__":
    unittest.main()
