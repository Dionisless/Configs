import os
import tempfile
import unittest
import zipfile

from src.container_reader import extract_container


class ContainerReaderTests(unittest.TestCase):
    def test_extract_zip(self):
        with tempfile.TemporaryDirectory() as td:
            zip_path = os.path.join(td, 'sample.zip')
            with zipfile.ZipFile(zip_path, 'w') as zf:
                zf.writestr('PROT001.bin', b'0A0B Param 10')

            out = extract_container(zip_path)
            self.assertEqual(out.source_type, 'zip')
            self.assertTrue(os.path.exists(os.path.join(out.work_dir, 'PROT001.bin')))


if __name__ == '__main__':
    unittest.main()
