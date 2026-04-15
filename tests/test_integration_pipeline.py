import os
import tempfile
import unittest

from src.cli import run_pipeline


class IntegrationPipelineTests(unittest.TestCase):
    def test_full_pipeline(self):
        with tempfile.TemporaryDirectory() as td:
            prot_path = os.path.join(td, 'PROT_A.bin')
            par_path = os.path.join(td, 'PAR_A.bin')
            with open(prot_path, 'wb') as f:
                f.write(b'0A0B Voltage 220 0A0C Current 5')
            with open(par_path, 'wb') as f:
                f.write(b'0A0D Mode TXT_7')

            out = os.path.join(td, 'out')
            result = run_pipeline(td, out, with_report=True)
            self.assertTrue(os.path.exists(result['output_xml']))
            self.assertTrue(os.path.exists(result['output_xrio']))
            self.assertTrue(os.path.exists(result['report_json']))


if __name__ == '__main__':
    unittest.main()
