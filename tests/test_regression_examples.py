import os
import unittest

from src.validators.crosscheck import compare_xml_xrio


class RegressionExamplesTests(unittest.TestCase):
    def test_existing_examples_are_comparable(self):
        xml_path = os.path.join('Примеры', 'QC2_АУВ_25_6MD (25-219).xml')
        xrio_path = os.path.join('Примеры', 'QC2_АУВ_25_6MD (25-219).xrio')
        if not (os.path.exists(xml_path) and os.path.exists(xrio_path)):
            self.skipTest('example files are missing')

        result = compare_xml_xrio(xml_path, xrio_path)
        self.assertIsInstance(result.summary(), dict)


if __name__ == '__main__':
    unittest.main()
