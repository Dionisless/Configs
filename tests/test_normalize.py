import unittest

from src.model.normalize import normalize_value, normalize_address


class NormalizeTests(unittest.TestCase):
    def test_inf(self):
        self.assertEqual(normalize_value('oo'), 'INF')

    def test_txt_enum(self):
        self.assertEqual(normalize_value('TXT_7'), '7')

    def test_address(self):
        self.assertEqual(normalize_address('0x0a0b'), '0A0B')


if __name__ == '__main__':
    unittest.main()
