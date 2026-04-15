import unittest

from src.prot_parser.prot_decoder import decode_prot_bytes


class ParserTests(unittest.TestCase):
    def test_decode_addresses(self):
        blob = b'ADDR 0A0B NAME Voltage VALUE 220 0C0D Current 5'
        params = decode_prot_bytes(blob, 'PROT1')
        addrs = {p.address for p in params}
        self.assertIn('0A0B', addrs)
        self.assertIn('0C0D', addrs)


if __name__ == '__main__':
    unittest.main()
