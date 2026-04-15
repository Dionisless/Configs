import os
import tempfile
import unittest
import xml.etree.ElementTree as ET

from src.emitters.xml_emitter import emit_xml
from src.emitters.xrio_emitter import emit_xrio
from src.model.relay_model import Parameter


class EmitterTests(unittest.TestCase):
    def test_emit_xml_and_xrio(self):
        params = [
            Parameter(
                address='0A0B',
                name_short='P',
                name_full='Param',
                value_raw='1',
                value_norm='1',
                datatype='INTEGER',
                enum_map={'1': 'On'},
                source_file='PROT',
            )
        ]
        with tempfile.TemporaryDirectory() as td:
            xml_path = os.path.join(td, 'output.xml')
            xrio_path = os.path.join(td, 'output.xrio')
            emit_xml(params, xml_path)
            emit_xrio(params, xrio_path)
            self.assertEqual(ET.parse(xml_path).getroot().tag, 'DeviceData')
            self.assertEqual(ET.parse(xrio_path).getroot().tag, 'XRio')


if __name__ == '__main__':
    unittest.main()
