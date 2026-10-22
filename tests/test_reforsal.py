import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pdf_a_excel_factura as parser


FIXTURES = Path(__file__).with_name('fixtures')
TEXT = (FIXTURES / 'reforsal_nuevo.txt').read_text(encoding='utf-8')
EXPECTED = {
    'nro': '101-1464',
    'razon': 'ENTE DE RECUPERACION DE FONDOS PARA EL FORTALECIMIENTO DEL SISTEMA DE SALUD',
    'cuit': '30-71861570-0',
    'iibb': '',
    'domicilio': 'Rafael Cubillos 2056 Piso 3 Dpto. 302',
    'localidad': 'Godoy Cruz',
    'provincia': 'Mendoza',
    'fecha': datetime(2026, 10, 16),
    'periodo': datetime(2025, 9, 3),
    'total': 5877131.10,
}


class ReforsalTests(unittest.TestCase):
    def test_debug_real(self):
        self.assertEqual(parser.parse_formato_san_luis(TEXT), EXPECTED)

    def test_espacios_saltos_acentos_y_ocr(self):
        variants = [
            TEXT.replace(' ', '   '),
            TEXT.replace(' ', '\n'),
            TEXT.replace('Razón', 'Razon').replace('Periodo', 'Período').replace('Facturacion', 'Facturación'),
            TEXT.replace('REFORSAL', 'REF0R5AL').replace('Razón', 'Raz0n').replace('Facturacion', 'Facturaci0n'),
            TEXT.replace('N° 00101', 'Nº 00101'),
            TEXT.replace('N° 00101', 'N0: 00101'),
            TEXT.replace('N° 00101', 'Nro.: 00101'),
            TEXT.replace('Total AR', 'Total ARS $'),
            TEXT.replace('Total AR', 'Total $'),
            TEXT.replace('30718615700', '30 - 71861570 - 0'),
            TEXT.replace('5.877.131,10', '5877131,10'),
            TEXT.replace('03-09-2025', '3 / 9 / 2025'),
            TEXT.replace('302- Godoy', '302\nGodoy'),
            TEXT.replace('Rafael Cubillos', 'Domicilio:\nRafael Cubillos'),
        ]
        for variant in variants:
            with self.subTest(text=variant[:100]):
                self.assertEqual(parser.parse_formato_san_luis(variant), EXPECTED)
        for variant in (TEXT.lower(), TEXT.upper()):
            result = parser.parse_formato_san_luis(variant)
            self.assertIsNotNone(result)
            for key, value in EXPECTED.items():
                self.assertEqual(result[key].upper() if isinstance(value, str) else result[key],
                                 value.upper() if isinstance(value, str) else value)

    def test_no_tomar_cuit_del_cliente(self):
        self.assertIsNone(parser.parse_formato_san_luis(TEXT.replace('CUIT: 30718615700', '')))

    def test_delimitacion_del_cliente_obligatoria(self):
        text = TEXT.replace('DATOS DEL CLIENTE', '').replace('Cliente:', 'Comprador:')
        self.assertIsNone(parser.parse_formato_san_luis(text))

    def test_sin_cuit_cliente_no_cambia_emisor(self):
        self.assertEqual(parser.parse_formato_san_luis(TEXT.replace('30679065382', '20999999999')), EXPECTED)

    def test_total_prevalece_sobre_subtotal_e_items(self):
        text = TEXT.replace('Subtotal: 5.877.131,10', 'Subtotal: 9.999.999,99')
        text = text.replace('5.779.749,66', '8.888.888,88')
        self.assertEqual(parser.parse_formato_san_luis(text), EXPECTED)

    def test_sin_total_no_usar_importes_ni_cae(self):
        self.assertIsNone(parser.parse_formato_san_luis(TEXT.replace('Total AR 5.877.131,10', '')))

    def test_otra_marca_no_es_reforsal(self):
        self.assertIsNone(parser.parse_formato_san_luis(TEXT.replace('REFORSAL', 'OTRA MARCA')))

    def test_fechas_invalidas(self):
        for old, new in [('16/10/2026', '32/10/2026'), ('03-09-2025', '40-09-2025')]:
            self.assertIsNone(parser.parse_formato_san_luis(TEXT.replace(old, new)))

    def test_original_y_duplicado_un_solo_resultado(self):
        self.assertEqual(parser.parse_formato_san_luis(TEXT + '\f' + TEXT.replace('ORIGINAL', 'DUPLICADO')), EXPECTED)

    def test_extraer_campos_usa_parser_existente(self):
        with patch.object(parser, 'read_pdf_text_first_page', return_value=TEXT):
            self.assertEqual(parser.extraer_campos('factura.pdf', Path('.'), Path('factura.pdf')), EXPECTED)

    def test_formato_anterior_real(self):
        text = (FIXTURES / 'reforsal_anterior.txt').read_text(encoding='utf-8')
        result = parser.parse_formato_san_luis(text)
        self.assertEqual(result['nro'], '102-956')
        self.assertEqual(result['cuit'], '30-71861570-0')
        self.assertEqual(result['fecha'], datetime(2026, 4, 6))
        self.assertEqual(result['periodo'], datetime(2025, 5, 27))
        self.assertEqual(result['total'], 6851677.79)
        with patch.object(parser, 'read_pdf_text_first_page', return_value=text):
            self.assertEqual(parser.extraer_campos('anterior.pdf', Path('.'), Path('anterior.pdf')), result)


if __name__ == '__main__':
    unittest.main()
