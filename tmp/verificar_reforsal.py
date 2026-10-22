import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pdf_a_excel_factura as m

p = Path(r'C:\Users\Usuario\Desktop\sellos\ERROR\101-1464.pdf')
print(json.dumps(m.extraer_campos(str(p), Path('tmp'), p), default=str, indent=2))
old = json.loads(Path('tmp/reforsal_baseline.json').read_text(encoding='utf-8'))
n = 0
for path, before in old.items():
    text = m.read_pdf_text_first_page(path)
    after = m.parse_formato_san_luis(text)
    actual = json.loads(json.dumps(after, default=str))
    if before is not None:
        n += 1
        assert before == actual, path
    else:
        assert after is not None, path
        print('Nuevo:', Path(path).name, after['nro'], after['total'])
    assert m.extraer_campos(path, Path('tmp'), Path(path)) == after, ('Orden', path)
print('Facturas anteriores con resultado identico:', n)
