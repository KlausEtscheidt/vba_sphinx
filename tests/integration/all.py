'''testet Gesamtablauf

- Auslesen der vba sourcen aus Excel oder Access
- Parsen der files und Ausgabe als rest-Dateien
- Erzeugen von html mit Sphinx
'''

from  build_sphinx_doku import buildit
from  vbasphinx.vba_reader import AccessReader, ExcelReader
from  vbasphinx.vba_parser import vba_parser

AccessReader.run()
ExcelReader.run()

vba_parser.run()
buildit('html', './tests/integration/sphinx')
