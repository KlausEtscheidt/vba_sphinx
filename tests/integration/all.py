'''test modul'''
from  vbasphinx.vba_reader import AccessReader, ExcelReader
from  vbasphinx.vba_parser import vba_parser
from  build_sphinx_doku import buildit

#AccessReader.run()
#ExcelReader.run()

vba_parser.run()
#buildit('html', './tests/integration/sphinx')
