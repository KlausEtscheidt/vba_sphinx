'''test modul'''
from  vbasphinx.xl_reader.xl_codereader import XLReader
from  vbasphinx.vba_parser import vba_parser
from  build_sphinx_doku import buildit

# XLReader.run()
vba_parser.run()
# buildit('html', './sphynx_vba_domain')

# def runtest():
#     '''only for debugging'''
#     setup_logger('./vba_parser.log')
#     log.setLevel(logging.DEBUG)
#     vbgr.currently_parsed_file = 'xxx'
#     erg = vbgr.var_decl.run_tests("""\
# '! speichert Status des Formulars
# Public ok_pressed As Boolean
#     """)
