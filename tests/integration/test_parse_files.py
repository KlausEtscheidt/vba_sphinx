'''parses all files from vba_parser.toml

vergleicht das Ergebnis mit den Werten aus vba_vorgabe.json'''

import vba_parser_summary as vbchk
from vbasphinx.vba_utils.config_reader_toml import ConfigReader
from  vbasphinx.vba_parser import vba_parser
# sys.path.append('./tests/integration')

cfg = ConfigReader('vba_parser.toml')
files2process = cfg.getfiles('filelist')
outdir = cfg.getdir('outdir')

for infile in files2process:
    tree = vba_parser.parse_file(infile)
    if tree:
        vba_parser.export_rst(tree, outdir, infile)
        vbchk.export_summary(tree, infile)

try:
    messages = vbchk.check_summary()
    # vbchk.write_summary()
except vbchk.VBAParserCheckExc as err:
    print(err)
else:
    if messages:
        for msg in messages:
            print(msg)
    else:
        print('='*30,' alle ok ','='*30)
