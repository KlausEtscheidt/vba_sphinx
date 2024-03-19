'''parse files with vba source code and store them in rst-format'''

import os
import sys
import logging
import time

# import tomlkit.toml_file
import pyparsing as pp

from vbasphinx.vba_utils.config_reader_toml import ConfigReader
from vbasphinx.vba_utils.vba_logging import setup_logger
import vbasphinx.vba_parser.vba_grammar as vbgr
from vbasphinx.vba_parser.vba_tree_export import export_rst
import vbasphinx.vba_parser.vba_parser_summary as vbchk

# Basis-Verzeichnis
# my_file = os.path.realpath(__file__) # Welcher File wird gerade durchlaufen
# my_dir = os.path.dirname(my_file)
# parent_dir = os.path.join(my_dir, '..')
# sys.path.insert(0, parent_dir)


log = logging.getLogger()

class VBAParserExc(Exception):
    '''class for exceptions'''

def strip_comments(fpath):
    '''ommit comment-lines from output but keep docstrings

    append separator line after each block of docstrings
    '''
    with open(fpath,"r",encoding="utf-8") as fp:
        text = fp.readlines()
    outlines = ''
    for line in text:
        line = line.strip()
        if not line:
            continue # ignore empty lines
        elif  len(line)==1:
            if not line == "'":
                # normally, this shouldn't be in vba source
                outlines += line +'\n'
        else:
            if line[0] == "'":
                if line[1] == "%":
                    # end of docstring block
                    outlines += '#######' +'\n'
                elif line[1] == "!":
                    # docstring
                    outlines += line +'\n'
            else:
                #regular line
                outlines += line +'\n'
    return outlines

def parse_file(fpath):
    '''parses one vba

    Args:
        fpath (str): file to parse
    '''
    try:
        text = strip_comments(fpath)
        erg = vbgr.vbagramm.parse_string(text)
    except pp.ParseException as err:
        log.error('\n\nFehler beim parsen von:\n%s\n', fpath)
        log.error('message: %s\n', err.msg,)
        log.error('found %s \n',err.line)
        log.error('in line no: %d\n', err.lineno)
        return None
        # raise err
    return erg

def run():
    '''parse files with vba source code and store them in rst-format'''

    #setup logging
    setup_logger('./vba_parser.log')
    # log.setLevel(logging.DEBUG)
    log.setLevel(logging.INFO)
    # log.setLevel(logging.ERROR)

    cfg = ConfigReader('vba_parser.toml')
    files2process = cfg.getfiles('filelist')
    outdir = cfg.getdir('outdir')
    start = time.time()
    for infile in files2process:
        _, name_ext = os.path.split(infile)
        name, _ = os.path.splitext(name_ext)
        log.error('\n\nstart to parse: %s\n',name_ext)

        vbgr.currently_parsed_file = infile
        tree = parse_file(infile)
        if tree:
            export_rst(tree, outdir, name)
            vbchk.export_summary(tree, name)
    print (time.time() -start)
    try:
        vbchk.check_summary()
        # vbchk.write_summary()
    except vbchk.VBAParserCheckExc as err:
        print(err)

if __name__ == '__main__':
    run()
