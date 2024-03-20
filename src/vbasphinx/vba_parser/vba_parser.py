'''parse files with vba source code and store them in rst-format'''

import os
import logging

# import tomlkit.toml_file
import pyparsing as pp

from vbasphinx.vba_utils.config_reader_toml import ConfigReader
from vbasphinx.vba_utils.vba_logging import setup_logger
import vbasphinx.vba_parser.vba_grammar as vbgr
from vbasphinx.vba_parser.vba_tree_export import export_rst

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
    '''parses one vba file

    Args:
        fpath (str): file to parse
    '''

    _, name_ext = os.path.split(fpath)
    # name, _ = os.path.splitext(name_ext)
    log.error('\n\nstart to parse: %s\n',name_ext)
    vbgr.currently_parsed_file = fpath

    try:
        text = strip_comments(fpath)
        erg = vbgr.vbagramm.parse_string(text)
    except pp.ParseException as err:
        log.error('\n\nFehler beim parsen von:\n%s\n', fpath)
        log.error('message: %s\n', err.msg,)
        log.error('found %s \n',err.line)
        log.error('in line no: %d\n', err.lineno)
        return None
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

    for infile in files2process:
        tree = parse_file(infile)
        if tree:
            export_rst(tree, outdir, infile)

if __name__ == '__main__':
    run()
