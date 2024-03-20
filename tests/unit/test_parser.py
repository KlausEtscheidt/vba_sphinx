import pytest
from pyparsing import ParseResults


# pylint: disable=import-error
# from  vbasphinx.vba_parser import vba_parser
import vbasphinx.vba_parser.vba_grammar as vbgr

def parse(grammar, txt):
    try:
        p_res = grammar.parse_string(txt)
    except:
        p_res = None
    assert isinstance(p_res, ParseResults)
    return p_res

# Public [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
# [ , [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]] . . .
@pytest.mark.parametrize("vartxt", [
    "Public varname As Boolean",
    "Public varname%",
    "Dim varname$",
    "Dim WithEvents varname$",
    "Global WithEvents varname As Double",
    #"Public Const WithEvents varname$",
])

def test_var(vartxt):
    '''var statements'''

    p_res = parse(vbgr.module_entity, vartxt)
    assert isinstance(p_res['vars'][0], ParseResults)
    var = p_res['vars'][0]
    assert var.obj_name == 'varname'
    # assert var.scope == 'Public'
    # assert var.vb_type_as == 'Boolean'

def test_mod():
    '''vbamodule statements'''
    txt = """\
========================================================
vbmodule: Modul1
========================================================
'! comment\n
Public ok_pressed As Boolean
<EndofFile>
    """
    p_res = parse(vbgr.vbamodule, txt)
    assert p_res[0].obj_name == 'Modul1'
    assert p_res[0].module_type == 'vbmodule'

if __name__ == '__main__':
    # vbgr.currently_parsed_file = 'xxx'
    # log.setLevel(logging.DEBUG)
    # setup_logger('./vba_parser.log')
    test_var()
