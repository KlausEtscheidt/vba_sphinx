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

#[ Public | Private ] Const constname [ As type ] = expression
@pytest.mark.parametrize("consttxt", [
    'Private Const cname$ = "sd" + "er"',
    "Public Const cname As Integer = 23",
    "Const cname As Integer = 2 * (1 + 23)",
])

def test_const(consttxt):
    '''const statements'''
    p_res = parse(vbgr.module_entity, consttxt)
    assert isinstance(p_res['const'][0], ParseResults)
    const = p_res['const'][0]
    assert const.obj_name == 'cname'

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

# [ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]

# [ Public | Private | Friend ] [ Static ] Property Get name [ (arglist) ] [ As type ]
# [ Public | Private | Friend ] [ Static ] Property Set name ( [ arglist ], reference )
# [ Public | Private | Friend ] [ Static ] Property Let name ( [ arglist ], value )
@pytest.mark.parametrize("proptxt", [
    "Private Property Get pname As Integer",
    'Private Property Get pname (xyz() As String = "asd") As Integer',
    "Private Static Property Get pname",
    "Public Property Set pname (i%, b As Int)",
    "Property Let pname(s)",
])

def test_params(proptxt):
    '''property statements'''
    p_res = parse(vbgr.prop, proptxt + ' lala\n End Property')
    assert isinstance(p_res['props'][0], ParseResults)
    prop = p_res['props'][0]
    assert prop.obj_name == 'pname'
    assert prop.prop_type in ('Get', 'Let', 'Set')

# [ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ]
# [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
@pytest.mark.parametrize("methtxt", [
    "Function mname$(i%, x As String)",
    "Private Sub mname()",
    "Public Static Function mname(i%, we²rt$) As Boolean",
    "Friend Function mname(i%) As Boolean",
    "Function mname(i%) As Boolean",
    "Static Function mname(i%)",
])

def test_method_statement(methtxt):
    '''method statements'''
    p_res = parse(vbgr.method_statement, methtxt)
    assert isinstance(p_res['methods'][0], ParseResults)
    p_res = p_res['methods'][0]
    assert p_res.obj_name == 'mname'
    assert p_res.method_type in ('Sub','Function')

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
