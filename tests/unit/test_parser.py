import pytest
from pyparsing import ParseResults


# pylint: disable=import-error
# from  vbasphinx.vba_parser import vba_parser
import vbasphinx.vba_parser.vba_grammar as vbgr

def parse(grammar, txt):
    p_res = grammar.parse_string(txt)
    try:
        p_res = grammar.parse_string(txt)
    except Exception as err:
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
    "Public varname As Boolean, i%",
    "Public varname%",
    "Dim varname$",
    "Dim WithEvents varname$",
    "Global WithEvents varname As Double",
    #"Public Const WithEvents varname$",
])

def test_var(vartxt):
    '''var statements'''
    p_res = parse(vbgr.module_entity, vartxt)
    # p_res = parse(vbgr.var_statement, vartxt)
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

def test_props(proptxt):
    '''property statements'''
    p_res = parse(vbgr.prop, proptxt + ' lala\n End Property')
    assert isinstance(p_res['props'][0], ParseResults)
    prop = p_res['props'][0]
    assert prop.obj_name == 'pname'
    assert prop.prop_type in ('Get', 'Let', 'Set')


#######################################################
# test method argument lists
# [ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]
@pytest.fixture(params=[
    '(xyz() As String = "asd", myint As Integer)',
    '(Optional xyz%() = "asd")',
    '(myint%)',
    '(ByVal ParamArray myint%)',
    ])
def get_method_param_data(request):
    # for each test data is a list of tuples(name,type_char,type_as,default)
    # one tuple for each arg in method argument lists
    data = [
    [('xyz', '', 'String', '"asd"'),('myint', '', 'Integer', '')],
    [('xyz', '%', '', '"asd"')],
    [('myint', '%', '', '')],
    [('myint', '%', '', '')],
    ]
    return request.param, data[request.param_index]

def test_method_params(get_method_param_data):
    toparse, resultdata = get_method_param_data
    # p_res = parse(vbgr.prop_params, toparse)
    p_res = parse(vbgr.method_params, toparse)
    assert 'param_detail' in p_res.keys()
    assert len(p_res['param_detail']) == len(resultdata)
    for i, arg_desc in enumerate(resultdata):
        p_res_arg = p_res['param_detail'][i]
        assert arg_desc[0] == p_res_arg.param_name
        if 'vb_type_char' in p_res_arg.keys():
            assert arg_desc[1] == p_res_arg.vb_type_char
        if 'vb_type_as' in p_res_arg.keys():
            assert arg_desc[2] == p_res_arg.vb_type_as
        if 'value' in p_res_arg.keys():
            assert arg_desc[3] == p_res_arg.value.strip()


#####################################
# test method statements
# [ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ]
# [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
@pytest.mark.parametrize("methtxt", [
    "Function mname$(i%, x() As String)",
    "Private Sub mname()",
    "Sub mname()",
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
    pass
