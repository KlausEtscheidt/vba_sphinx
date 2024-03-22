'''Test Export to Rest-Files'''

import pytest
from pyparsing import ParseResults


# pylint: disable=import-error
import vbasphinx.vba_parser.vba_grammar as vbgr
import vbasphinx.vba_parser.vba_tree_export as vbexp

def export_module_test(module, ftype, resultstr):
    '''export module node and tests the result

    module is exported to rst-file, which content is compared with resultstr
    full resultstring is composed from constant part at begin, ftype and resultstr

    Args:
        module (ParseResults): ParseResults node to export
        ftype (str): part of expected resultstr, 
                    type of sphinx directive name e.q. var for vbvar
        resultstr (str): expected content of rst file
    '''
    rstfile = './test.rst'
    with open(rstfile, 'w'  , encoding="utf-8") as outfile:
        vbexp.Directive.outfile = outfile
        vbexp.export_module(module)
    with open(rstfile, 'r'  , encoding="utf-8") as infile:
        text = infile.read()
    resultstring = f'\n   .. vba:dummy:: testmodul\n\n      .. vba:vb{ftype}:: name{resultstr}'
    assert text == resultstring

def get_dummy_module(target_list_name, parse_results):
    '''generate dummy module node (type ParseResults) for export

    ParseResults node of type 'module' is generated, with a subnode
    named target_list_name, which is filled with parse_results

    Args:
        target_list_name (str): name of module subnode, which stores parse_results
        parse_results (ParseResults): ParseResults to store in module

    Returns:
        ParseResults: dummy ParseResults node of type 'module'
    '''
    module = ParseResults()
    module['module_type'] = 'dummy'
    module['obj_name'] = 'testmodul'
    module[target_list_name] = parse_results
    return module

#######################################################
# test const exports
@pytest.fixture(params=[
    'Private Const name$ = "sd" + "er"',
    'Const name As Integer = 2 * (1 + 23)',
    ])
def get_const_data(request):
    data = [
    '$ = "sd" + "er"\n         :scope: Private\n',
    ' As Integer = 2 * (1 + 23)\n',
    ]
    return request.param, data[request.param_index]

def test_const_statement(get_const_data):
    toparse, resultstring = get_const_data
    p_res = vbgr.const_statement.parse_string(toparse)
    module = get_dummy_module('const', p_res.const)
    export_module_test(module, 'const', resultstring)

#######################################################
# test var exports
@pytest.fixture(params=[
    "Public name As Boolean",
    'Global WithEvents name As Double',
    ])
def get_var_data(request):
    data = [
    ' As Boolean\n         :scope: Public\n',
    ' As Double\n         :scope: Global\n         :withevents:\n',
    ]
    return request.param, data[request.param_index]

def test_var_statement(get_var_data):
    toparse, resultstring = get_var_data
    p_res = vbgr.var_statement.parse_string(toparse)
    module = get_dummy_module('vars', p_res.vars)
    export_module_test(module, 'var', resultstring)


#######################################################
# test method exports
@pytest.fixture(params=[
    'Function name$(i%, x As String)',
    'Private Sub name()',
    'Public Static Function name(i%, we²rt$) As Boolean',
    'Friend Function name(i%) As Boolean',
    'Function name(i%) As Boolean',
    'Static Function name(i%)',
    ])
def get_method_data(request):
    data = [
    ('func', '$(i%, x As String)\n\n         :arg % i:\n         :arg String x:\n         :returns:\n         :returntype: $\n\n'),
    ('sub', '()\n         :scope: Private\n\n\n'),
    ('func', '(i%, we²rt$) As Boolean\n         :scope: Public\n         :static:\n\n         :arg % i:\n         :arg $ we²rt:\n' 
        + '         :returns:\n         :returntype: Boolean\n\n'),
    ('func', '(i%) As Boolean\n         :scope: Friend\n\n         :arg % i:\n         :returns:\n         :returntype: Boolean\n\n'),
    ('func', '(i%) As Boolean\n\n         :arg % i:\n         :returns:\n         :returntype: Boolean\n\n'),
    ('func', '(i%)\n         :static:\n\n         :arg % i:\n\n'),

    ]
    ftype, resultstring = data[request.param_index]
    return request.param, ftype, resultstring

def test_method_statement(get_method_data):
    toparse, ftype, resultstring = get_method_data
    p_res = vbgr.method_statement.parse_string(toparse)
    module = get_dummy_module('methods', p_res.methods)
    export_module_test(module, ftype, resultstring)

if __name__ == '__main__':
    pass
