'''Test Export to Rest-Files'''

import pytest
from pyparsing import ParseResults


# pylint: disable=import-error
import vbasphinx.vba_parser.vba_grammar as vbgr
import vbasphinx.vba_parser.vba_tree_export as vbexp

def parse(grammar, txt):
    '''parse text into pyparse results'''
    p_res = grammar.parse_string(txt)
    return p_res

# 
@pytest.fixture(params=[
    'Function mname$(i%, x As String)',
    'Private Sub mname()',
    'Public Static Function mname(i%, we²rt$) As Boolean',
    'Friend Function mname(i%) As Boolean',
    'Function mname(i%) As Boolean',
    'Static Function mname(i%)',
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
    ftype, result_part2 = data[request.param_index]
    resultstring = f'\n   .. vba:dummy:: testmodul\n\n      .. vba:vb{ftype}:: mname{result_part2}'
    return request.param, resultstring

# def test_method_statement(methtxt):
def test_method_statement(get_method_data):
    '''export method statements'''
    toparse,  resultstring = get_method_data
    p_res = parse(vbgr.method_statement, toparse)
    module = ParseResults()
    module['module_type'] = 'dummy'
    module['obj_name'] = 'testmodul'
    module['methods'] = p_res.methods
    rstfile = './test.rst'
    with open(rstfile, 'w'  , encoding="utf-8") as outfile:
        vbexp.Directive.outfile = outfile
        vbexp.export_module(module)
    with open(rstfile, 'r'  , encoding="utf-8") as infile:
        text = infile.read()
    assert text == resultstring


if __name__ == '__main__':
    text = "Private Sub mname()"
    test_method_statement(text)
