'''defines the parsing grammar vor VBA sources'''
import logging
import pyparsing as pp

# pylint: disable=global-statement, invalid-name
log = logging.getLogger()

# pylint: disable=unnecessary-pass

currently_parsed_file = '' # pylint: disable=invalid-name

# chars valid for types in vb (like dim i% for integer)
TYPEDEF_CHAR = '%&^@!#$'

# for some of the next expressions, we allow more chars than VBA does
# to make sure that we don't miss a token
ALL_CHAR = pp.alphas + pp.alphas8bit + pp.punc8bit + pp.printables

LPAR, RPAR, EQUAL, COLON = map (pp.Suppress, '()=:')

# Keywords
PRIV, PUBL, FRIEND, STATIC = map (pp.Keyword, ['Private', 'Public', 'Friend', 'Static'])
DIM, GLOB = map (pp.Keyword, ['Dim', 'Global'])
CONST, SUB, FUNC, EOF = map (pp.Keyword, ['Const', 'Sub', 'Function', '<EndofFile>'])
PROP, GET, LET, SET = map (pp.Keyword, ['Property', 'Get', 'Let', 'Set'])
BYVAL, BYREF, OPTIONAL, PARRAY = map (pp.Keyword, ['ByVal', 'ByRef', 'Optional', 'ParamArray'])
VBMODULE, VBCLASS, VBXLOBJ, VBFORM = \
                           map (pp.Keyword, ['vbmodule', 'vbclass', 'vb_office_obj', 'vbform'])
AS, NEW, WITH, END = map (pp.Keyword, ['As', 'New', 'WithEvents', 'End'])

# start of new module
DIVIDER = pp.Suppress(pp.Word('=', min=40))

# docstrings
docstring = pp.Suppress(pp.Literal("'!")) + pp.rest_of_line

# used in variable definitions (Dim, Public, etc) for var names
# vbvarname =  pp.Word(ALL_CHAR, exclude_chars=TYPEDEF_CHAR)

# type definitions
# --------------------------------------------------------------------------
# used in 'As'-type definitions for typename
vbtype = pp.Word(ALL_CHAR + '.',exclude_chars="(),'\n")
type_as = pp.Suppress(AS + pp.Opt(NEW)) + vbtype('vb_type_as')
# typedef per special char like "dim x%""
type_char = pp.Char(TYPEDEF_CHAR)('vb_type_char')

# name of const, var, property, sub, function with optional char for type
vb_typedname = pp.Word(ALL_CHAR, exclude_chars=TYPEDEF_CHAR+"(),'\n")


# declaration of a const value (only parsed if outside of sub or function)
# ------------------------------------------------------------------------
#[ Public | Private ] Const constname [ As type ] = expression
const_value = pp.Word(pp.printables + ' "', exclude_chars="'")
const_scope = pp.Opt(PRIV | PUBL | GLOB)
const_statement = pp.Group(const_scope('scope') + pp.Suppress(CONST)\
                      + vb_typedname('obj_name') + pp.Opt(type_char) +  pp.Opt(type_as)\
                      + EQUAL + const_value('value'))('const*')

# declaration of a variable (only parsed if outside of sub or function)
# ------------------------------------------------------------------------
# Public [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
# [ , [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]] . . .
var_scope = PRIV | PUBL | GLOB | DIM
var_subscript = LPAR + pp.SkipTo(')', include=False) + RPAR
not_method = ~ ( SUB | FUNC | STATIC | PROP)
var_statement = pp.Group(var_scope('scope') + not_method\
                + pp.Opt(WITH)('withevents') + vb_typedname('obj_name') + pp.Opt(type_char)\
                + pp.Opt(var_subscript)('subscripts') + pp.Opt(type_as))('vars*')

# properties
# ------------------------------------------------------------------------
# [ Public | Private | Friend ] [ Static ] Property Get name [ (arglist) ] [ As type ]
# [ Public | Private | Friend ] [ Static ] Property Set name ( [ arglist ], reference )
# [ Public | Private | Friend ] [ Static ] Property Let name ( [ arglist ], value )
method_scope = pp.Optional(PRIV | PUBL | FRIEND)
prop_name = vb_typedname('obj_name')
prop_begin = method_scope('scope') + pp.Opt(STATIC)('Static')\
                            + PROP('method_type')
prop_params = pp.originalTextFor(pp.nestedExpr())('prop_params')
# prop_params = LPAR + pp.SkipTo(')', include=False)('prop_params') + RPAR

prop_get = prop_begin + GET('prop_type') + prop_name + pp.Opt (prop_params) + pp.Opt(type_as)
prop_let = prop_begin + (SET | LET)('prop_type') + prop_name + prop_params
prop_statement = pp.Group(prop_get | prop_let)('props*')
prop_end = pp.Suppress((END + PROP))
prop_content = pp.SkipTo(prop_end, include=True)('prop_body')
prop = prop_statement + prop_content

# functions and subs (procedures)
# ------------------------------------------------------------------------
# [ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ]
# [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
method_types = SUB | FUNC
method_params = pp.Opt(pp.originalTextFor(pp.nestedExpr())('method_params'))
method_statement = pp.Group(method_scope('scope')  + pp.Opt(STATIC)('Static')\
            + method_types('method_type') \
            + vb_typedname('obj_name') + pp.Opt(type_char)\
            + method_params + pp.Opt(type_as))('methods*')
method_end = pp.Suppress((END + method_types))
method_content = pp.SkipTo(method_end, include=True)('method_body')
method = method_statement + method_content

# -------------------------------------------------------------------------------------
# vba module
# -------------------------------------------------------------------------------------

# module header
# -------------
# description of the module
mod_types = VBMODULE | VBCLASS | VBXLOBJ | VBFORM
module_header_content = mod_types('module_type') + COLON + pp.Word(ALL_CHAR)('obj_name')
module_header = DIVIDER + module_header_content + DIVIDER

# module end
# ----------
# after all items of interest, we skip over anything else to the next module or end of file
end_targets = DIVIDER | EOF
module_end = pp.SkipTo(end_targets, include=False)('module_end')

# module content
# --------------
# here we have to list all items, which should be in the result tree
target_entities = const_statement | var_statement | method | prop
module_entity = pp.SkipTo(target_entities, include=True, ignore=docstring)
module_content = pp.ZeroOrMore(module_entity, stop_on=end_targets)

# complete module
vbamodule = pp.Group(module_header + module_content + module_end)

# -------------------------------------------------------------------------------------
# complete file
# -------------------------------------------------------------------------------------
vbagramm = pp.ZeroOrMore(vbamodule('vbamodules*')) + EOF

def get_method_arguments(_text, _loc, toks):
    '''parse argument list into single arguments'''
    if not toks.method_params:
        return
    # [ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]
    vbparam = pp.Group(pp.Opt(OPTIONAL)('opt') + pp.Opt(BYREF|BYVAL)('by') +pp.Opt(PARRAY)\
                    + vb_typedname('param_name') + pp.Opt(type_char) + pp.Opt(type_as)
                    )
    vbparameter = LPAR + pp.Opt(pp.delimited_list(vbparam, delim=',')) + RPAR
    erg = vbparameter.parse_string(toks.method_params)
    # insert node in result tree if not empty ()
    if erg:
        toks['param_detail'] = erg
        toks.append(erg)
    pass

def get_docs_before_item(text, loc, toks):
    '''inserts docstrings, written before items into result tree.'''
    docs = []

    # first we look in the line of the statement itself (docstring right beside)
    lines_after = text[loc:].split('\n')
    current_line_parts = lines_after[0].split("'!")
    if len(current_line_parts) > 1:
        docs.append(current_line_parts[1])

    # now we look before the statement and add docstrings found here
    lines = text[0:loc].split('\n')
    # search start of docstring
    isdoc = True
    max_line = len(lines)
    lnr = max_line-2
    while isdoc and lnr>0:
        line = lines[lnr].strip()
        isdoc = line[0:2] == "'!"
        if isdoc:
            docs.insert(0,line[2:])
        lnr -= 1
    # insert node in result tree
    toks[0]['docstrings'] = docs
    toks[0].append(docs)

def get_docs_after_modheader(text, loc, toks):
    '''inserts docstrings for modules into result tree.'''
    docs = []
    lines = text[loc:].split('\n')
    # search end of docstring
    for line in lines[3:]:
        isdoc = line[0:2] == "'!"
        if isdoc:
            docs.append(line[2:])
        else:
            break
    toks['docstrings']=docs
    toks.append(docs)

module_header.set_parse_action(get_docs_after_modheader)
method_statement.set_parse_action(get_docs_before_item)
prop_statement.set_parse_action(get_docs_before_item)
const_statement.set_parse_action(get_docs_before_item)
var_statement.set_parse_action(get_docs_before_item)

method_params.set_parse_action(get_method_arguments)

pp.autoname_elements()


#-----------------------------------------------------------------------
# debug
#-----------------------------------------------------------------------
def print_search(s, loc, expr, _cache):
    '''logs message when starts seraching for expression (for debugging)'''
    log.debug('suche %s bei %d', expr.customName, loc)
    log.debug('suche ab >>>%s ....\n', s[loc:loc+200])
    pass

def print_found(text, loc_start, loc_end, expr, _toks, _cache):
    '''logs message when found expression (for debugging)'''
    log.info('%s gefunden:', expr.customName)
    log.debug('von %d bis %d:\n>>>%s<<<\n', loc_start, loc_end, text[loc_start:loc_end])
    pass

def nop_err(_text, _loc, _expr, _exc, _cache):
    '''does nothing'''
    pass

def print_all_err(_text, _loc, expr, exc, _cache):
    '''logs message when expression was not found (for debugging)'''
    log.error('\n\n %s\n', currently_parsed_file)
    log.error('Fehler: %s\nin zeile %d gefunden:\n%s', exc.msg, exc.lineno, exc.line)
    log.error('nicht gefunden: %s\n', expr.customName)
    # loc_min = max(0, loc-100)
    # loc_max = min(len(text)-1, loc+100)
    # log.error('Text nach errloc %d\n>>>\n%s\n<<<', loc, text[loc:loc_max])
    # log.error('Text vor errloc %d\n>>>\n%s\n<<<', loc, text[loc_min:loc])
    pass

def print_finished(_text, _loc_start, _loc_end, _expr, _toks, _cache):
    '''prints message when file is parsed completely'''
    log.error('parsing of %s \nfinished OK !!!\n', currently_parsed_file)

vbagramm.set_debug_actions(print_search,print_finished,print_all_err)
# vbamodule.set_debug_actions(print_search,print_found,print_all_err)

#module_header.set_debug_actions(print_search,print_found,print_all_err)
# module_content.set_debug_actions(print_search,print_found,print_all_err)
# module_entity.set_debug_actions(print_search,print_found,print_all_err)
# module_end.set_debug_actions(print_search,print_found,print_all_err)

# const_statement.set_debug_actions(print_search,print_found,print_all_err)
# var_statement.set_debug_actions(print_search,print_found,print_all_err)
#method_statement.set_debug_actions(print_search,print_found,print_all_err)

# skip targets
# DIVIDER.set_debug_actions(print_search,print_found,print_all_err)
# EOF.set_debug_actions(print_search,print_found,print_all_err)
