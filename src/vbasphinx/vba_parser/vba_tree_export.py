'''exports parsing result tree to rst outfile'''
import os
import logging

log = logging.getLogger()

# pylint: disable=global-statement, invalid-name
# pylint: enable=global-statement, invalid-name

class Directive:

    level = 0
    outfile = None

    def __init__(self, node, name, argument, complete=False):
        self.node = node
        self.direc_name = name
        self.direc_arg = argument
        self.complete= complete

    def  rst_out(self):
        '''writes directive to rst outfile'''
        indent = ' '*3*self.level
        # append arguments, type etc to the name, so that we get the complete signature
        if self.complete:
            self.__complete_argument()
        # write the directive
        directive = f'\n{indent}.. vba:{self.direc_name}:: {self.direc_arg}\n'
        self.outfile.write(directive)
        # write options
        self.__options_rst_out()
        # write docstrings as directive content
        self.__docstring_rst_out()
        # write args
        if self.direc_name in ('vbsub', 'vbfunc'):
            self.__arglist_rst_out()

    def __arglist_rst_out(self):
        self.outfile.write('\n')
        indent = ' '*3*(self.level+1)
        for arg in self.node.param_detail:
            #if arg.vb_type_as:
            vb_type = arg.vb_type_as + arg.vb_type_char
            self.outfile.write(f'{indent}:arg {vb_type} {arg.param_name}:\n')
        # for functions: write return info
        if self.direc_name == 'vbfunc':
            vb_type = self.node.vb_type_as + self.node.vb_type_char
            if vb_type:
                self.outfile.write(f'{indent}:returns:\n')
                self.outfile.write(f'{indent}:returntype: {vb_type}\n')
        self.outfile.write('\n')

    def __docstring_rst_out(self):
        if not self.node.docstrings:
            return

        self.outfile.write('\n')
        indent = ' '*3*(self.level+1)
        for doc in self.node.docstrings:
            doc = doc.strip()
            self.outfile.write(f'{indent}{doc}\n')
        # self.outfile.write('\n')

    def __options_rst_out(self):
        indent = ' '*3*(self.level+1)
        if self.node.scope:
            self.outfile.write(f'{indent}:scope: {self.node.scope}\n')
        if self.node.withevents:
            self.outfile.write(f'{indent}:withevents:\n')
        if self.node.Static:
            self.outfile.write(f'{indent}:static:\n')

    def __complete_argument(self):
        '''adds argumetlist and type to the name'''
        if self.node.vb_type_char: # add char for type like in Dim i%
            self.direc_arg += self.node.vb_type_char
        # add parameterlist for sub and function
        if self.direc_name in ('vbsub', 'vbfunc'):
            self.direc_arg += self.node.method_params
        # add 'As type' if exists
        if self.node.vb_type_as:
            self.direc_arg += ' As ' + self.node.vb_type_as
        # add value for const
        if self.direc_name == 'vbconst':
            self.direc_arg  += ' = ' + self.node.value

def export_module(module):
    '''writes directive for module and all its entities to rst outfile'''
    log.info('\n%s : %s', module['module_type'], module['obj_name'])
    Directive.level = 1
    direc = Directive(module, module['module_type'], module['obj_name'])
    direc.rst_out()
    Directive.level = 2

    log.info('       %d Konstanten gefunden:', len(module.const))
    for const in module.const:
        log.info('      %s %s %s = %s', const.scope, const.obj_name, const.vb_type, const.value)
        direc = Directive(const, 'vbconst', const.obj_name, True)
        direc.rst_out()

    log.info('       %d glob. Variable gefunden:', len(module.vars))
    for var in module.vars:
        log.info('       %s %s %s', var.scope, var.obj_name, var.vb_type)
        direc = Directive(var, 'vbvar', var.obj_name, True)
        direc.rst_out()

    log.info('       %d properties gefunden:', len(module.props))
    # we have multiple statements (Let, Get, Set) for properties,
    # but we want to export only one property node
    # so first, we collect them into a dict with prop_name as key
    # the value is a tuple (docs, sub-dict)
    # with sub-dict as dictionary of all statements found for this prop_name
    all_props = {}
    for prop in module.props:
        if not prop.obj_name in all_props.keys():
            # first statement for property named prop_name
            # we generate dict-entry under key of prop_type
            one_prop = {prop.prop_type: prop}
            # insert first statement into dict
            all_props[prop.obj_name] = (prop.docstrings, one_prop)
        else:
            # second statement for property named prop_name
            # we collect docstrings from every statement for this prop_name
            docs = all_props[prop.obj_name][0]
            docs += prop.docstrings
            # append statement under key of prop_type
            all_props[prop.obj_name][1][prop.prop_type] = prop

    for statement in all_props.values():
        docs = statement[0]
        if 'Get' in statement[1].keys():
            # We prefer Get, because we have the property type (return type) here
            prop = statement[1]['Get']
        else:
            # export random one (Let or Set)
            prop = next(iter(statement[1].values()))

        prop['docstrings'] = docs
        log.info('      %s %s %s %s (%s) %s ', prop.scope, prop.method_type,
                prop.prop_type, prop.obj_name, prop.prop_params, prop.vb_type)
        direc = Directive(prop, 'vbprop', prop.obj_name, True)
        direc.rst_out()

    log.info('       %d Methoden gefunden:', len(module.methods))
    for meth in module.methods:
        log.info('      %s %s %s (%s) %s', meth.scope, meth.method_type,
                meth.obj_name, meth.method_params, meth.vb_type)
        if  meth.method_type == 'Sub':
            direc = Directive(meth, 'vbsub', meth.obj_name, True)
        else:
            direc = Directive(meth, 'vbfunc', meth.obj_name, True)
        direc.rst_out()


def export_rst(topnode, exportdir, fullpath):
    '''exports parsing result tree to rst outfile

    Args:
        topnode (ParseResults): top node of the parsed structure
        exportdir (str): dir to export to
        fullpath (str): path to VBA-source (Exel file, etc)
    '''

    _, name_ext = os.path.split(fullpath)
    filename, _ = os.path.splitext(name_ext)

    # global outfile, level # pylint: disable=global-statement
    outpath = os.path.join(exportdir, filename + '.rst')

    log.info('   %d Module gefunden:', len(topnode.vbamodules))

    with open(outpath , 'w'  , encoding="utf-8") as outfile:
        outfile.write(f'{filename}\n')
        line = '='*len(filename)
        outfile.write(line+'\n')
        Directive.outfile = outfile
        Directive.level = 0
        direc = Directive(topnode, 'module', filename)
        direc.rst_out()
        for module in topnode.vbamodules:
            export_module(module)

    log.info('Rest File %s geschrieben', outpath)
