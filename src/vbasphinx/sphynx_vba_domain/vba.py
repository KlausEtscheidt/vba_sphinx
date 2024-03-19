"""The VBA (Visual Basic for Applications) domain."""

from __future__ import annotations

import contextlib
from typing import TYPE_CHECKING, Any, cast

from collections import defaultdict

from docutils import nodes
from docutils.parsers.rst import directives

from sphinx import addnodes
from sphinx.directives import ObjectDescription
from sphinx.domains import Domain, ObjType, Index
from sphinx.domains.python import _pseudo_parse_arglist
from sphinx.locale import _, __
from sphinx.roles import XRefRole
from sphinx.util import logging
from sphinx.util.docfields import Field, GroupedField, TypedField
from sphinx.util.docutils import SphinxDirective
from sphinx.util.nodes import make_id, make_refnode, nested_parse_with_titles

if TYPE_CHECKING:
    from collections.abc import Iterator

    from docutils.nodes import Element, Node

    from sphinx.addnodes import desc_signature, pending_xref
    from sphinx.application import Sphinx
    from sphinx.builders import Builder
    from sphinx.environment import BuildEnvironment
    from sphinx.util.typing import OptionSpec

logger = logging.getLogger(__name__)


class VBAObject(ObjectDescription[tuple[str, str]]):
    """
    Description of a VBA object.
    """
    #: If set to ``True`` this object is callable and a `desc_parameterlist` is
    #: added
    has_arguments = False

    #: If ``allow_nesting`` is ``True``, the object prefixes will be accumulated
    #: based on directive nesting
    allow_nesting = False

    object_type_prefix = ''

    option_spec: OptionSpec = {
        'no-index': directives.flag,
        'no-index-entry': directives.flag,
        'no-contents-entry': directives.flag,
        'no-typesetting': directives.flag,
        'noindex': directives.flag,
        'noindexentry': directives.flag,
        'nocontentsentry': directives.flag,
        'single-line-parameter-list': directives.flag,
        'static': directives.flag,
        'scope': directives.unchanged,
        # 'vba_type': directives.unchanged,
        'withevents': directives.flag,
    }

    def get_sig_keyword_entry(self, name: str) -> list[Node]:
        '''convenience function returns desc_sig_keyword + desc_sig_space'''
        return [addnodes.desc_sig_keyword(name, name), addnodes.desc_sig_space()]

    def get_display_prefix(self) -> list[Node]:
        #: what is displayed right before the documentation entry
        nodes = []

        if 'scope' in self.options:
            val = self.options['scope']
            nodes += self.get_sig_keyword_entry(val)
        if 'static' in self.options:
            nodes += self.get_sig_keyword_entry('Static')
        if 'withevents' in self.options:
            nodes += self.get_sig_keyword_entry('WithEvents')
        nodes += self.get_sig_keyword_entry(self.object_type_prefix)
        return nodes

    def parse_signature(self, sig: str)  -> tuple[str, str]:
        sig = sig.strip()
        restofline = ''
        arglist = ''
        name = sig

        # sub and function have at least an empty parameter list
        if self.objtype in ('vbsub', 'vbfunc'):
            pos = sig.find('(')
            if pos > 0:
                name = sig[:pos]
                pos2 = sig.find(')')
                arglist = sig[pos+1:pos2]
                restofline = sig[pos2+1:]

        # for const and properties name can be followed by whitespace and 'As type'
        if self.objtype in ('vbprop', 'vbconst'):
            words = sig.split()
            name = words[0]
            restofline = ' '.join(words[1:])

        # variables could have '(subscripts)' and/or 'As type' after subscripts
        # so we try '(' first and then 'As'
        if self.objtype == 'vbvar':
            pos = sig.find('(')
            if pos > 0:
                name = sig[:pos]
                restofline = sig[pos:]
            else:
                pos = sig.find('As')
                if pos > 0:
                    name = sig[:pos]
                    restofline = sig[pos:]

        return name.strip(), arglist, restofline.strip()

    def handle_signature(self, sig: str, signode: desc_signature) -> tuple[str, str]:

        name, arglist, restofline = self.parse_signature(sig)

        parent = self.env.ref_context.get('vba:object', None)
        mod_name = self.env.ref_context.get('vba:module')

        add_name = addnodes.desc_addname('', '')
        if parent:
            fullname = '.'.join([parent, name])
            add_name += addnodes.desc_sig_name(parent, parent)
            add_name += addnodes.desc_sig_punctuation('.', '.')
        else:
            fullname = name

        signode['module'] = mod_name
        signode['object'] = parent
        signode['fullname'] = fullname

        display_prefix = self.get_display_prefix()
        if display_prefix:
            signode += addnodes.desc_annotation('', '', *display_prefix)

        if self.env.config.vba_display_fullname:
            signode += add_name
        signode += addnodes.desc_name('', '', addnodes.desc_sig_name(name, name))

        # if arglist:
        if self.objtype in ('vbsub', 'vbfunc'):
            paramlist = addnodes.desc_parameterlist()
            paramlist += addnodes.desc_parameter(arglist, arglist)
            signode += paramlist

        if restofline:
            signode += [addnodes.desc_sig_space(),
                    #   addnodes.desc_type(restofline, restofline),
                    #   addnodes.desc_returns(restofline, restofline),
                      addnodes.desc_sig_keyword(restofline, restofline),
                    ]

        # a, b = self.xhandle_signature(sig, signode)
        return fullname, parent

    def _object_hierarchy_parts(self, sig_node: desc_signature) -> tuple[str, ...]:
        if 'fullname' not in sig_node:
            return ()
        modname = sig_node.get('module')
        fullname = sig_node['fullname']

        if modname:
            return (modname, *fullname.split('.'))
        else:
            return tuple(fullname.split('.'))

    def add_target_and_index(self, name_obj: tuple[str, str], sig: str,
                             signode: desc_signature) -> None:
        mod_name = self.env.ref_context.get('vba:module')
        fullname = (mod_name + '.' if mod_name else '') + name_obj[0]
        node_id = make_id(self.env, self.state.document, '', fullname)
        signode['ids'].append(node_id)
        self.state.document.note_explicit_target(signode)

        domain = cast(VBADomain, self.env.get_domain('vba'))
        domain.note_object(fullname, self.objtype, node_id, location=signode)

        if 'no-index-entry' not in self.options:
            indextext = self.get_index_text(mod_name, name_obj)  # type: ignore[arg-type]
            if indextext:
                self.indexnode['entries'].append(('single', indextext, node_id, '', None))

    def get_index_text(self, objectname: str, name_obj: tuple[str, str]) -> str:
        name, obj = name_obj

        name = name.split('.')[-1]

        if self.objtype == 'vbclass':
            return _('%s (class in %s)') % (name, objectname)
        elif self.objtype == 'vbform':
            return _('%s (Formular in %s)') % (name, objectname)
        elif self.objtype == 'vbmodule':
            return _('%s (VB-Modul in %s)') % (name, objectname)
        elif self.objtype == 'vboffice_obj':
            return _('%s (Office-Objekt in %s)') % (name, objectname)

        elif self.objtype in ('vbfunc', 'vbsub'):
            if not obj:
                return _('%s() (built-in function)') % name
            return _('%s() (%s procedure)') % (name, obj)

        elif self.objtype == 'vbprop':
            return _('%s (%s property)') % (name, obj)
        elif self.objtype == 'vbvar':
            return _('%s (%s var)') % (name, obj)
        elif self.objtype == 'vbconst':
            return _('%s (%s const)') % (name, obj)

        # elif self.objtype == 'data':
        #     return _('%s (global variable or constant)') % name
        # elif self.objtype == 'attribute':
        #     return _('%s (%s attribute)') % (name, obj)
        return ''

    def before_content(self) -> None:
        """Handle object nesting before content

        :py:class:`VBAObject` represents VBA language constructs. For
        constructs that are nestable, this method will build up a stack of the
        nesting hierarchy so that it can be later de-nested correctly, in
        :py:meth:`after_content`.

        For constructs that aren't nestable, the stack is bypassed, and instead
        only the most recent object is tracked. This object prefix name will be
        removed with :py:meth:`after_content`.

        The following keys are used in ``self.env.ref_context``:

            vba:objects
                Stores the object prefix history. With each nested element, we
                add the object prefix to this list. When we exit that object's
                nesting level, :py:meth:`after_content` is triggered and the
                prefix is removed from the end of the list.

            vba:object
                Current object prefix. This should generally reflect the last
                element in the prefix history
        """
        prefix = None
        if self.names:
            (obj_name, obj_name_prefix) = self.names.pop()
            prefix = obj_name_prefix.strip('.') if obj_name_prefix else None
            if self.allow_nesting:
                prefix = obj_name
        if prefix:
            self.env.ref_context['vba:object'] = prefix
            if self.allow_nesting:
                objects = self.env.ref_context.setdefault('vba:objects', [])
                objects.append(prefix)

    def after_content(self) -> None:
        """Handle object de-nesting after content

        If this class is a nestable object, removing the last nested class prefix
        ends further nesting in the object.

        If this class is not a nestable object, the list of classes should not
        be altered as we didn't affect the nesting levels in
        :py:meth:`before_content`.
        """
        objects = self.env.ref_context.setdefault('vba:objects', [])
        if self.allow_nesting:
            with contextlib.suppress(IndexError):
                objects.pop()

        self.env.ref_context['vba:object'] = (objects[-1] if len(objects) > 0
                                             else None)

    def _toc_entry_name(self, sig_node: desc_signature) -> str:
        if not sig_node.get('_toc_parts'):
            return ''

        config = self.env.app.config
        objtype = sig_node.parent.get('objtype')
        if config.add_function_parentheses and objtype in {'function', 'sub'}:
            parens = '()'
        else:
            parens = ''
        *parents, name = sig_node['_toc_parts']
        if config.toc_object_entries_show_parents == 'domain':
            return sig_node.get('fullname', name) + parens
        if config.toc_object_entries_show_parents == 'hide':
            return name + parens
        if config.toc_object_entries_show_parents == 'all':
            return '.'.join(parents + [name + parens])
        return ''


class VBACallable(VBAObject):
    """Description of a VBA function or sub."""
    has_arguments = True

    doc_field_types = [
        TypedField('vbargs', label=_('Argumente'), names=('arg',),
                   typerolename='vbdata', typenames=('type',)),
        GroupedField('errors', label=_('Throws'), rolename='vbdata',
                     names=('raise', ),
                     can_collapse=True),
        Field('returnvalue', label=_('Returns'), has_arg=False,
              names=('returns', 'return')),
        Field('returntype', label=_('Return type'), has_arg=False,
              names=('rtype',)),
    ]

class VBAClass(VBAObject):
    '''description of VBA-classes'''
    object_type_prefix = 'Class'
    allow_nesting = True

class VBAModule(VBAObject):
    '''description of VBA-modules'''
    object_type_prefix = 'vbmodule'
    allow_nesting = True

class VBAForm(VBAObject):
    '''description of VBA-forms'''
    object_type_prefix = 'vbform'
    allow_nesting = True

class VBAOfficeObject(VBAObject):
    '''description of VBA-Objects (like Worksheets)'''
    object_type_prefix = 'office_object'
    allow_nesting = True

class VBAFunction(VBACallable):
    '''description of a vba function'''
    object_type_prefix = 'Function'

class VBASub(VBACallable):
    '''description of a vba subroutine'''
    object_type_prefix = 'Sub'

class VBAprop(VBAObject):
    '''description of a vba property'''
    object_type_prefix = 'Property'

class VBAvar(VBAObject):
    '''description of a vba variable (global or class-member)'''
    object_type_prefix = ''

class VBAConst(VBAObject):
    '''description of a vba Const-Value (global or class-member)'''
    object_type_prefix = 'Const'


class VBAFile(SphinxDirective):
    """
    Directive to mark description of a new VBA File

    This directive specifies the module name that will be used by objects that
    follow this directive.

    Options
    -------

    no-index
        If the ``:no-index:`` option is specified, no linkable elements will be
        created, and the module won't be added to the global module index. This
        is useful for splitting up the module definition across multiple
        sections or files.

    :param mod_name: Module name
    """

    has_content = True
    required_arguments = 1
    optional_arguments = 0
    final_argument_whitespace = True
    option_spec: OptionSpec = {
        'no-index': directives.flag,
        'no-contents-entry': directives.flag,
        'no-typesetting': directives.flag,
        'noindex': directives.flag,
        'nocontentsentry': directives.flag,
    }

    def run(self) -> list[Node]:
        mod_name = self.arguments[0].strip()
        self.env.ref_context['vba:module'] = mod_name
        no_index = 'no-index' in self.options or 'noindex' in self.options

        content_node: Element = nodes.section()
        # necessary so that the child nodes get the right source/line set
        content_node.document = self.state.document
        nested_parse_with_titles(self.state, self.content, content_node, self.content_offset)

        ret: list[Node] = []
        if not no_index:
            domain = cast(VBADomain, self.env.get_domain('vba'))

            node_id = make_id(self.env, self.state.document, 'module', mod_name)
            domain.note_module(mod_name, node_id)
            # Make a duplicate entry in 'objects' to facilitate searching for
            # the module in VBADomain.find_obj()
            domain.note_object(mod_name, 'module', node_id,
                               location=(self.env.docname, self.lineno))

            # The node order is: index node first, then target node
            indextext = _('%s (module)') % mod_name
            inode = addnodes.index(entries=[('single', indextext, node_id, '', None)])
            ret.append(inode)
            target = nodes.target('', '', ids=[node_id], ismod=True)
            self.state.document.note_explicit_target(target)
            ret.append(target)
        ret.extend(content_node.children)
        return ret


class VBAXRefRole(XRefRole):
    def process_link(self, env: BuildEnvironment, refnode: Element,
                     has_explicit_title: bool, title: str, target: str) -> tuple[str, str]:
        # basically what sphinx.domains.python.PyXRefRole does
        refnode['vba:object'] = env.ref_context.get('vba:object')
        refnode['vba:module'] = env.ref_context.get('vba:module')
        if not has_explicit_title:
            title = title.lstrip('.')
            target = target.lstrip('~')
            if title[0:1] == '~':
                # changed from here (from JS domain). Doesn't seem to make sense for vba
                raise Exception('did not expect ~')
                # title = title[1:]
                # dot = title.rfind('.')
                # if dot != -1:
                #     title = title[dot + 1:]
            #we use only last part for title
            title = title.split('.')[-1]
        if target[0:1] == '.':
            target = target[1:]
            refnode['refspecific'] = True
        return title, target


class VBAProcedureIndex(Index):
    """Index of subs and functions of VBA."""

    name = 'procedureindex'
    localname = _('Liste der Prozeduren')
    shortname = _('Prozeduren')

    def generate(self, docnames=None):
        content = defaultdict(list)

        # sort the list of methods in alphabetical order
        procedures = self.domain.get_objects()
        procedures = sorted(procedures, key=lambda procedure: procedure[0])

        # generate the expected output, shown below, from the above using the
        # first letter of the recipe as a key to group thing
        #
        # name, subtype, docname, anchor, extra, qualifier, description

        for procedure in procedures:
            _name, dispname, typ, docname, anchor, _priority = procedure
            nameparts = dispname.split('.')
            if typ in ('vbfunc', 'vbsub'):
                localname = self.domain.object_types[typ].lname
                content[nameparts[2][0].lower()].append(
                    (nameparts[2], 0, docname , anchor, '', _(localname), nameparts[0] + '.' + nameparts[1]))
                    # (nameparts[-1], typ, docname , anchor, nameparts[1], '', nameparts[0]))

        # convert the dict to the sorted list of tuples expected
        content = sorted(content.items())

        return content, True


class VBADomain(Domain):
    """VBA language domain."""
    name = 'vba'
    label = 'Visual Basic'
    # if you add a new object type make sure to edit VBObject.get_index_string
    # obj_type: ObjType(localname, role)
    object_types = {
        'module':       ObjType(_('module'),     'mod'),
        'vboffice_obj': ObjType(_('office_obj'), 'vbmod'),
        'vbform':       ObjType(_('form'),       'vbmod'),
        'vbclass':      ObjType(_('class'),      'vbmod'),
        'vbmodule':     ObjType(_('vbmodule'),   'vbmod'),
        'vbfunc':       ObjType(_('function'),   'vbproc'),
        'vbsub':        ObjType(_('Sub'),        'vbproc'),
        'vbprop':       ObjType(_('property'),   'vbdata'),
        'vbvar':        ObjType(_('var'),        'vbdata'),
        'vbconst':      ObjType(_('const'),      'vbdata'),
        # 'method':     ObjType(_('method'),    'meth'),
        # 'data':       ObjType(_('data'),      'data'),
        # 'attribute':  ObjType(_('attribute'), 'attr'),
    }
    directives = {
        'module':        VBAFile,    # Exel-, Wordfile or other Office-File
        'vb_office_obj': VBAOfficeObject,
        'vbform':        VBAForm,
        'vbclass':       VBAClass,   #software "container"
        'vbmodule':      VBAModule,

        'vbfunc':        VBAFunction, # Procedures
        'vbsub':         VBASub,

        'vbprop':        VBAprop,     # data
        'vbvar':         VBAvar,
        'vbconst':       VBAConst,
        # 'method':       VBACallable,
        # 'data':        VBAObject,
        # 'attribute': VBAObject,
    }
    roles = {
        'mod':     VBAXRefRole(),
        'vbmod':   VBAXRefRole(),
        'vbproc':  VBAXRefRole(fix_parens=True),
        'vbdata':  VBAXRefRole(),
        # 'meth':  VBAXRefRole(fix_parens=True),
        # 'prop':  VBAXRefRole(),
        # 'class': VBAXRefRole(),
        # 'frm':   VBAXRefRole(),
        # 'var':  VBAXRefRole(),
        # 'attr':  VBAXRefRole(),
    }
    indices = {
        VBAProcedureIndex,
    }
    initial_data: dict[str, dict[str, tuple[str, str]]] = {
        'objects': {},  # fullname -> docname, node_id, objtype
        'modules': {},  # modname  -> docname, node_id
    }

    @property
    def objects(self) -> dict[str, tuple[str, str, str]]:
        return self.data.setdefault('objects', {})  # fullname -> docname, node_id, objtype

    def note_object(self, fullname: str, objtype: str, node_id: str,
                    location: Any = None) -> None:
        if fullname in self.objects:
            docname = self.objects[fullname][0]
            logger.warning(__('duplicate %s description of %s, other %s in %s'),
                           objtype, fullname, objtype, docname, location=location)
        self.objects[fullname] = (self.env.docname, node_id, objtype)

    @property
    def modules(self) -> dict[str, tuple[str, str]]:
        return self.data.setdefault('modules', {})  # modname -> docname, node_id

    def note_module(self, modname: str, node_id: str) -> None:
        self.modules[modname] = (self.env.docname, node_id)

    def clear_doc(self, docname: str) -> None:
        for fullname, (pkg_docname, _node_id, _l) in list(self.objects.items()):
            if pkg_docname == docname:
                del self.objects[fullname]
        for modname, (pkg_docname, _node_id) in list(self.modules.items()):
            if pkg_docname == docname:
                del self.modules[modname]

    def merge_domaindata(self, docnames: list[str], otherdata: dict[str, Any]) -> None:
        # XXX check duplicates
        for fullname, (fn, node_id, objtype) in otherdata['objects'].items():
            if fn in docnames:
                self.objects[fullname] = (fn, node_id, objtype)
        for mod_name, (pkg_docname, node_id) in otherdata['modules'].items():
            if pkg_docname in docnames:
                self.modules[mod_name] = (pkg_docname, node_id)

    def find_obj(
        self,
        env: BuildEnvironment,
        mod_name: str,
        prefix: str,
        name: str,
        typ: str | None,
        searchorder: int = 0,
    ) -> tuple[str | None, tuple[str, str, str] | None]:
        if name[-2:] == '()':
            name = name[:-2]

        searches = []
        if mod_name and prefix:
            searches.append('.'.join([mod_name, prefix, name]))
        if mod_name:
            searches.append('.'.join([mod_name, name]))
        if prefix:
            searches.append('.'.join([prefix, name]))
        searches.append(name)

        if searchorder == 0:
            searches.reverse()

        newname = None
        object_ = None
        for search_name in searches:
            if search_name in self.objects:
                newname = search_name
                object_ = self.objects[search_name]

        return newname, object_

    def resolve_xref(self, env: BuildEnvironment, fromdocname: str, builder: Builder,
                     typ: str, target: str, node: pending_xref, contnode: Element,
                     ) -> Element | None:
        mod_name = node.get('vba:module')
        prefix = node.get('vba:object')
        searchorder = 1 if node.hasattr('refspecific') else 0
        name, obj = self.find_obj(env, mod_name, prefix, target, typ, searchorder)
        if not obj:
            return None
        return make_refnode(builder, fromdocname, obj[0], obj[1], contnode, name)

    def resolve_any_xref(self, env: BuildEnvironment, fromdocname: str, builder: Builder,
                         target: str, node: pending_xref, contnode: Element,
                         ) -> list[tuple[str, Element]]:
        mod_name = node.get('vba:module')
        prefix = node.get('vba:object')
        name, obj = self.find_obj(env, mod_name, prefix, target, None, 1)
        if not obj:
            return []
        return [('vba:' + self.role_for_objtype(obj[2]),  # type: ignore[operator]
                 make_refnode(builder, fromdocname, obj[0], obj[1], contnode, name))]

    def get_objects(self) -> Iterator[tuple[str, str, str, str, str, int]]:
        for refname, (docname, node_id, typ) in list(self.objects.items()):
            yield refname, refname, typ, docname, node_id, 1

    def get_full_qualified_name(self, node: Element) -> str | None:
        modname = node.get('vba:module')
        prefix = node.get('vba:object')
        target = node.get('reftarget')
        if target is None:
            return None
        else:
            return '.'.join(filter(None, [modname, prefix, target]))


def setup(app: Sphinx) -> dict[str, Any]:
    app.add_domain(VBADomain)
    app.add_config_value('vba_display_fullname', False, 'env', types={None},)
    return {
        'version': 'builtin',
        'env_version': 3,
        #'parallel_read_safe': True,
        #'parallel_write_safe': True,
    }
