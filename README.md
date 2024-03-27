# vba_sphinx
 Python tools to document Visual Basic Software with Sphinx.

 The package consists of three tools, which can be used independant from each other.

 - VBA Codereader can read the VBA source code from an Office Application.
 - VBA Parser can parse VBA source code and convert it into rest format
 - VBA Domain is a Sphinx extension, which enables Sphinx to read Rest files with VBA documentation

 # The VBA Domain

 The VBA domain (name **vba**) can be used to document Visual Basic for Applications software.

## Directives
**.. vba:module::** filename\
describes the source of the software e.g. an Excel Workbook. This directive sets the module name for object declarations that follow after. The module name is used in the global module index and in cross references. For all objects that belong to this module, the filename is shown in the index as source for these objects.

**.. vba:vb_office_obj::** name\
**.. vba:vbform::** name\
**.. vba:vbclass::** name\
**.. vba:vbmodule::** name\
these directives describe modules inside a vba file, which are classical software modules (vbmodule), class modules (vbclass), user forms (vbform) or office objects (vb_office_obj) e.g. an Excel-Sheet.
All this modules can contain software.

**.. vba:vbfunc::** funcname(signature) As vbtype

**.. vba:vbsub::** funcname(signature) As vbtype

**.. vba:vbprop::** propertyname

**.. vba:vbvar::** varname

**.. vba:vbconst::** varname

## Roles

**:vba:mod:** Modul1\
link to a module

**:vba:vbmod:** Modul1.myVBModul\
link to a vbmodule like VB form or VB class module

**:vba:vbproc:** Modul1.vbaclass.mysub\
link to a procedure like Sub or Function

**:vba:vbdata:** Modul1.vbaclass.myprop\
link to property, variable or constant value

options for module:
- no-index'
- no-contents-entry'
- no-typesetting'
- noindex'
- nocontentsentry'

nesting allowed:
- vb_office_obj
- vbform
- vbclass
- vbmodule

fields for vbfunc, vbsub (VBACallable)
```python
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
```

options for all directives

```python
class VBAObject(ObjectDescription[tuple[str, str]]):
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
        'withevents': directives.flag,
    }
```



