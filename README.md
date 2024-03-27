# vba_sphinx
 Python tools to document Visual Basic Software with Sphinx.

 The package consists of three tools, which can be used independant from each other.

 - VBA Codereader can read the VBA source code from an Office Application.
 - VBA Parser can parse VBA source code and convert it into rest format
 - VBA Domain is a Sphinx extension, which enables Sphinx to read Rest files with VBA documentation

 # The VBA Domain

 The VBA domain (name **vba**) can be used to document Visual Basic for Applications software.

**.. vba:module::** filename
: describes the source of the software e.g. an Excel Workbook

Term *with Markdown*
: Definition [with reference](syntax/definition-lists)

  A second paragraph
: A second definition

- module (File)
- vb_office_obj
- vbform
- vbclass
- vbmodule
- vbfunc
- vbsub
- vbprop
- vbvar
- vbconst

roles:
- mod
- vbmod
- vbproc
- vbdata

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



