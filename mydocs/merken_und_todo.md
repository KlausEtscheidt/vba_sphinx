# links 
[packaging](https://py-pkgs.org/welcome)

## pyparse
[pyparse](https://pyparsing-docs.readthedocs.io/en/latest/)\
[Common-Pitfalls](https://github.com/pyparsing/pyparsing/wiki/Common-Pitfalls-When-Writing-Parsers)\
[wiki](https://github.com/pyparsing/pyparsing/wiki)

# ToDo
Was machen wir mit subscripts

prüfen Kommentare unten
```python
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
```

options für File (module directive) prüfen

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




