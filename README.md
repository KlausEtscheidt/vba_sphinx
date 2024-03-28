# vba_sphinx
 Python tools to document Visual Basic Software with Sphinx.

 The package consists of three tools, which can be used independant from each other.

 - VBA Codereader can read the VBA source code from an Office Application.
 - VBA Parser can parse VBA source code and convert it into rest format
 - VBA Domain is a Sphinx extension, which enables Sphinx to read Rest files with VBA documentation

 # The VBA Domain ####################################################

 The VBA domain (name **vba**) can be used to document Visual Basic for Applications software.

## Directives
**.. vba:module::** filename\
Describes the source of the software e.g. an Excel Workbook. This directive sets the module name for object declarations that follow after. This "module" has am similar function as in th epython or java script domain. The module name is used in the global module index and in cross references. For all objects that belong to this module, the filename is shown in the index as source for these objects.

Beside this kind of module VB knows internal modules, which act as a kind of container for software and are described below. 

---
**.. vba:vb_office_obj::** name\
**.. vba:vbform::** name\
**.. vba:vbclass::** name\
**.. vba:vbmodule::** name\
These directives describe modules inside a vba file, which are classical software modules (vbmodule), class modules (vbclass), user forms (vbform) or office objects (vb_office_obj) e.g. an Excel-Sheet.
All these modules can contain software in VB, so these directives can have content. Regarding indizes and references they behave like the vba:module directive.

---
**.. vba:vbsub::** subname(arguments)\
**.. vba:vbfunc::** funcname(arguments) As vbtype\
These two describe callable vb procedures, that means subroutines or functions. You can use the full statement like in VB, but without everything preceding the name of the procedure. If you want to see the scope (*Public*, *Private*, ..) or the "*Static*" keyword in your documentation, please use the according directive options for that.

Example:
```code
    .. vba:vbsub:: mysub(a%, b As Integer)
       :scope: Public
       :static:
```
will be rendered to
> *Public Static Sub* mysub(a%, b As Integer)

In the description of a sub or function you can use the following info fields:

- arg: Description of an argument
- raise: Description of an exception, which could be raised by the procedure
- returns, return: Description of the return value.
- rtype: Return type

Example
```code
   .. vba:vbsub:: vbasub(myvar1 As Double, i%)
      :scope: Private

      Here you can describe the sub, 
      with as much of text, as it doesn't bother your reader.

      :arg Double myvar1: description of myvar1
      :arg % i: description of integer variable i
      :returns: description of what is returned
      :rtype: type of what is returned
```

will be rendered as (german version):

> <dl class="vba vbsub"><dt class="sig sig-object vba" id="Modul1.vbaclass.vbasub"><em class="property"><span class="k"><span class="pre">Private</span></span><span class="w"> </span><span class="k"><span class="pre">Sub</span></span><span class="w"> </span></em><span class="sig-name descname"><span class="n"><span class="pre">vbasub</span></span></span><span class="sig-paren">(</span><em class="sig-param"><span class="pre">myvar1</span> <span class="pre">As</span> <span class="pre">Double,</span> <span class="pre">i%</span></em><span class="sig-paren">)</span></dt><dd><p>Here you can describe the sub, with as much of text, as it doesn’t bother your reader.</p>
> <dl class="field-list simple"><dt class="field-odd">Argumente<span class="colon">:</span></dt><dd class="field-odd"><ul class="simple"><li><p><strong>myvar1</strong> (<span><code class="xref vba vba-vbdata docutils literal notranslate"><span class="pre">Double</span></code></span>) – description of myvar1</p>
> </li><li><p><strong>i</strong> (<span><code class="xref vba vba-vbdata docutils literal notranslate"><span class="pre">%</span></code></span>) – description of integer variable i</p></li></ul></dd>
> <dt class="field-even">Rückgabe<span class="colon">:</span></dt><dd class="field-even"><p>description of what is returned</p></dd><dt class="field-odd">Rückgabetyp<span class="colon">:</span></dt><dd class="field-odd"><p>type of what is returned</p></dd></dl></dd></dl>
---
**.. vba:vbprop::** propertyname\
Directive to document VBA properties. Although properties are procedures in VBA, we document only their data (like we would for a variable or const). The propertyname must not be preceded by any keywords. You can use the options `:scope:` and `:static:` for that.
the propertyname can be followed by a VB type character (as in Dim i%) or by ' As type' (e.g. 'As Variant')

Example:
```code
.. vba:vbprop:: myprop As Variant
   :scope: Public
   :static:
```
will be rendered as:

> *Public Static Property* myprop As Variant

and
```code
.. vba:vbprop:: anotherprop%
   :scope: Private
```
will result in:
> *Private Property* anotherprop%
---
**.. vba:vbvar::** varname\
Directive to document VBA variables. Similar to the handling of properties. varname can be followed by a type character or 'As type' string.

In addition to the options for the other vba-objects, the vbvar directive has the :withevents: option, which generates the keyword *'WithEvents'* in the output.

Example:
```code
.. vba:vbvar:: withvar As Application
   :scope: Public
   :withevents:
```
leeds to:
> *Public WithEvents* withvar As Application

---

**.. vba:vbconst::** constname\
Directive to document VBA constant values. Similar to the handling of properties, but the constname can although be followed by an equal sign and the value of the const.

example:
```code
.. vba:vbconst:: mconst2 As Double = 10 * 10
   :scope: Public
```
leeds to:
> *Public Const* mconst2 As Double = 10 * 10
## Roles

**:vba:mod:** Modul1\
link to a module

**:vba:vbmod:** Modul1.myVBModul\
link to a vbmodule like VB form or VB class module

**:vba:vbproc:** Modul1.vbaclass.mysub\
link to a procedure like Sub or Function

**:vba:vbdata:** Modul1.vbaclass.myprop\
link to property, variable or constant value


