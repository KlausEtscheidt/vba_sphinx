# vba_sphinx
 Python tools to document Visual Basic for Applications (VBA) software with [Sphinx](https://www.sphinx-doc.org/en/master/index.html).

For those who don't know Sphinx: It's a tool which generates various nice outputs like html or pdf
from simple text files, written in *reStructuredText* ([reST](https://docutils.sourceforge.io/rst.html)) format.

You can write these text files manualy, but there are tools for some programming languages like pyhton,
which can extract informations from the source code and transform them into reST.

This process can be improved, if the source code contains so called [Docstrings](https://en.wikipedia.org/wiki/Docstring)
or [Docblocks](https://en.wikipedia.org/wiki/Docblock), which give additional information for specific elements of the sources.

Due to the nature of VBA, which is integrated in applications like Excel, the source code can't be accessed from outside of these applications without more ado.

So we have three steps to generate a documentation for VBA and this package consists of three tools, which can be used independant from each other:

 - VBA Codereader: exports the VBA source code from an Office Application into an external file.
 - VBA Parser: can parse VBA source code and convert it into reST formated files.
 - VBA Domain: is a Sphinx extension, which enables Sphinx to process reST files with VBA documentation

All three tools are written in python, but can be used without python knowledge.

# The VBA Codereader ################################################
The codereader can be used to export the VBA source code out of office into plain text files.
For the moment the tool can handle Excel- and Access-files.
After installation of the package you should generate a directory of your choice as working directory.

Inside this working directory we need a configuration file named *vba_codereader.toml* to use the tool.
This file defines the output directory, where the exported files will be stored and the office files
which will be searched for software. You can download an example file from [vba_codereader.zip](https://github.com/KlausEtscheidt/vba_sphinx/files/14805141/vba_codereader.zip).

If you are interested in the toml details [see toml.io](https://toml.io/en/).

For our purpose we first define the output directory for the exported VBA files with:
```code
outdir = '.\mydir_of_choice'
```
The generated files will be named like the office file where the software was found
but with '*.txt' extension. The directory must exist.  
```{important}Use single backslashes for windows pathes.  
Pathes starting with a dot like '.\mydir' are relativ to the current working dir.
```
Next we define a bunch of office files, which shall be searched for VBA software to export.

```{important}
You can not export VBA from Excel and Access in one run of the tool.
So don't mix them in the configuration file.
```
Example of files to be searched:
```code
[[filelist]]
path = 'V:\Tools\Excel Makros'
files = [
    #'MyExcelWorkbook.xlsm',
    '*.xl*',
]
```
The block shown above can be used multiple times in a configuration file.

`[[filelist]]` has to be followed by one `path = " "` and one `files = []` statement.  
The path statement defines a directory and has to follow the same rules as `outdir` above.\
The files-statement defines a comma separated list of files inside the directory defined by the path-statement.\
An asterix * as wildcard is allowed.
With # you can mark a line as comment, so that the entry is inactivated.

In the example above, we define that we will search for VBA software in every *.xl* file in the directory 'V:\Tools\Excel Makros'.

To start the export process, open a windows command shell, go to your working directory (the one with the configuration file) and type the command:
```code
python -m vbasphinx.vba_reader Excel
```
if you have Excel files in your configuration or 
```code
python -m vbasphinx.vba_reader Access
```
for Access.

# The VBA Parser ################################################

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
---
## Procedure index
In addition to the general index the role
```code
:ref:`vba-procedureindex`
```
generates an index, which lists all vba subroutines and functions
in alphabetical order. The name of the procedure is followed by its type (Sub or Function)
and the file and vbmodule where the procedure is located.

---
## Roles

**:vba:mod:** Modul1\
link to a module

**:vba:vbmod:** Modul1.myVBModul\
link to a vbmodule like a VB-form or a VB-classmodule

**:vba:vbproc:** Modul1.vbaclass.mysub\
link to a procedure like Sub or Function

**:vba:vbdata:** Modul1.vbaclass.myprop\
link to property, variable or constant value


