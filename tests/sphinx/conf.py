import os
import sys

ext_dir=os.path.abspath(os.path.join('.','ext'))
ext_dir=os.path.abspath('.')
print (ext_dir)
sys.path.append(ext_dir)

project = 'sphinx_ext'
copyright = '2024, Klaus Etscheidt'
author = 'Klaus Etscheidt'
release = '1.0'
# root_doc = 'Readme'  macht Probleme root muss immer index heißen
htmlhelp_basename = project

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = ['sphinx.ext.autodoc', 'sphinx.ext.coverage', 'sphinx.ext.napoleon',
              'sphinx_rtd_theme', 'myst_parser', 'vba']

# templates_path = ['_templates']
# exclude_patterns = []

language = 'de'
source_encoding = 'utf-8'

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

# html_theme = 'alabaster'
html_theme = 'sphinx_rtd_theme'
html_theme_options = {
    'collapse_navigation': False,
}

# vba_display_fullname = True
toc_object_entries_show_parents = 'hide'

napoleon_use_param = True

myst_heading_anchors = 6
myst_enable_extensions = ["deflist", "attrs_inline", "attrs_block", "colon_fence", "fieldlist"]

# def change(app, what, name, obj, options, lines):
#     print(lines)
#     #name += '#######'

# def setup(app):
#     #     app.connect("autodoc-process-docstring", change)
#     app.add_object_type('vbsub', 'sub', 'single: VBA-Routine; %s', objname='Subroutine')