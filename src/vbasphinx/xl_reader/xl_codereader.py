'''This module reads Excel-Files and exports all the visual-basic routines inside.

    The vba-sources are exported in one Text-File per Excel-File 
    with the same name but *.txt extension.
    Directory for output and files for input have to be defined in file named xl_codereader.toml
    in the current working directory.

    Because vba-software is organized in components (like forms, modules, classmodules, etc)
    this structure can be found in the outfile through headers for each component, like so:

    ============================================================
    form: mainform
    ============================================================
    
    with "form: mainform" as "type: name" of the component.

    blank lines are stripped (not exported) and continuation lines (_ at the end) are joined.

    The strucure is designed to be read by a parser, which can write RST-Files for Sphinx.

    Python-Packages required:

    - win32com

    if error win32com.gen_py has no attribute 'CLSIDToPackageMap' occurs:
             clear contents of c:/users/<username>/Appdata/Local/Temp/gen_py
'''

import os
import logging
import sys

import win32com.client as win32
import pywintypes

# Basis-Verzeichnis
my_file = os.path.realpath(__file__) # Welcher File wird gerade durchlaufen
my_dir = os.path.dirname(my_file)
parent_dir = os.path.join(my_dir, '..')
sys.path.insert(0, parent_dir)

dontclose = True

from vbasphinx.vba_utils.config_reader_toml import ConfigReader
from vbasphinx.vba_utils.vba_logging import setup_logger

log = logging.getLogger()

# pylint doesnt know pywintypes.com_error
# pylint: disable=no-member

# text for output
DIVIDER = '='*80

class XlReaderException(Exception):
    '''class for exceptions'''

class XLReader:
    '''Reading Excel-Files and exporting vba-code'''

    xl_app = None
    workbook = None
    files2process = []
    xl_outdir = ''

    @classmethod
    def run(cls):
        '''Load Excel-files and export software'''

        #setup logging
        setup_logger('./xl_codereader.log')

        # read what to do
        cfg = ConfigReader('xl_codereader.toml')
        cls.xl_outdir = cfg.getdir('outdir')
        allfiles = cfg.getfiles('filelist')
        # exclude backupfiles
        for file in allfiles:
            if not '~' in file:
                cls.files2process.append(file)

        cls.__start_excel()

        for i, fpath in enumerate(cls.files2process):
            log.info ('\n%s', DIVIDER)
            log.info ('file %d of %d', i+1, len(cls.files2process))
            log.info ('try to open%s', fpath)
            cls.__handle_xl_file(fpath)

        if not dontclose:
            cls.xl_app.Application.Quit()
            log.info ('Stopped Excel\n')

    @classmethod
    def __start_excel(cls):
        '''starts Excel

        Raises:
            XlReaderException: raised if Excel is already running
            XlReaderException: raised if Excel can't be started
        '''
        # start Excel
        try:
            # Check, if it's already running
            cls.xl_app = win32.GetActiveObject("Excel.Application")
            # raise XlReaderException('Excel is running. Please close all instances.')
        except (AttributeError, pywintypes.com_error) as err: # pylint: disable=unused-variable

            # Ok, Excel is not running, so start it
            try:
                # cls.xl_app = win32.gencache.EnsureDispatch('Excel.Application')
                cls.xl_app = win32.Dispatch('Excel.Application')
            except pywintypes.com_error as err2:
                raise XlReaderException('Could not start Excel.\n{err2}') from err2

        cls.xl_app.Visible = True
        log.info ('Excel is running\n')

    @classmethod
    def __handle_xl_file(cls, xl_path):
        '''exports all software components of the xl_path workbook

        Args:
            xl_path (str): path to excel-workbook
        '''

        # try to open workbook (raises error if we fail)
        cls.__open_xl_file(xl_path)
        log.info ('opened\n')

        # get name of output-file
        _, name_ext = os.path.split(xl_path)
        name, _ = os.path.splitext(name_ext)
        outpath = os.path.join(cls.xl_outdir, name + '.txt')

        with open(outpath, 'w', encoding='utf-8') as outf:
            for comp in cls.workbook.VBProject.VBComponents:
                cls.__handle_vba_component(outf, comp)
            outf.write('<EndofFile>')
        log.info ('\nWritten %s', outpath)
        opened_by_someone = False

        if not dontclose: #dont close for tests
            try:
                # If Workbook is opened by some else, we can't close it
                cls.workbook.Close()
            except pywintypes.com_error as err:
                if err.hresult == -2147352567:
                    if err.excepinfo[5] == -2146827284:
                        #opened by someone else
                        opened_by_someone = True
                if not opened_by_someone:
                    cls.xl_app.Application.Quit()
                    raise err

        if opened_by_someone:
            log.info('\n %s still opened by someone.', name_ext)
        else:
            log.info ('\nClosed %s', xl_path)
        log.info (DIVIDER)


    @classmethod
    def __open_xl_file(cls, xl_path):
        '''opens xl-workbook

        Args:
            xl_path (str): path to Excel file

        Raises:
            XlReaderException: raised, if file can't be openend
        '''
        _, name_ext = os.path.split(xl_path)
        try:
            # already opened ?
            cls.workbook = cls.xl_app.Workbooks[name_ext]
        except pywintypes.com_error as err:  # pylint: disable=unused-variable

            try:
                # open
                cls.workbook = cls.xl_app.Workbooks.Open(xl_path)
            except pywintypes.com_error as err2:
                cls.xl_app.Application.Quit()
                raise XlReaderException(f'Could not open {name_ext}') from err2

    @classmethod
    def __handle_vba_component(cls, outf, comp):
        '''export sourcecode of one vba.component (form, module, worksheet, etc)

        Args:
            outf (TextIOWrapper): File for output of comp's data
            comp (com-object): VBComponent

        Raises:
            XlReaderException: Type of VBComponent not implemented
        '''

        log.info('found: %s', comp.name)
        #log.info(codmod.CountOfDeclarationLines)

        if comp.type == 1:
            c_type = 'vbmodule'
        elif comp.type == 2:
            c_type = 'vbclass'
        elif comp.type == 100:
            c_type = 'vb_office_obj'
        elif comp.type == 3:
            c_type = 'vbform'
        else:
            outf.write(f'{DIVIDER}\n')
            outf.write(f'error !!!!! component type {comp.type} not yet implemented\n')
            outf.write(f'no export for {comp.name}\n')
            outf.write(f'{DIVIDER}\n')
            return

        codmod = comp.CodeModule
        # is there a CodeModule for this vba-component
        if codmod:
            # isn't it empty
            if codmod.CountOfLines:
                # write header for this component
                outf.write(f'{DIVIDER}\n{c_type}: {comp.name}\n{DIVIDER}\n')
                # get the code out of the module and change crlf to lf
                codetext = codmod.Lines(1, codmod.CountOfLines).replace('\r', '')
                # split text into lines
                i = 0
                lines = codetext.split('\n')
                while i < len(lines):
                    out_line = ''
                    line = lines[i].rstrip()
                    # ignore empty lines
                    if line.strip():
                        # is there a line continuation char at the end of the line
                        while line.strip() and line[-1] == '_':
                            #join them
                            out_line += ' ' + line[0:-2].strip()
                            i += 1
                            line = lines[i].rstrip()
                        out_line += line
                        outf.write(out_line+'\n')
                    i += 1
                outf.write('\n\n')

if __name__ == '__main__':
    XLReader.run()
