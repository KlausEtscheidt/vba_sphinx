'''This module reads Excel-Files and exports all the visual-basic routines inside.

    The vba-sources are exported in one text file per office file 
    with the same name but *.txt extension.
    Directory for output and files for input have to be defined in file named vba_codereader.toml
    in the current working directory.

    Because vba-software is organized in components (like forms, modules, classmodules, etc)
    this structure can be found in the outfile through headers for each component, like so:

    ============================================================
    form: mainform
    ============================================================
    
    with "form: mainform" as "type: name" of the component.

    blank lines are stripped (not exported) and continuation lines (_ at the end) are joined.

    The strucure is designed to be read by a parser, which can write reST-Files for Sphinx.

    Python-Packages required:

    - win32com

    if error win32com.gen_py has no attribute 'mcsIDToPackageMap' occurs:
             clear contents of c:/users/<username>/Appdata/Local/Temp/gen_py
'''

import os
import logging
import sys
from abc import ABCMeta, abstractmethod

import win32com.client as win32
import pywintypes

from vbasphinx.vba_utils.config_reader_toml import ConfigReader
from vbasphinx.vba_utils.vba_logging import setup_logger

# Basis-Verzeichnis
# my_file = os.path.realpath(__file__) # Welcher File wird gerade durchlaufen
# my_dir = os.path.dirname(my_file)
# parent_dir = os.path.join(my_dir, '..')
# sys.path.insert(0, parent_dir)

DONTCLOSE = False

log = logging.getLogger()

# pylint doesnt know pywintypes.com_error
# pylint: disable=no-member

# text for output
DIVIDER = '='*80

class VBReaderException(Exception):
    '''class for exceptions'''

class VBAReader(ABCMeta):
    '''Reading Office-Files and exporting vba-code'''
    # __metaclass__ = abc.ABCMeta
    app = None
    appname = ''
    act_file = None
    files2process = []
    vba_outdir = ''

    @classmethod
    def run(mcs):
        '''Load Office-files and export software'''

        #setup logging
        setup_logger('./vba_codereader.log')

        # read what to do
        cfg = ConfigReader('vba_codereader.toml')
        mcs.vba_outdir = cfg.getdir('outdir')
        allfiles = cfg.getfiles('filelist')
        # exclude backupfiles
        for file in allfiles:
            if not '~' in file:
                mcs.files2process.append(file)

        mcs.__start_app()

        for i, fpath in enumerate(mcs.files2process):
            log.info ('\n%s', DIVIDER)
            log.info ('file %d of %d', i+1, len(mcs.files2process))
            log.info ('try to open%s', fpath)
            mcs.__handle_file(fpath)

        if not DONTCLOSE:
            mcs.app.Application.Quit()
            log.info ('Stopped %s\n', mcs.appname)

    @classmethod
    def __start_app(mcs):
        '''starts Office-App

        Raises:
            VBReaderException: raised if App is already running
            VBReaderException: raised if App can't be started
        '''
        # start app
        try:
            # Check, if it's already running
            # mcs.xl_app = win32.GetActiveObject(appname + '.Application')
            mcs.app = win32.GetActiveObject(mcs.appname + '.Application')
            # raise XlReaderException('Excel is running. Please close all instances.')
        except (AttributeError, pywintypes.com_error) as err: # pylint: disable=unused-variable

            # Ok, Excel is not running, so start it
            try:
                mcs.app = win32.gencache.EnsureDispatch(mcs.appname + '.Application')
                # mcs.xl_app = win32.Dispatch('Excel.Application')
                # mcs.xl_app = win32.Dispatch('Access.Application')
            except pywintypes.com_error as err2:
                raise VBReaderException('Could not start {mcs.appname}.\n{err2}') from err2

        mcs.app.Visible = True
        log.info ('%s is running\n', mcs.appname)

    @classmethod
    def __handle_file(mcs, path):
        '''exports all software components of the path db

        Args:
            xl_path (str): path to access db
        '''

        # try to open workbook (raises error if we fail)
        mcs.open_file(path)
        log.info ('opened\n')

        # get name of output-file
        _, name_ext = os.path.split(path)
        name, _ = os.path.splitext(name_ext)
        outpath = os.path.join(mcs.vba_outdir, name + '.txt')

        with open(outpath, 'w', encoding='utf-8') as outf:
            for comp in mcs.app.VBE.ActiveVBProject.VBComponents:
                mcs.__handle_vba_component(outf, comp)
            outf.write('<EndofFile>')
        log.info ('\nWritten %s', outpath)
        opened_by_someone = False

        if not DONTCLOSE: #dont close for tests
            try:
                # If abstractmethod is opened by some else, we can't close it
                mcs.act_file.Close()
            except pywintypes.com_error as err:
                if err.hresult == -2147352567:
                    if err.excepinfo[5] == -2146827284:
                        #opened by someone else
                        opened_by_someone = True
                if not opened_by_someone:
                    mcs.app.Application.Quit()
                    raise err

        if opened_by_someone:
            log.info('\n %s still opened by someone.', name_ext)
        else:
            log.info ('\nClosed %s', path)
        log.info (DIVIDER)

    @classmethod
    @abstractmethod
    def open_file(mcs, _path):
        '''opens office file'''

    @classmethod
    def __handle_vba_component(mcs, outf, comp):
        '''export sourcecode of one vba.component (form, module, worksheet, etc)

        Args:
            outf (TextIOWrapper): File for output of comp's data
            comp (com-object): VBComponent

        Raises:
            VBReaderException: Type of VBComponent not implemented
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

class AccessReader(VBAReader):
    '''Reading Access-Files and exporting vba-code'''

    appname = 'Access'

    @classmethod
    def open_file(mcs, path):
        '''opens Access database

        Args:
            path (str): path to Access file

        Raises:
            VBReaderException: raised, if file can't be openend
        '''
        _, name_ext = os.path.split(path)
        try:
            mcs.app.OpenCurrentDatabase(path)
            mcs.act_file = mcs.app.CurrentDb()
        except pywintypes.com_error as err:  # pylint: disable=unused-variable
            if err.hresult == -2147352567:
                if err.excepinfo[5] == -2146820421:
                    # was already open
                    mcs.act_file = mcs.app.CurrentDb()
            else:
                raise VBReaderException(f'Could not open {name_ext}') from err

        if not mcs.act_file:
            mcs.app.Application.Quit()
            raise VBReaderException(f'{mcs.appname} could not open {name_ext}')

class ExcelReader(VBAReader):
    '''Reading Excel-Files and exporting vba-code'''

    appname = 'Excel'

    @classmethod
    def open_file(mcs, path):
        '''opens Excel database
        
        Args:
            path (str): path to Excel file

        Raises:
            VBReaderException: raised, if file can't be openend
        '''
        _, name_ext = os.path.split(path)
        try:
            # already opened ?
            mcs.act_file = mcs.app.Workbooks[name_ext]
        except pywintypes.com_error as err:  # pylint: disable=unused-variable
            try:
                # open
                mcs.act_file = mcs.app.Workbooks.Open(path)
            except pywintypes.com_error as err2:
                mcs.app.Application.Quit()
                raise VBReaderException(f'{mcs.appname} could not open {name_ext}') from err2

if __name__ == '__main__':
    app = sys.argv[1]
    if app == "Excel":
        ExcelReader.run()
    elif app == "Access":
        AccessReader.run()
    else:
        print ('\nusage:\n\npython -m vbasphinx.vba_reader Excel\n')
        print ('or:\n\npython -m vbasphinx.vba_reader Access')
