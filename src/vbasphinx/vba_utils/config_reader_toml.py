'''This module reads a configuration-file in toml format.

    - tomlkit
'''

import os
import glob

import tomlkit.toml_file
import tomlkit.exceptions as tk_exc

class ConfigReaderException(Exception):
    '''class for exceptions resulting from configuration problems.'''

class ConfigReader:
    '''Reading configuration and files to process'''

    def __init__(self, tomlfilename):
        '''reads input directory and files to handle.

        directories are read from toml-file in working dir.

        Raises:
            ConfigReaderException: raised if directory does not exist

        Returns:
            list(str): list of filenames (full path) to handle
        '''

        # reading configuration data from active working directory
        tdir = os.getcwd()
        # tdir = os.path.dirname(os.path.realpath(__file__))

        conf = os.path.join(tdir, tomlfilename)
        try:
            self.toml = tomlkit.toml_file.TOMLFile(conf).read()
        except tk_exc.TOMLKitError as err:
            raise ConfigReaderException(f'error reading configuration file {tomlfilename}:\n{err.args[0]}') from err

    def getdir(self, tml_token):
        '''checks if `path` is an existing dir

        if path is relativ (starting with a `.`),
        complete path (os.realpath) will be returned.

        Args:
            path (str): path to check

        Returns:
            bool: True if path exists
            str: complete path
        '''
        # reading dir for output
        try:
            path = self.toml[tml_token]
        except tk_exc.NonExistentKey:
            raise ConfigReaderException(f'no <{tml_token}> entry in configuration (toml-file).')
        path = os.path.realpath(path)
        if os.path.isdir(path):
            return path
        raise ConfigReaderException(f'directory >>{path}<< does not exist. check your configuration (toml-file).')

    def getfiles(self, tomltoken):
        '''read a filelist from toml

        if fname contains wildcards (*), files in my_dir will be globbed and returned.
        Otherwise my_dir and fname are joined and returned, if fname exists

        Args:
            my_dir (str): directory which should contain fname
            fname (str): filename without path

        Raises:
            XlReaderException: raised, if fname does not exist in my_dir
            XlReaderException: raised, if globbing results in 0 files

        Returns:
            list(str): list of filenames (full path) existing
        '''

        filelist = []

        for flist in self.toml[tomltoken]:
            my_dir = flist['path']
            my_dir = os.path.realpath(my_dir) # if my_dir starts with '.'
            # if list of files is not empty, check if dir exists
            if flist['files']:
                if not os.path.isdir(my_dir):
                    raise ConfigReaderException(f'directory >>{my_dir})'
                                    + '<< does not exist. check your configuration (toml-file).')

            for fname in flist['files']:

                errmsg = f'file >>{fname}<< does not exist in >>{my_dir}<< check your configuration (toml-file).'

                mypath = os.path.join(my_dir, fname)
                if '*' in fname:
                    # glob files
                    files = glob.glob(mypath)
                    if not files:
                        raise ConfigReaderException(errmsg)
                    filelist += files
                else:
                    if os.path.isfile(mypath):
                        filelist.append(mypath)
                    else:
                        raise ConfigReaderException(errmsg)

        return filelist
