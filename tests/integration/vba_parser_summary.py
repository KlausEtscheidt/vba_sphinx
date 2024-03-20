''' Modul zur Grobkontrolle des Parse-Vorgangs

    Für jeden geparsten File wird eine Zusammenfassung erzeugt.
    - Namen und Typ aller im FIle gefundenen Module
    - Je Modul Anzahl der Konstanten, Varibalen und Methoden

    Die Zusammenfassung wird bei Bedarf als JSON in eine Datei geschrieben.
    Diese dient dann als Referenz.

    Nach jedem Parsen wird die neue Zusammenfassung mit dieser Referenz verglichen.
'''

import json
import os

files = {}

JSONFILE = './tests/vba_vorgabe.json'

class VBAParserCheckExc(Exception):
    '''class for exceptions'''

def export_summary(tree, fullpath):
    '''Zusammenfassung der wichtigsten Knoten erzeugen'''

    _, name_ext = os.path.split(fullpath)
    name, _ = os.path.splitext(name_ext)

    new_file = {}
    files[name] = new_file
    for module in tree.vbamodules:
        new_module = {}
        new_file[module.obj_name] = new_module
        new_module['vbtype'] = module.module_type
        new_module['n_const'] = len(module.const)
        new_module['n_vars'] = len(module.vars)
        new_module['n_methods'] = len(module.methods)

def check_summary():
    '''Neue Zusammenfassung mit Referenz aus Datei vergleichen.'''
    with open(JSONFILE, "r", encoding='utf-8') as fp:
        ref = json.load(fp)

    for fname, file in files.items():
        if not fname in ref:
            answer = input(f'Include new VBA-File {fname} to referene data ? Y/N')
            if answer in ('Y', 'y'):
                ref[fname] = file
                write_summary()
            else:
                raise VBAParserCheckExc(f'Keine Daten für File {fname} in Referenzdatei.')

        ref_file = ref[fname]

        for modname, module in file.items():
            if not modname in ref_file:
                answer = input(f'Include new module {fname}/{modname}  to referene data ? Y/N')
                if answer in ('Y', 'y'):
                    ref_file[modname] = module
                    write_summary()
                else:
                    raise VBAParserCheckExc(f'Modul {fname}/{modname} nicht in Referenz.')
                    
            ref_mod = ref_file[modname]
            msg = f'In Modul {fname}/{modname}' + ' {}: {:d} != {:d} (ref).'
            count_list = ['n_const', 'n_vars', 'n_methods']
            for count in count_list:
                if module[count] != ref_mod[count]:
                    msg = msg.format(count, module[count], ref_mod[count])
                    raise VBAParserCheckExc(msg)

def write_summary():
    '''Zusammenfassung als JSON expotieren.'''
    pretty_json = json.dumps(files, indent=4)
    # print(pretty_json)

    with open(JSONFILE, 'w', encoding='utf-8') as fp:
        fp.writelines(pretty_json)
