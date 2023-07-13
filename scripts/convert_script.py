"""Responsible for deal with all tasks involving the macro script gerated by SAP
"""
import os
import pathlib
import importlib
from types import ModuleType
import scripts.import_data as import_data
from scripts.logger import logger


SPACE = '    '
STD_MACRO_DIR = 'Macro'
STD_DIR_PYFILE = 'scripts'
STD_PYFILE_NAME = 'macro'
STD_SAVE_PATH = str(pathlib.Path(__file__).resolve().parents[1]).replace('\\', '\\\\')



def import_script(name: str, path: str = STD_MACRO_DIR, key: str = 'findById') -> str:
    """
    Reads the script file and returns the text inside as a string
    
    path: str
        Path to the desired file
    name: str
        Desired file name
    key: str
        Key valor searched in script to separete header
    
    return: str
        All script founded in file without header
    """
    if os.path.isfile(path + '\\' + name):
        with open(path + '\\' + name, 'r') as script:
            text = script.readlines()
        line = find_first_key(text, key)
        text = text[line:]
        return text
    else:
        logger.error(f'None Macro called {name} founded.')


def edit_script(text: str, name_foo: str = 'script_foo', param: str = 'session',
                save_dir: str = 'SAP data', save_file: str = 'test.txt',
                path: str = STD_SAVE_PATH) -> str:
    """
    Edit the sintaxe from the text from VBS to python
    
    text: str
        Script that will be edited
    name_foo: str
        Name of the function that will contain the python script
    param: str
        Name of the required parameter of the function
    save_dir: str
        Name of the directory that the macro will save the exported data
    save_file: str
        Name of the file (with extention) that the macro will criate
    path: str
        Path that the exported data will be saved
    
    return: str
        Edited script in python
    """
    
    for i, line in enumerate(text):
        if not '=' in line:
            line = line.removesuffix('\n')
            words = line.split(' ')
            line = words[0] + '('
            try: 
                line += words[1] + ')\n'
            except:
                line += ')\n'
        if line[0] == "'":
            line = line.replace("'", "#")
        if 'true' in line:
            index = line.find('true')
            line = line[:index] + 'T' + line[index+1:]
        if 'false' in line:
            index = line.find('false')
            line = line[:index] + 'F' + line[index+1:]
        if 'ctxtDY_PATH' in line:
            check_direc(save_dir)
            line = line.split('=')[0] + f" = '{path}\\{save_dir}'\n"
        if 'ctxtDY_FILENAME' in line:
            line = line.split('=')[0] + f" = '{save_file}'\n"
        
        text[i] = line
    text = [SPACE + line for line in text]
    text.insert(0, f'def {name_foo}({param}):\n')    
    
    return text


def find_first_key(text: str, key: str) -> int:
    """
    Find first ocurrence of a string in a text
    
    text: str
        Text where the string will be searched
    key: str
        String searched
    
    return: int or None
        Return the index of the start of the key passed. If nothing is founded,
        return None
    """
    for i, line in enumerate(text):
        if line.find(key) != -1:
            return i


def save_as(text: str, name: str, extention: str = '.py',
            path: str = STD_DIR_PYFILE) -> None:
    if os.path.isdir(path):
        with open(path + '\\' + name + extention, 'a') as script:
            for line in text:
                script.write(line)
    else:
        logger.error(f'{path} not found.')


def check_direc(path: str) -> None:
    """
    Check if a directory exists in a given path. If the directory id missing,
        create the directory.
    
    path: str
        Path to the directory that will be verified. 
    
    """
    if not os.path.isdir(path):
        import_data.make_dir(path)
        logger.warning(f'None directory was founded in {path}. Dir was created.')
        

def file_names(path: str = STD_MACRO_DIR) -> list:
    """
    Returns a list with all file names in path
    
    path: str
        Path thah all files will be listed
        
    return: list
        A list with all names in specified path
    """
    return [file.name for file in os.scandir(path)]


def get_import(module_name: str = STD_PYFILE_NAME, path: str = STD_DIR_PYFILE) -> ModuleType: 
    """
    Import file with specified name
    
    module_name: str
        Name of file that will be imported
    path: str
        Path to the py file desired
    
    return: ModuleType
        Python module imported
    """
    try:
        module = importlib.import_module(f'{path}.{module_name}')
        logger.info('Macro was successfully imported.')
        return module
    except ImportError as error:
        logger.exception(error)


def get_functions(module: ModuleType, path: str = STD_MACRO_DIR) -> list:
    """
    Get all function in specified module
    
    module: ModuleType
        Module with function that will be returned
    path: str
        Path to the module
        
    return: list
        List with all functions (callable) founded in module  
    """
    names = file_names(path)
    if names:
        return [getattr(module, name.split('.')[0].replace(' ', '_')) for name in names]
            
    
def convert_macro(save_name: str = STD_PYFILE_NAME, path: str = STD_MACRO_DIR) -> None:
    """
    Main function that controls the conversion of VBS macro in a Python module
    
    save_name: str
        Name that file will be saved after conversion (without extention)
    path: str
        Path that the file will be saved
    
    """
    check_direc(path)
    files_in_path = file_names(path)
    if files_in_path:
        for name_file in files_in_path:
            text = import_script(name_file)
            text = edit_script(text, name_foo = name_file.split('.')[0].replace(' ', '_'))
            save_as(text, save_name)


def clear_script(name: str = STD_PYFILE_NAME, path: str = STD_DIR_PYFILE) -> None:
    """
    Clear a py file
    
    name: str
        Name of python file that will be cleared
    path: str
        Path of the python file
    """
    name += '.py'
    if name in file_names(path):
        with open(path + '\\' + name, 'r+') as file:
            file.truncate(0)


if __name__ == '__main__':
    #convert_macro(std_macro_name=True)
    print(STD_SAVE_PATH)
