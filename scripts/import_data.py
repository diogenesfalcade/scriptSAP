"""Module that have tools to handle files, directory and internal data that must be saved
"""
from tkinter import filedialog
import os
import json
from scripts.logger import logger

STD_DIR_PATH = 'Saves'
FILE_NAME = 'internal data.json'


def save_json(data: str, path: str, file_name: str) -> None:
    """Save data in a json file, if the file doesn't exists it will be created
    
    data: str
        Data that will be saved
    path: str
        Path where the json file will be saved
    file_name: str
        Name of the json file
        
    return: None
    """
    old_data = read_json(path, file_name)
    data.update(old_data)
    if os.path.isdir(path):
        with open(path + '\\' + file_name, 'w') as file:
            json.dump(data, file)
    else:
        make_dir(path)
        with open(path + '\\' + file_name, 'w') as file:
            json.dump(data, file)


def read_json(path: str = STD_DIR_PATH, file_name: str = FILE_NAME) -> dict:
    """Loads data from a json file into a dictionary
    
    path: str 
        Path to the desired file
    file_name: str
        Name of the file that will be open
        
    return: dict
        A dictionary containing all data loaded from json file. 
    """
    if os.path.isdir(path):
        try:
            with open(path + '\\' + file_name, 'r') as file:
                paths = json.load(file)
            return paths
        except FileNotFoundError as error:
            logger.exception(error)
            return {}
    else: 
        return {}


def make_dir(path: str) -> None:
    """Makes a directory in specified path
    
    path: str
        Path where the new directory will be created
    
    return: None
    """
    try:
        os.mkdir(path)
    except OSError as error:
        logger.exception(error)


def get_data(path: str):
    with open(path, 'r') as file:
        data = file.readlines()       
    return data


def custom_dir():
    direc = filedialog.askdirectory()
    return direc

