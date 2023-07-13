"""Module that can kill all running programs in OS


"""

import os
import subprocess

from scripts.logger import logger

APLICATION = 'EXCEL.EXE'

def close_app(app:str = APLICATION):
    """Close aplicattion application via cmd command

    app: str
        kill task, standard value: EXCEL.EXE
    """
    try:
        os.system("TASKKILL /F /IM " + app)
    except SystemError as error:
        logger.info(f"Coudn't close app. Error: {error}" )
    else:
        logger.info('Application closed.')

def get_running_apps():
    cmd = 'powershell "gps | where {$_.MainWindowTitle } | select ProcessName'
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    running_apps = list()
    for line in proc.stdout:
        if line.rstrip():
            # only print lines that are not empty
            # decode() is necessary to get rid of the binary string (b')
            # rstrip() to remove `\r\n`
            running_apps.append(line.decode().rstrip())
    return running_apps

def close_running_apps(exception: list):
    lista = get_running_apps()
    exception = [item.lower() for item in exception]
    for item in lista:
        if not item.lower() in exception:
             close_app(item + '.exe')
