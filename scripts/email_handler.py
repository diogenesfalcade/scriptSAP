"""Deal with all operations to send a report by email
"""
import win32com.client
import win32ui
import time
import os
import subprocess
from scripts.logger import logger


OUTLOOK_PATH = r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'
OUTLOOK_WINDOW_NAME = 'Microsoft Outlook'
OUTLOOK_RUNNING_NAME = 'outlook.exe'


class EmailHandler:
    """Handles with all tasks to send the email with Python context manager
    """
    def __init__(self):
        self._outlook_path = OUTLOOK_PATH
        self._outlook_window_name = OUTLOOK_WINDOW_NAME
        self._outlook_running_name = OUTLOOK_RUNNING_NAME
        self._delay = 5
    
    def __enter__(self):
        """Returns the mail object to be edited and sended        
        """
        self.get_email()
        return self.mail

    
    def __exit__(self, exception_type, exception_value, exception_traceback):
        """Close application after a delay
        """
        time.sleep(self.delay) #TO DO: Ajustar delay
        self.close_outlook()
        if exception_type or exception_value or exception_traceback:
            logger.error(f'{exception_type} {exception_value} {exception_traceback}')


    def get_email(self):
        """Get outlook object, if application is not running, launch it first
        """
        if self.check_outlook():
            self._outlook = win32com.client.Dispatch('outlook.application')
            self._mail = self._outlook.CreateItem(0)
        else:
            self.launch_outlook()
            time.sleep(10) #TO DO: Ajustar delay
            self._outlook = win32com.client.Dispatch('outlook.application')
            self._mail = self._outlook.CreateItem(0)
        try:
            logger.info(f'Outlook object was obtained with success. {self.mail}')
        except:
            logger.error(f'Outlook object was missing.')

       
    def launch_outlook(self):
        """Launch Outlook application
        """
        try:
            subprocess.Popen(self._outlook_path)
            time.sleep(self.delay) #TO DO: Ajustar delay
        except SystemError as error:
            logger.error(f"Coudn't launch outlook application. {error}")
        else:
            logger.info('Outlook lauched.')


    def check_outlook(self):
        """Check if Outlook is running
        
        return: bool
            If is running return True, else return False
        """
        try:
            win32ui.FindWindow(None, self._outlook_window_name)
            return True
        except win32ui.error as error:
            logger.warning(f"Outlook wasn't running: {error}")
            return False
    
    
    def close_outlook(self):
        """Close de Outlook application with cmd command
        """
        try:
            os.system("TASKKILL /F /IM " + self._outlook_running_name)
        except SystemError as error:
            logger.info(f"Coudn't kill task. {error}" )
        else:
            logger.info('Outlook closed.')


    def send(self):
        """Send the email and log it.
        """
        if self.to:
            logger.info('Email sended.')
            self.mail.Send()
        else:
            logger.warning('None recipient founded.')


    @property
    def mail(self):
        return self._mail
    
    @property
    def delay(self):
        return self._delay
    
    @delay.setter
    def delay(self, valor):
        self._delay = valor
        

