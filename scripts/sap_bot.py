"""Module that can get SAP session and execute macros
"""

import sys
import win32com.client
import subprocess
import time
import os
import pythoncom
from threading import Thread


from scripts.logger import logger
from scripts.session_handler import SessionHandler

SAP_PATH = r'C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe'
SAP_CONNECTION_NAME = '01 : RP1 - AKSO S/4 HANA 2020 Production System'
SAP_RUNNING_NAME = 'saplogon.exe'
SAP_MAX_WORKERS = 2


#TODO: Garantir que as sessões estejam na tela principal
def wrapper(foo):
    """Função que envelopa a função objetivo, criando as condições necessárias para
    a execução da macro sem a necessidade de altera-la.
    """
    def inner(session, *args, **kwargs):
        pythoncom.CoInitialize()
        # Get instance from the id
        sap_session = win32com.client.Dispatch(
                pythoncom.CoGetInterfaceAndReleaseStream(session.thread_id, pythoncom.IID_IDispatch)
        )
        foo(sap_session, *args, **kwargs)
        pythoncom.CoUninitialize()
    return inner



class SapBot:
    """Responsible for get all the necessarie objects form SAP applicantion and 
        set the conditions to execute a macro. Only one instance of this class 
        is allowed.
    """
    keys_map = {"Enter": 0, "F3": 3, "F8": 8}
    
    def __new__(cls, *args, **kwargs):
        """Only permit the construction of one instance of SapRobot
        """
        if not hasattr(cls, '_criated'):
            cls._criated = super().__new__(cls)    
        return cls._criated
    
    
    def __init__(self, sap_path: str = SAP_PATH, sap_connection_name: str = SAP_CONNECTION_NAME):
        """Constructor method, set the initial informatio.
        
        sap_path: str
            Path to SAP executable. Normally this is standart.
        sap_connection_name: str
            Name of the connection where the session will be requested.
        """
        self.sap_path = sap_path
        self.sap_connection_name = sap_connection_name
        self.loged = False
        self.window = 0
        self.waiting_time = 1 #seconds
        self.extra_workers = SAP_MAX_WORKERS


    def run(self, function, *args, **kwargs):
        """Run the passed function passing the session as first param
        """
        if not callable(function):
            logger.error("Function must be callable")
            #raise
        try:
            function(self.session, *args, **kwargs)
        except Exception as e:
            logger.error(f"{e}\nThe passed function can't be executed.")
            #raise


    def run_thread(self, function, list_arg, *args, **kwargs):
        
        #Inicializando as janelas e sessões do SAP
        self.open_windows(self.extra_workers)
        
        #Contador responsável por gerenciar quando a alimentação de novas threads deve terminar
        count = 0
        arg_iter = iter(list_arg)
        
        #Definição inicial das threads - Todas sessions estão livres
        for _ in range(len(self.sessions.get_available())):
            session = self.sessions["available"]
            foo_args = (session, next(arg_iter)) + args
            th = Thread(target=wrapper(function), args=foo_args, kwargs=kwargs)
            session.define_thread(th)
            count += 1
            th.start()
        
        #Fase de manutenção - Alocar apenas sessions que terminaram execução anterior
        while count < len(list_arg):
            session = self.sessions["available"]
            if session:
                foo_args = (session, next(arg_iter)) + args
                th = Thread(target=wrapper(function), args=foo_args, kwargs=kwargs)
                th.start()
                session.define_thread(th)
                count += 1
            for session in self.sessions.get_unavailable():
                if not session.thread.is_alive():
                    session.remove_thread()    
            time.sleep(self.waiting_time)
        
        #Garantindo que todas as threads estejam terminadas antes de continuar
        for session in self.sessions.get_unavailable():
            session.thread.join()
        time.sleep(20) #TODO: Ajustar delay
  
    def launch_sap_app(self):
        try:
            subprocess.Popen(self.sap_path)
            time.sleep(20) #TODO: Ajustar delay
        except:
            logger.error("Coundn't launch SAP application.")
            raise
        else:
            logger.info('SAP lauched.')
            return True
            
            
    def close_sap(self):
        """Close SAP application via cmd command
        """
        
        try:
            self.sessions.close_session()
            self._connection.CloseSession('ses[0]')
            os.system("TASKKILL /F /IM " + SAP_RUNNING_NAME)
        except Exception as e:
            logger.error(e)
            raise
        except: 
            logger.error(f"Unknown exception. SAP was not closed.")
            raise
        else:
            logger.info('Application closed.')


    def login(self):
        """Main method to get all necessarie objects from SAP engine
        """
        if self.get_application():
            if self.get_connection():
                if self.get_session():
                    self.loged = True
  
        
    def get_application(self):
        try:
            self._sap_gui_auto = win32com.client.GetObject("SAPGUI")
            self._application = self._sap_gui_auto.GEtScriptingEngine           
        except:
            try:
                self.launch_sap_app()
                self._sap_gui_auto = win32com.client.GetObject("SAPGUI")
                self._application = self._sap_gui_auto.GetScriptingEngine          
            except:    
                logger.error("Coudn't launch sap application.")
                raise
        finally:
            if not type(self._sap_gui_auto) == win32com.client.CDispatch:
                logger.error('Invalid type for sap_gui_auto')
                raise TypeError('Invalid type for sap_gui_auto')
            
            logger.info('Application obteined.')
            return True


    def get_connection(self):
        try:
            self._connection = self._application.Children(0)
        except:
            try:
                self._application.OpenConnection(self.sap_connection_name, True)
                self._connection = self._application.Children(0)
            except:
                logger.error("Coudn't launch sap connection. Verify if name has changed.")
                raise
        finally:
            if not type(self._connection) == win32com.client.CDispatch:
                logger.error('Invalid type for connection')
                raise TypeError('Invalid type for connection')
            logger.info('Connection obteined.')
            return True
    
        
    def get_session(self):
        try:
            self.sessions = SessionHandler(self._connection)
        except:
            logger.error("Coudn't get sap session.")
            raise
        finally:
            logger.info('Session obteined.')
            return True


    def open_windows(self, workers: int):
        for i in range(workers):
            self.sessions.new_session()


    @property
    def session(self):
        if self.sessions.main_session.available:
            return self.sessions.main_session.session
        else:
            Exception("Main session not available") 
        
        
    def open_transaction(self, transaction):
        if self.window != 0:
            transaction = "/n" + transaction
        
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = transaction
            self.press_key("Enter")
            self.window += 1
        except:
            logger.error("Transaction can't be opened.")
            raise
    
    
    def go_back(self):
        self.press_key("F3")
        self.window -= 1
    
    
    def press_key(self, key):
        if not key in self.keys_map.keys():
            logger.warning("Key is not implemented.")
            raise NotImplementedError("Key is not implemented.")
        
        self.session.findById("wnd[0]").sendVKey(self.keys_map[key])
        
    
    def save_to_txt(self, save_path, save_name):
        if not os.path.isdir(save_path):
            os.mkdir(save_path)
        if os.path.isfile(f"{save_path}\\{save_name}"):
            os.remove(f"{save_path}\\{save_name}")
            
        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = save_path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = save_name
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
                