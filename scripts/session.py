import win32com.client
import pythoncom


from scripts.logger import logger


class Session:
    session_counter = 0
    def __new__(cls, *args, **kwargs):
        """Define the order of session based on how many was already created
        """
        cls.session_counter += 1
        created = super().__new__(cls)    
        created.number = cls.session_counter
        return created
    
    
    def __init__(self, session):
        self.check_session(session)
        self.session = session
        self.available = True
        self.thread = None
        if self.number == 1:
            self.main_flag = True
        else:
            self.main_flag = False
        self.generate_id()
   
        
    def update(self, available):
        self.available = available
    
        
    def check_session(self, session):
        if not type(session) == win32com.client.CDispatch:
            logger.error('Invalid type for session')
            raise TypeError('Invalid type for session')
            
    
    def define_thread(self, thread):
        self.available = False
        self.thread = thread
    
    
    def remove_thread(self):
        self.available = True
        self.thread = None
        self.generate_id()
    
    
    def generate_id(self):        
        pythoncom.CoInitialize()
        self.thread_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, self.session)
        
        
    def __str__(self):
        return f'{self.__class__.__name__}'