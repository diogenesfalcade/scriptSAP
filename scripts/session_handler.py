import time


from scripts.session import Session


class SessionHandler:
    def __init__(self, connection):
        self.__index = 0
        self._connection = connection
        self.main_session = Session(self._connection.Children(0))
        self.__sap_windows = [self.main_session]
        self.session_quantity = 1
        self.timeout = 30
        
    
    def new_session(self):
        try:
            session = self._connection.Children(self.session_quantity)
        except:
            self.main_session.session.findById("wnd[0]").sendVKey(74)
            time0 = time.time()
            while True:
                try:
                    session = self._connection.Children(self.session_quantity)
                    break
                except:
                    if (time0 - time.time()) > self.timeout:
                        raise TimeoutError("Coudn't get sap session")
                    
        self.session_quantity += 1
        session_obj = Session(session)     
        self.__sap_windows.append(session_obj)
        
    
    def close_session(self):
        if len(self.get_unavailable()) > 0:
            raise Exception("Some sessions still running.")
        for i in range(1, len(self.__sap_windows)):
            try:
                self._connection.CloseSession(f'ses[{i}]')
            except Exception as e:
                raise e
            
      
    def get_available(self):
        return [session for session in self.__sap_windows if session.available]
    
    
    def get_unavailable(self):
        return [session for session in self.__sap_windows if not session.available]
    
    
    def update(self, index, available):
        if type(index) == list: 
            for item in index:
                self.update(item, available)        
        elif type(index) == int:
            if index == 0:
                self.main_session.update(available=available)
            self.__sap_windows[index-1].update(available=available)
        else:
            raise AssertionError(f"Type must be list or int, but {type(index)} was passed")
        
        
    def __getitem__(self, item):
        assert type(item) == int or type(item) == str
        if type(item) == int:
            return self.__sap_windows[item]
        if item.lower() == "available":
            available = self.get_available()
            if len(available) > 0:
                return available[0]
        elif item.lower() == "unavailable":
            unavailable = self.get_unavailable()
            if len(unavailable) > 0:
                return unavailable[0]
            

    def __setitem__(self, index, session):
        self.__sap_windows[index] = Session(session)
        
        
    def __iter__(self):
        return self
    
    
    def __next__(self):
        try:
            for i in range(self.__index, len(self.__sap_windows)+1):
                session = self.__sap_windows[i]
                self.__index += 1
                if session.available:
                    return session
                
        except IndexError:
            self.__index = 0
            raise StopIteration
    
    
    def __str__(self):
        return f'{self.__class__.__name__}({self.__sap_windows})'
    
    





    