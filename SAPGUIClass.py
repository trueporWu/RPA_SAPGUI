# -*- coding: utf-8 -*-
"""
Created on Wed Jul 24 15:47:09 2019

@author: pssotosa
"""

import win32com.client
import subprocess
import time

### START CLASS

class SAPGUIError(Exception):
    pass

class SAPGUI():
    
    def __init__(self, _nameConnection, _user, _pass, _pathSapLogon="C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"):
        self.pathSapLogon = _pathSapLogon
        self.nameConnection = _nameConnection
        self.username = _user
        self.password = _pass
    
    @property
    def nameConnection(self): return self.__nameConnection
    @nameConnection.setter
    def nameConnection(self, value): self.__nameConnection = value
    
    @property
    def username(self): return self.__username
    @username.setter
    def username(self, value): self.__username = value
    
    @property
    def pathSapLogon(self): return self.__pathSapLogon
    @pathSapLogon.setter
    def pathSapLogon(self, value): self.__pathSapLogon = value
    
    @property
    def password(self): return self.__password
    @password.setter
    def password(self, value): self.__password = value
    
    @property
    def session(self): return self.__session
    @session.setter
    def session(self, value): self.__session = value
    
        
    def openSAPLogon(self):
        try:
            path = self.pathSapLogon
            subprocess.Popen(path)
            time.sleep(10)            
        except Exception as exp:
            raise SAPGUIError('Ha ocurrido un al ejecutar la aplicacion: '+str(exp))
    
    def loginSap(self):
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return None
            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return None
            connection = application.OpenConnection(self.nameConnection, True)
            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return None
            self.session = connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return None
            
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]").sendVKey(0)            
        except Exception as exp:
            raise SAPGUIError('Ha ocurrido un Error en el Login: '+str(exp))
    
    def openTransaction(self,transationName):
        try:
            self.session.StartTransaction(transationName)
        except Exception as exp:
            raise SAPGUIError('Error al abrir la transaccion: '+exp)    
    
    def execute_workflow(self,steps):
        step_id = 0
        try:
            for step in steps:
                step_id    = step['step_id']
                type_find  = step['type_find']
                cod_find   = step['cod_find']
                action     = step['action']
                comentario = step['comment']
                print('step id ',step_id,comentario)
                if step['value_type'] == 'na': valor = ''
                else: valor= step['value']         
                
                if type_find == 'id': objeto = self.session.findById(cod_find)
                #if type_find == 'text': objeto = self.session.findByText(cod_find)
                self.swithAction(action,objeto,valor)
                
        except Exception as exp:
            raise SAPGUIError('Ha ocurrido un Error en el WorkFlow ID_STEP: '+str(step_id),str(exp))
    
    def swithAction(self,action,objeto,value):
        try:
            if(action=='select'):
                objeto.select()
            elif(action=='pressToolbarContextButton'):
                objeto.pressToolbarContextButton(value)
            elif(action=='selectContextMenuItem'):
                objeto.selectContextMenuItem(value)
            elif(action=='key'):
                objeto.key = value
            elif(action=='press'):
                objeto.press()
            elif(action=='text'):
                objeto.text = value            
        except Exception as exp:
            raise SAPGUIError('Ha ocurrido un Error en la Accion: '+str(exp))
        
    
    def disconnectSession(self):
        self.session = None

### END CLASS  