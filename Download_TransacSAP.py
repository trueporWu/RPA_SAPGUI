# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 14:43:02 2019

@author: pssotosa
"""

from SAPGUIClass import SAPGUI
from SAPGUIClass import SAPGUIError
import datetime
import json
import getpass


def read_json_config():
    try:
        with open('config_rpa_sapgui.json') as jfile:
            data = json.load(jfile)
        
        to_large_print = 20
        for k in range(0,to_large_print): print('#', end=' ')
        print('\nRPA SAP GUI v0.1 P.Soto')
        for k in range(0,to_large_print): print('#', end=' ')
        
        if data['username'] == '':
            data['username'] = input('Ingrese Usuario: ')
                                    
        if data['password'] == '':
            data['password'] = getpass.getpass('Ingrese Password: ')
                      
        #creation workflow
        for index_a,element_a in enumerate(data['transactions_execute']):
            print('\n\n Datos Necesarios para Transaccion '+element_a['transaction_code'])
            for index_b,element_b in enumerate(element_a['executions_steps']):
                value_type = element_b['value_type']
                if value_type in ['input','fixed']:
                    step =element_b['step_id']
                    comentario = element_b['comment']
                    print('Paso '+str(step), end=' ')
                    print(comentario, end=' : \t')
                    if value_type == 'input':
                        valor = input('Ingrese Valor: ')
                        data['transactions_execute'][index_a]['executions_steps'][index_b]['value'] = valor
                    else:
                        print(element_b['value'])
        
        for k in range(0,to_large_print): print('#', end=' ')
        print('\n\n ->-> WorkFlow Generado...')
        
        return data
    except Exception as exp:
        raise Exception('Error de lectura de json: '+str(exp))


def main():
    start = datetime.datetime.now()
    
    data_json = read_json_config()
    sapg = SAPGUI(data_json['conectionName'],data_json['username'],data_json['password'],data_json['saplogonpath'])
    try:
        
        sapg.openSAPLogon()
        sapg.loginSap()
        #sapg.openTransaction('ZCO_0137')
        
        for i in data_json['transactions_execute']:
            print('Codigo Transaccion',i['transaction_code'])
            sapg.openTransaction(i['transaction_code'])
            sapg.execute_workflow(i['executions_steps'])
            sapg.session.findbyid('wnd[1]').close()  
        
        
    except Exception as exp:
        print(exp)
    except SAPGUIError as guierr:
        print(guierr)
    finally:
        sapg.disconnectSession()
        end = datetime.datetime.now()
        tiempo = end - start
        print('\n\n Proceso Finalizado en ',tiempo)
    

if __name__ == "__main__":
    main()

       
        
    
    
    
    
    
    
    
    