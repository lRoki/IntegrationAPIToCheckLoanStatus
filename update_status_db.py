from mysql.connector import connect, Error
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import load_workbook
import os
import sys
import zeep
import time
import random
import logging
root = tk.Tk()
root.withdraw()


logging.basicConfig(filename='filelog.log', filemode="w",level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')




##settings = zeep.Settings(raw_response=True)
try:
    flamex_wsdl = 'http://ws.flamexnet.com.br/WebServiceJoinConsig.asmx?wsdl'
    client = zeep.Client(wsdl=flamex_wsdl)#,settings=settings)
except Exception as e:
    print(e)
    logging.info("Error: %s",e)
    messagebox.showerror("Error",e)
    sys.exit()
    
sLogin = 'LOGIN'
sSenah = 'PASSWORD'


HOST = "IP NUMBER"
DATABASE = "DATABASE NAME"
USER = "DATABASE USER"
PASSWORD = "DATABASE PASWORD"


##HOST = "localhost"
##DATABASE = "conse491_sistema-interno"
##USER = "root"
##PASSWORD = ""


try:
    connection = connect(host=HOST, database=DATABASE, user=USER, password=PASSWORD,autocommit=True,buffered=True)
except Error as e:
    print(e)
    logging.info("Error: %s",e)
    messagebox.showerror("Error",e)
    sys.exit()


cursor = connection.cursor()


def check_contract_exist(contract):
    sql = f'SELECT Situacao FROM usuarios WHERE Contrato = "{contract}"'
    cursor.execute(sql)
    if cursor.fetchone():
        return True
    else:
        return False


def get_contract(limit=None):
    if not limit:
        limit = ""
    else:
        limit = " LIMIT " + str(limit)
        
    sql = f'SELECT Contrato FROM usuarios WHERE usuario = "Flamex"{limit}'
    cursor.execute(sql)
    rows = cursor.fetchall()
    if rows:
        return rows
    else:
        return False

def update_contract_situacao(contract,situacao):
    val = (situacao,contract)
    sql = 'UPDATE usuarios SET Situacao = %s WHERE Contrato = %s'
    cursor.execute(sql,val)
    count_row = cursor.rowcount
    if count_row > 0:
        print("Contrato Actualizado",contract)
        return True
    else:
        print("No Actualizado",contract)
        return False
    

def main():
    logging.info("main Iniciado")
    rows = get_contract()
##    print(rows)
##    situacao = client.service.ConsultaSituacaoProposta(sLogin,sSenah,'751303005')
##    print(situacao)   
##    input()
    print("Cantidad Contractos Flamex en DB:",len(rows))
    count = 0
    count_updated = 0

    

    if rows:
        for contract in rows:
            contract_number = contract[0]
            print("Contract number:",contract_number)
            try:
                logging.info("Contrato: %s",contract_number)
                situacao = client.service.ConsultaSituacaoProposta(sLogin,sSenah,contract_number)
                if "não foi localizado" in situacao:
                    situacao = "Não localizado"
                    print("Não localizado")
                    if update_contract_situacao(contract_number,situacao):
                        logging.info("Updated Não localizado: %s",contract_number)
                        count_updated += 1
                    else:
                        logging.info("Not Updated Existente: %s",contract_number)
                        
                else:
                    print("Situacao:",situacao)
                    if update_contract_situacao(contract_number,situacao):
                        logging.info("Updated %s: %s",situacao,contract_number)
                        count_updated += 1
                    else:
                        logging.info("Not Updated Existente: %s",contract_number)          
                sec = random.randint(3,5)
                count += 1
##                time.sleep(sec)           
            except Exception as e:
                logging.info("Error: %s",e)
                print("Erro:",e)
    else:
        logging.info("No existen contratos en la DB")
        messagebox.showerror("Error","No existen contratos en la DB")

    logging.info("Finalizado - Cantidad Contratos: %s - Cantidad Updates: %s",count,count_updated)
    print(f"Finalizado - Cantidad Contratos: {count} - Cantidad Updates: {count_updated}")

        
            

if __name__ == "__main__":
    main()

