import win32com.client
import os
import json
import random
import time
import os

def unglue_Message ():
    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\MessageQueue"
    queue=qinfo.Open(1,0) 
    msg=queue.Receive()
    print("Body :",msg.Body)
    print("------------------------")
    print("")

print("                  __________________________________")
print("                 |  AGENTE 2, ALMACENANDO ARCHIVOS! |")
print("                  ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ")

print("ALMACENANDO LOS ARCHIVOS")
while True:
    print("------------------------")
    print("Nombre del archivo almacenado:")   
    time.sleep(4)
    unglue_Message()