import win32com.client
import os
import json
import random
import time

def enqueue_Message (x):
    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\MessageQueue"
    queue=qinfo.Open(2,0)   # Open a ref to queue
    msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
    msg.Label="TestMsg"
    msg.Body = x
    msg.Send(queue)
    queue.Close()

palabra=""    
abecedario=['a','b','c','d','e','f','g','h','i','j','k','l','m','n','ñ','o','p','q','r','s','t','u','v','w','x','y','z']
extensiones=['.exe', '.doc', '.mp3','.txt','.docx','.zip','.rar']

class Serializar:
    def __init__(self):
        self.Archivo=palabra    

class Deserializar:
    def __init__(self,j):
        self.__dict__=json.loads(j)


almacenamiento=open("Almacenamiento.txt","a")
print("                  __________________________________")
print("                 |  AGENTE 1, GENERADOR ARCHIVOS!   |")
print("                  ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ̅ ") 

while True:    
    for i in range(5):
        aleatorio =random.randint(0,26)
        palabra += abecedario[aleatorio]
    aleatorio =random.randint(0,6)
    palabra += extensiones[aleatorio]
    pal = Serializar()
    pal.Archivo=palabra

    print("")
    print("----------------------------------------------------------")
    print("Generando en nombre del archivo. Un momento por favor....")
    time.sleep(5)
    print("Excelente el archivo: "+palabra+" se genero correctamente")
    time.sleep(1)
    print("Encolando el archivo a MSMQ, un momento")
    time.sleep(2)
    print("El archivo: "+ palabra+" se encolo correctamente")
    print("----------------------------------------------------------")
    palabra=""

    almacenamiento=open("Almacenamiento.txt","a")
    extejson=json.dumps(pal.__dict__)
    almacenamiento.write(extejson+"\n")
    enqueue_Message(extejson)
    almacenamiento.close()