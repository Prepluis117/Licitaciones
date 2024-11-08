import re

class Validar:
    def __init__(self) -> None:
        pass
    @staticmethod
    def ValNum(texto):
        aux=False
        resp=''
        while aux==False:
            resp=input(texto)
            aux=resp.isnumeric()
        return int(resp)
    
    @staticmethod
    def ValEle(texto):
        aux=False
        resp=''
        while aux==False:
            resp=input(texto)
            try:
                resp=int(resp)
                if resp!=0 and resp!=1:
                    raise Exception("Mal")
            except:
                aux=False
            else:
                aux=True
        return resp

    @staticmethod
    def ValFlo(texto):
        aux=False
        resp=''
        while aux==False:
            resp=input(texto)
            try:
                resp=float(resp)
            except:
                aux=False
            else:
                aux=True
        return resp
    
    @staticmethod
    def ValGen(texto):
        aux=False
        resp=''
        while aux==False:
            resp=input(texto)
            if not resp:aux=False
            else:aux=True
        return resp
    
    @staticmethod
    def ValEnt(texto,num):
        aux=False
        resp=''
        while aux==False:
            resp=input(texto)
            try:
                resp=int(resp)
                if resp<0 or resp>num:
                    raise Exception("Mal")
            except:
                aux=False
            else:
                aux=True
        return resp
    
