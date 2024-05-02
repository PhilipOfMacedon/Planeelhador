import xlsxwriter as xlsx
from datetime import datetime
import locale

class PlaneelhaOutputer:
    def getDateFromString(self, strDate):
        
    
    def __init__(self, params=None):
        self.orgao = params["orgao"]
        self.codLicitacao = params["codLicitacao"]
        self.codProcesso = params["codProcesso"]
        self.dataAbertura = params["dataAbertura"]
        self.horaAbertura = params["horaAbertura"]
        self.empresa = params["empresa"]
        
        
        "orgao": self.orgao.get(),
        "codLicitacao": self.codLicitacao.get(),
        "codProcesso": self.codProcesso.get(),
        "dataAbertura": self.dataAbertura.get(),
        "horaAbertura": self.horaAbertura.get(),
        "empresa": self.empresa.get(),
        "tipo": self.tipo.get(),
        "agrupamento": self.agrupamento.get(),
        "qtd": self.qtd.get(),
        "lotesQtd": self.tkVars2Integers()