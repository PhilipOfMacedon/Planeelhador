import xlsxwriter as xlsx
from datetime import datetime
import locale

class PlaneelhaOutputer:
    def get_datetime_from_string(self, date_string, time_string):
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        date_formats = ["%d/%m/%Y", "%d/%m/%y", "%d/%m"]
        
        date_object = None
        whole_datetime = date_string + " " + time_string
        
        for date_format in date_formats:
            try:
                date_object = datetime.strptime(whole_datetime, date_format + " %H:%M")
                if "%Y" not in date_format:
                    current_year = datetime.now().year
                    date_object = date_object.replace(year=current_year)
                self.dia = str(date_object.day)
                self.mes = str(date_object.month)
                self.ano = str(date_object.year)
                self.diaSemana = date_object.strftime("%A")
                self.mesExtenso = date_object.strftime("%B")
                self.hora = date_object.strftime("%H:%M")
                if int(self.dia) < 10: self.dia = "0" + self.dia
                return
            except ValueError:
                continue

        return
    
    def __init__(self, params=None):
        self.orgao = params["orgao"]
        self.codLicitacao = params["codLicitacao"]
        self.codProcesso = params["codProcesso"]
        self.dia = ""
        self.mes = ""
        self.ano = ""
        self.diaSemana = ""
        self.mesExtenso = ""
        self.hora = ""
        self.get_datetime_from_string(params["dataAbertura"], params["horaAbertura"])
        self.empresa = params["empresa"]
        self.tipo = params["tipo"]
        self.qtd = params["qtd"]
        self.lotesQtd = params["lotesQtd"]
        
        print(f"DATA E HORÁRIO DE ABERTURA: {self.dia} de {self.mesExtenso}, {self.diaSemana}, {self.hora} horas.")