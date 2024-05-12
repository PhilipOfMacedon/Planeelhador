import xlsxwriter as xlsx
import json
import pprint
from datetime import datetime
import locale

def load_data():
    data = None
    data1 = None
    data2 = None
    data3 = None
    data4 = None
    
    with open("./text_presets/formats.json") as dict_file:
        data1 = json.load(dict_file)
    
    with open("./text_presets/conditions.json", encoding="utf8") as dict_file:
        data2 = json.load(dict_file)
        
    with open("./text_presets/reps.json", encoding="utf8") as dict_file:
        data3 = json.load(dict_file)
    
    with open("./text_presets/signataries.json", encoding="utf8") as dict_file:
        data4 = json.load(dict_file)
        
    data = {**data1, **data2, **data3, **data4}
    
    pp = pprint.PrettyPrinter(depth=8)
    pp.pprint(data["FORMATS"]["default"])
    return data
class PlaneelhaOutputer:
    def get_datetime_from_string(self, date_string, time_string):
        locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
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
    
    def generate_file(self):
        
        data = load_data()
        formats = {}
        
        opt = {
            "strings_to_numbers": True
        }
        
        wb = xlsx.Workbook(filename=self.arquivo, options=opt)
        ws = wb.add_worksheet("PROPOSTA")
        
        
        formats["A8R"] = {
            "Text": wb.add_format(properties=data["FORMATS"]["Arial8Regular"]["Text"]),
            "Decimal": wb.add_format(properties=data["FORMATS"]["Arial8Regular"]["Decimal"]),
            "Currency": wb.add_format(properties=data["FORMATS"]["Arial8Regular"]["Currency"])
        }
        
        formats["A8B"] = {
            "Text": wb.add_format(properties=data["FORMATS"]["Arial8Bold"]["Text"]),
            "Decimal": wb.add_format(properties=data["FORMATS"]["Arial8Bold"]["Decimal"]),
            "Currency": wb.add_format(properties=data["FORMATS"]["Arial8Bold"]["Currency"])
        }
        
        formats["default"] = data["FORMATS"]["default"]
        formats["logo"] = data["FORMATS"]["logo"]
        formats["rows"] = data["FORMATS"]["rowHeights"]
        formats["cols"] = data["FORMATS"]["colWidths"]
        formats["merged_left"] = data["FORMATS"]["merged_left"]
        
        wb.formats[0].font_name = formats["default"]["font_name"]
        wb.formats[0].font_size = formats["default"]["font_size"]
        wb.formats[0].set_align(formats["default"]["align"])
        wb.formats[0].set_align(formats["default"]["valign"])
        
        ws.set_row_pixels(0, formats["rows"]["timbra"])
        ws.set_column_pixels(0, 0, formats["cols"]["id"])
        ws.set_column_pixels(1, 1, formats["cols"]["desc"])
        ws.set_column_pixels(2, 4, formats["cols"]["muq"])
        ws.set_column_pixels(5, 6, formats["cols"]["valores"])
        ws.insert_image(0, 1, "./images/" + self.empresa + ".png", formats["logo"])
        
        wb.close()
        
    
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
        self.arquivo = params["caminhoArquivo"]
        
        self.generate_file()
        
        print(f"DATA E HORÁRIO DE ABERTURA: {self.dia} de {self.mesExtenso}, {self.diaSemana}, {self.hora} horas.")