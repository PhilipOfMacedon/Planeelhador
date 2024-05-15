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
    
    def set_default_document_format(self, wb, ws, data):
        # Set all cells with the same font and content alignment
        wb.formats[0].font_name = data["FORMATS"]["default"]["font_name"]
        wb.formats[0].font_size = data["FORMATS"]["default"]["font_size"]
        wb.formats[0].set_align(data["FORMATS"]["default"]["align"])
        wb.formats[0].set_align(data["FORMATS"]["default"]["valign"])
        
        # Adjust the main column widths and the logo line height 
        ws.set_row_pixels(0, data["FORMATS"]["rowHeights"]["timbra"])
        ws.set_column_pixels(0, 0, data["FORMATS"]["colWidths"]["id"])
        ws.set_column_pixels(1, 1, data["FORMATS"]["colWidths"]["desc"])
        ws.set_column_pixels(2, 4, data["FORMATS"]["colWidths"]["muq"])
        ws.set_column_pixels(5, 6, data["FORMATS"]["colWidths"]["valores"])
    
    def write_document_header(self, wb, ws, data):
        formats = data["FORMATS"]
        
        ws.insert_image(0, 1, "./images/" + self.empresa + ".png", data["FORMATS"]["logo"])
        line_A2G2_format = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["middle_left"])
        ws.merge_range("A2:G2", "Ao Illm.o Sr. Pregoeiro de", line_A2G2_format)
    
    def generate_file(self):
        
        data = load_data()
        
        opt = {
            "strings_to_numbers": True
        }
        
        wb = xlsx.Workbook(filename=self.arquivo, options=opt)
        ws = wb.add_worksheet("PROPOSTA")
        
        self.set_default_document_format(wb, ws, data)
        self.write_document_header(wb, ws, data)
        
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