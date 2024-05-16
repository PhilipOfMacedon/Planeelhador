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
    data5 = None
    
    with open("./text_presets/formats.json") as dict_file:
        data1 = json.load(dict_file)
    
    with open("./text_presets/statements.json", encoding="utf8") as dict_file:
        data2 = json.load(dict_file)
        
    with open("./text_presets/reps.json", encoding="utf8") as dict_file:
        data3 = json.load(dict_file)
    
    with open("./text_presets/signataries.json", encoding="utf8") as dict_file:
        data4 = json.load(dict_file)
        
    with open("./text_presets/header.json", encoding="utf8") as dict_file:
        data5 = json.load(dict_file)
        
    data = {**data1, **data2, **data3, **data4, **data5}
    
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
        header = data["HEADER"]
        proposal = data["STATEMENTS"][self.empresa]["PROPOSTA"]
        
        ws.insert_image(0, 1, "./images/" + self.empresa + ".png", data["FORMATS"]["logo"])
        A8RML_text = wb.add_format(formats["Arial8Regular"] | formats["middle_left"])
        table_TITLE_text = wb.add_format(formats["Arial8Regular"] | \
            formats["middle_center"] | \
            formats["bold_text"] | \
            formats["border_thin"] | \
            formats["BGColors"][self.empresa])
        table_DESC_text = wb.add_format(formats["Arial8Regular"] | \
            formats["top_left"] | \
            formats["border_thin"] | \
            formats["BGColors"][self.empresa])
        table_IMUQ_text = wb.add_format(formats["Arial8Regular"] | \
            formats["middle_center"] | \
            formats["decimal"] | \
            formats["border_thin"] | \
            formats["BGColors"][self.empresa])
        table_MONEYA8R_text = wb.add_format(formats["Arial8Regular"] | \
            formats["middle_center"] | \
            formats["currency"] | \
            formats["border_thin"] | \
            formats["BGColors"][self.empresa])
        table_MONEYA8B_text = wb.add_format(formats["Arial8Regular"] | \
            formats["middle_center"] | \
            formats["bold_text"] | \
            formats["currency"] | \
            formats["border_thin"] | \
            formats["BGColors"][self.empresa])
        A8RML_text = wb.add_format(formats["Arial8Regular"] | formats["middle_left"])
        A8RMJ_text = wb.add_format(formats["Arial8Regular"] | formats["middle_just"])
        A8BML_text = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["middle_left"])
        A8BMC_text = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["middle_center"])
        A8BUMC_text = wb.add_format(formats["Arial8Regular"] | formats["bold_underline_text"] | formats["middle_center"])
        
        
        ws.set_row_pixels(1, formats["rowHeights"]["default"])
        ws.merge_range("A2:G2", header["A2"], A8BML_text)
        
        ws.set_row_pixels(2, formats["rowHeights"]["orgao"])
        ws.merge_range("A3:G3", header["A3"].format(self.orgao), A8BMC_text)
        
        ws.set_row_pixels(3, formats["rowHeights"]["default"])
        ws.write("A4", header["A4"], A8BML_text)
        ws.merge_range("B4:G4", "", A8BMC_text)
        
        ws.set_row_pixels(4, formats["rowHeights"]["default"])
        ws.merge_range("A5:G5", header["A5"].format(proposal), A8BMC_text)
        
    
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