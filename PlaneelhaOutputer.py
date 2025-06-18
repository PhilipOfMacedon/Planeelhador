import xlsxwriter as xlsx
import json
from datetime import datetime
import locale

def init_format_variables(wb, formats, empresa):
    global A8RML_text
    global table_HEAD_text
    global table_TITLE_text
    global table_DESC_text
    global table_IMUQ_text
    global table_MONEYA8R_text
    global table_MONEYA8B_text
    global table_HIDDEN_header
    global table_HIDDEN_header2
    global table_HIDDEN_body
    global table_HIDDEN_disabled
    global A8RML_text
    global A8RBC_text
    global A8RMC_text
    global A8RMJ_text
    global A8BML_text
    global A8BMC_text
    global A8BUMC_text
    global A8RMC_number
    global A8BBC_text
    global multiplier_text
    
    A8RML_text = wb.add_format(formats["Arial8Regular"] | formats["middle_left"])
    table_HEAD_text = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["bold_text"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["title"])
    table_TITLE_text = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["bold_text"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["body"])
    table_DESC_text = wb.add_format(formats["Arial8Regular"] | \
        formats["top_left"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["body"])
    table_IMUQ_text = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["decimal"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["body"])
    table_MONEYA8R_text = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["currency"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["body"])
    table_MONEYA8B_text = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["bold_text"] | \
        formats["currency"] | \
        formats["border_thin"] | \
        formats["BGColors"][empresa]["body"])
    table_HIDDEN_header = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["bold_text"] | \
        formats["border_thin"] | \
        formats["BGColors"]["COMMON"]["header"])
    table_HIDDEN_header2 = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["bold_text"] | \
        formats["border_thick"] | \
        formats["BGColors"]["COMMON"]["header_2"])
    table_HIDDEN_disabled = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["border_thin"] | \
        formats["BGColors"]["COMMON"]["disabled"])
    table_HIDDEN_body = wb.add_format(formats["Arial8Regular"] | \
        formats["middle_center"] | \
        formats["border_thin"] | \
        formats["decimal_2"] | \
        formats["BGColors"]["COMMON"]["body"])
    A8RML_text = wb.add_format(formats["Arial8Regular"] | formats["middle_left"])
    A8RBC_text = wb.add_format(formats["Arial8Regular"] | formats["bottom_center"])
    A8RMC_text = wb.add_format(formats["Arial8Regular"] | formats["middle_center"])
    A8RMJ_text = wb.add_format(formats["Arial8Regular"] | formats["middle_just"])
    A8BML_text = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["middle_left"])
    A8BMC_text = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["middle_center"])
    A8BUMC_text = wb.add_format(formats["Arial8Regular"] | formats["bold_underline_text"] | formats["middle_center"])
    A8RMC_number = wb.add_format(formats["Arial8Regular"] | formats["middle_center"] | formats["decimal_2"])
    A8BBC_text = wb.add_format(formats["Arial8Regular"] | formats["bold_text"] | formats["bottom_center"])
    multiplier_text = wb.add_format(formats["multiplier"])

def load_data():
    data = None
    data1 = None
    data2 = None
    data3 = None
    data4 = None
    data5 = None
    data6 = None
    
    with open("./text_presets/formats.json") as dict_file:
        data1 = json.load(dict_file)
    
    with open("./text_presets/proposals.json", encoding="utf8") as dict_file:
        data2 = json.load(dict_file)
        
    with open("./text_presets/reps.json", encoding="utf8") as dict_file:
        data3 = json.load(dict_file)
    
    with open("./text_presets/signataries.json", encoding="utf8") as dict_file:
        data4 = json.load(dict_file)
        
    with open("./text_presets/header.json", encoding="utf8") as dict_file:
        data5 = json.load(dict_file)
        
    with open("./text_presets/statement.txt", encoding="utf8") as statement_file:
        data6 = {
            "STATEMENT": statement_file.read()
        }
    
    data = {**data1, **data2, **data3, **data4, **data5, **data6}
    
    return data
class PlaneelhaOutputer:
    def get_datetime_from_string(self, date_string, time_string):
        locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")
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
                self.diaSemana = date_object.strftime("%A").encode('utf-8').decode('utf-8')
                self.mesExtenso = date_object.strftime("%B").encode('utf-8').decode('utf-8')
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
        wb.formats[0].set_text_wrap()
        
        # Adjust the main column widths and the logo line height
        # ws.set_default_row(data["FORMATS"]["rowHeights"]["default"])
        ws.set_row_pixels(0, data["FORMATS"]["rowHeights"]["timbra"])
        ws.set_column_pixels(0, 0, data["FORMATS"]["colWidths"]["id"])
        ws.set_column_pixels(1, 1, data["FORMATS"]["colWidths"]["desc"])
        ws.set_column_pixels(2, 4, data["FORMATS"]["colWidths"]["muq"])
        ws.set_column_pixels(5, 12, data["FORMATS"]["colWidths"]["valores"])
    
    def set_database_sheet(self, wb, data):
        bd = wb.add_worksheet("BD")
        bd.set_column_pixels(0, 1, data["FORMATS"]["colWidths"]["valores"])
        bd.write_string("A1", "CHAVE", A8BMC_text)
        bd.write_string("A2", "ESTIMADO", A8RMC_text)
        bd.write_string("A3", "MÍNIMO", A8RMC_text)
        bd.write_string("A4", "REAJUSTADO", A8RMC_text)
        bd.write_string("B1", "VALOR", A8BMC_text)
        bd.write_string("B2", "H", A8RMC_text)
        bd.write_string("B3", "I", A8RMC_text)
        bd.write_string("B4", "K", A8RMC_text)

    def write_document_header(self, ws, data):
        formats = data["FORMATS"]
        header = data["HEADER"]
        proposal = data["PROPOSALS"][self.empresa]["PROPOSTA"]
        
        ws.insert_image(0, 1, "./images/" + self.empresa + ".png", data["FORMATS"]["logo"])   
        ws.merge_range("A2:G2", header["A2"], A8BML_text)
        ws.set_row_pixels(2, formats["rowHeights"]["orgao"])
        ws.merge_range("A3:G3", header["A3"].format(self.orgao.upper()), A8BMC_text)
        
        ws.write("A4", header["A4"], A8BML_text)
        ws.merge_range("B4:G4", "", A8BMC_text)
        
        bid = "{} Nº {} - ".format(self.tipo, self.codLicitacao)
        process = "" if self.codProcesso == "" else "PROCESSO LICITATÓRIO Nº {} - ".format(self.codProcesso)
        lictype = "Menor preço por {}".format("ITEM" if self.agrupamento == 0 else "LOTE")
        opening = "DATA E HORÁRIO DE ABERTURA: {}, {} de {} de {} às {} horas".format(self.diaSemana, self.dia, self.mesExtenso, self.ano, self.hora)
        ws.merge_range("A5:G5", header["A5"].format(proposal), A8BMC_text)
        ws.merge_range("A6:G6", bid + process + lictype, A8BMC_text)
        ws.merge_range("A7:G7", opening, A8BMC_text)
        
        ws.write("H1", "MULT. MÍNIMO:", A8BMC_text)
        ws.write_number("I1", 1.0, multiplier_text)

        ws.write("H2", "VISUALIZAR:", A8BMC_text)
        ws.data_validation("I2", {"validate": "list", "source": "=BD!$A$2:$A$4"})
        ws.write("I2", "ESTIMADO", table_HIDDEN_header2)
 
    def write_item_table(self, ws, data, title, start_line, item_count):
        marca = data["PROPOSALS"][self.empresa]["MARCA"]
        
        ws.set_row_pixels(start_line - 1, data["FORMATS"]["rowHeights"]["tabela_pontas"])
        ws.set_row_pixels(start_line + item_count, data["FORMATS"]["rowHeights"]["tabela_pontas"])
        
        ws.write("A{}".format(start_line), "ITEM", table_HEAD_text)
        ws.write("B{}".format(start_line), title, table_HEAD_text)
        ws.write("C{}".format(start_line), "MARCA", table_HEAD_text)
        ws.write("D{}".format(start_line), "UNID.", table_HEAD_text)
        ws.write("E{}".format(start_line), "QUANT.", table_HEAD_text)
        ws.write("F{}".format(start_line), "V.UNIT.", table_HEAD_text)
        ws.write("G{}".format(start_line), "V.TOTAL.", table_HEAD_text)
        ws.write("H{}".format(start_line), "ESTIMADO", table_HIDDEN_header)
        ws.write("I{}".format(start_line), "MÍNIMO", table_HIDDEN_header)
        ws.write("J{}".format(start_line), "CUSTO", table_HIDDEN_header)
        ws.write("K{}".format(start_line), "REAJUSTADO", table_HIDDEN_header)
        
        for i in range(item_count):
            line = start_line + i + 1
            ws.write("A{}".format(line), str(i + 1), table_IMUQ_text)
            ws.write("B{}".format(line), "", table_DESC_text)
            ws.write("C{}".format(line), marca, table_IMUQ_text)
            ws.write("D{}".format(line), "UN", table_IMUQ_text)
            ws.write("E{}".format(line), "", table_IMUQ_text)
            ws.write_formula("F{}".format(line), \
                "=TRUNC(INDIRECT(CONCATENATE(VLOOKUP($I$2,BD!$A$2:$B$4,2,FALSE),ROW())),2)", table_MONEYA8R_text)
            ws.write_formula("G{}".format(line), \
                "=E{}*F{}".format(line, line), table_MONEYA8R_text)
            ws.write("H{}".format(line), "", table_HIDDEN_body)
            ws.write_formula("I{}".format(line), \
                "=$I$1*J{}".format(line), table_HIDDEN_body)
            ws.write("J{}".format(line), "", table_HIDDEN_body)
            ws.write("K{}".format(line), "", table_HIDDEN_body)

        tipo_total = "GERAL" if title == "DESCRIÇÃO DO PRODUTO" else title
        line = start_line + item_count + 1
        first = start_line + 1
        last = start_line + item_count
        ws.merge_range("A{}:F{}".format(line, line), "TOTAL {}".format(tipo_total), table_TITLE_text)
        ws.write_formula("G{}".format(line), \
                "=SUM(G{}:G{})".format(first, last), table_MONEYA8B_text)
        ws.write("H{}".format(line), "", table_HIDDEN_disabled)
        ws.write("I{}".format(line), "", table_HIDDEN_disabled)
        ws.write("J{}".format(line), "", table_HIDDEN_disabled)
        ws.write("K{}".format(line), "", table_HIDDEN_disabled)

    def write_tables(self, ws, data):
        ws.set_row_pixels(8, data["FORMATS"]["rowHeights"]["tabela_pontas"])
        ws.merge_range("A9:G9", "RELAÇÃO DE ITENS", table_TITLE_text)
        ws.write("H9", "", table_HIDDEN_disabled)
        ws.write("I9", "", table_HIDDEN_disabled)
        ws.write("J9", "", table_HIDDEN_disabled)
        ws.write("K9", "", table_HIDDEN_disabled)
        if self.agrupamento == 0:
            self.write_item_table(ws, data, "DESCRIÇÃO DO PRODUTO", 10, self.qtd)
            return self.qtd + 11
        else:
            nome = "LOTE {}"
            item_sum = 0
            skipped = 0
            for lote in range(self.qtd):
                if self.lotesQtd[lote] == 0:
                    skipped += 1
                    continue
                start_line = 10 + item_sum + 2 * (lote - skipped)
                self.write_item_table(ws, data, nome.format(lote + 1), start_line, self.lotesQtd[lote])
                item_sum += self.lotesQtd[lote]
            line = 10 + item_sum + 2 * (self.qtd - skipped)
            if self.qtd - skipped > 1:
                ws.set_row_pixels(line -1, data["FORMATS"]["rowHeights"]["tabela_pontas"])
                ws.merge_range("A{0}:F{0}".format(line), "TOTAL GLOBAL", table_HEAD_text)
                ws.write_formula("G{}".format(line), "=SUMPRODUCT(--(ISNUMBER(SEARCH(\"TOTAL\",A11:A{0}))),G11:G{0})".format(line-1), table_MONEYA8B_text)
                ws.write("H{}".format(line), "", table_HIDDEN_disabled)
                ws.write("I{}".format(line), "", table_HIDDEN_disabled)
                ws.write("J{}".format(line), "", table_HIDDEN_disabled)
                ws.write("K{}".format(line), "", table_HIDDEN_disabled)
            return line

    def write_details(self, ws, data, start_line):
        proposal = data["PROPOSALS"][self.empresa]
        statement = data["STATEMENT"]
        rep = data["REPS"][self.empresa]
        sign = data["SIGNATARIES"]["CRISTIAN"]
        conditions = "- Condições de pagamento: {}".format(proposal["CONDICOES"])
        expires = "- Validade da proposta: {}".format(proposal["VALIDADE"])
        banks = []
        for bank_info in proposal["BANCOS"]:
            bank = "- Dados bancários - {}: Agência: {} / Conta: {}{}".format(\
                    bank_info["NOME"], bank_info["AG"], bank_info["CC"],\
                    "" if bank_info["PIX"] == "" else " / PIX: {}".format(bank_info["PIX"])\
                )
            banks.append(bank)
        adr = rep["ENDERECO"]
        rep_title = "REPRESENTANTE LEGAL PARA FINS DE ASSINATURA DE CONTRATO:"
        rep_name = "Nome: {}".format(rep["NOME"])
        rep_id = "Identidade: {} - Org. Expedidor: {}".format(rep["ID"], rep["ORG"])
        rep_cpf = "CPF: {}     Estado Civil: {}".format(rep["CPF"], rep["ECIVIL"])
        rep_end = "Endereço: {} - {} - {} - Cidade: {}-{}".format(\
            adr["LOGRADOURO"], adr["BAIRRO"], adr["CEP"], adr["CIDADE"], adr["ESTADO"])
        sign_place_time = "{}, {} de {} de {}".format(proposal["ENDERECO"]["CIDADE"],\
            self.dia, self.mesExtenso, self.ano)
        sign_underline = "_______________________________________________________"
        sign_name = sign["NOME"]
        sign_id = "Rep. Comercial - CPF: {} RG: {}".format(sign["CPF"], sign["RG"])
        
        ws.merge_range("A{}:G{}".format(start_line, start_line), conditions, A8RML_text)
        ws.merge_range("A{}:G{}".format(start_line + 1, start_line + 1), expires, A8RML_text)
        line = start_line + 2
        for bank in banks:
            ws.merge_range("A{}:G{}".format(line, line), bank, A8RML_text)
            line += 1
        statement_line = start_line + len(banks) + 3
        ws.set_row_pixels(statement_line - 1, data["FORMATS"]["rowHeights"]["statement"])
        ws.merge_range("A{}:G{}".format(statement_line, statement_line), statement, A8RMJ_text)
        
        rep_line = statement_line + 2
        ws.merge_range("A{}:G{}".format(rep_line, rep_line), rep_title, A8BMC_text)
        ws.merge_range("A{}:G{}".format(rep_line + 1, rep_line + 1), rep_name, A8RMC_text)
        ws.merge_range("A{}:G{}".format(rep_line + 2, rep_line + 2), rep_id, A8RMC_text)
        ws.merge_range("A{}:G{}".format(rep_line + 3, rep_line + 3), rep_cpf, A8RMC_text)
        ws.merge_range("A{}:G{}".format(rep_line + 4, rep_line + 4), rep_end, A8RMC_text)
        
        ws.set_row_pixels(rep_line + 4, data["FORMATS"]["rowHeights"]["rep_x_sign"])
        sign_line = rep_line + 6
        
        ws.merge_range("A{}:G{}".format(sign_line, sign_line), sign_place_time, A8BBC_text)
        ws.set_row_pixels(sign_line, data["FORMATS"]["rowHeights"]["sign"])
        ws.merge_range("A{}:G{}".format(sign_line + 1, sign_line + 1), sign_underline, A8BBC_text)
        ws.merge_range("A{}:G{}".format(sign_line + 2, sign_line + 2), sign_name, A8BBC_text)
        ws.merge_range("A{}:G{}".format(sign_line + 3, sign_line + 3), sign_id, A8BBC_text)
        
        return sign_line + 3

    def generate_file(self):
        data = load_data()
        
        opt = {
            "strings_to_numbers": True
        }

        wb = xlsx.Workbook(filename=self.arquivo, options=opt)
        init_format_variables(wb, data["FORMATS"], self.empresa)
        
        self.set_database_sheet(wb, data)

        ws = wb.add_worksheet("PROPOSTA")
        
        self.set_default_document_format(wb, ws, data)
        self.write_document_header(ws, data)
        last_table_line = self.write_tables(ws, data)
        last_document_line = self.write_details(ws, data, last_table_line + 2)
        
        ws.set_paper(9)
        ws.set_margins(1, 1, 1, 1)
        ws.fit_to_pages(1, 0)
        ws.center_horizontally()
        ws.print_area("A1:G{}".format(last_document_line))
        ws.repeat_rows(0)
        
        ws.activate()
        
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
        self.agrupamento = params["agrupamento"]
        self.lotesQtd = params["lotesQtd"]
        self.arquivo = params["caminhoArquivo"]
        
        self.generate_file()
        
        print("SUCCESS!")
        print(f"DATA E HORÁRIO DE ABERTURA: {self.dia} de {self.mesExtenso}, {self.diaSemana}, {self.hora} horas.")