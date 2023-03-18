from datetime import datetime
import os
import sys
import re
import json

import PyPDF2


def float_by_string(strNumber) -> float:
    return float(re.findall("[\d\.,]+", strNumber)[0].replace(".", "").replace(",", "."))


class readFGTSPDF:
    def __init__(self, root) -> None:
        self.fgtss = []
        self.root = root
        if(not os.path.exists(self.root)):
            os.mkdir(self.root)

    def header(self, fulltext):
        fields = re.findall(
            'EMPREGADOR(.*)CARTEIRA DE TRABALHO(.*)DATA DE OPÇÃO(.*)TIPO DE CONTA(.*)DATA DE ADMISSÃO(.*)INCRIÇÃO DO EMPREGADOR(.*)' \
            + 'DATA E CÓDIGO DE AFASTAMENTO(.*)TAXA DE JUROS(.*)PIS/PASEP(.*)Nº DA CONTA \(COD. ESTABELECIMENTO/CONTA\)(.*)CATEGORIA(.*)VALOR PARA FINS RECISÓRIOS(.*)Histórico de Movimentaçõe',
            fulltext[:500].replace("\n", " ").replace("\xa0", ""))
        fields = [x.strip() for x in fields[0]]
        return {
            'EMPREGADOR': fields[0],
            'CARTEIRA DE TRABALHO': fields[1],
            'DATA DE OPÇÃO': fields[2],
            'TIPO DE CONTA': fields[3],
            'DATA DE ADMISSÃO': fields[4],
            'INCRIÇÃO DO EMPREGADOR': fields[5],
            "DATA E CÓDIGO DE AFASTAMENTO": fields[6],
            'TAXA DE JUROS': fields[7],
            'PIS/PASEP': fields[8],
            'Nº DA CONTA': fields[9],
            'CATEGORIA': fields[10],
            'VALOR PARA FINS RECISÓRIOS': fields[11],
        }

    def jamstract(self, move, conta, typeLine, linhaAnterior=None):
        move = move.replace("\n", " ").replace("\xa0", "")
        fields = re.findall(
            "(\d{2}/\d{2}/\d{4})[ ]+CREDITO DE JAM ([\d\.\,-]+)[ ]+R\$[ ]+([\d\.\,-]+)[ ]+R\$[ ]+([\d\.\,-]+)", move)
        fields = [x.strip() for x in fields[0]]
        return {
            'Descrição da Movimentação': "JUROS AO MES APLICADO EM " + fields[0],
            'Tipo Movimentação': typeLine,
            # 'conta':conta,
            'Data Movimentação': "01{}".format(fields[0][2:]),
            "Base de Calculo": linhaAnterior["Acumulado"] if linhaAnterior else 0,
            'Juros Aplicado': float_by_string(fields[1]),
            'TR da Caixa': float_by_string(fields[1]) - 0.002466,
            'Resultado dos Juros': float_by_string(fields[2]),
            'Acumulado': float_by_string(fields[3])
        }

    def depositostract(self, move, conta, typeLine):
        move = move.replace("\n", " ").replace("\xa0", "")
        fields = re.findall("(\d{2}/\d{2}/\d{4})([^\$]+)R\$[ ]+([\d\.\,-]+)[ ]+R\$[ ]+([\d\.\,-]+)", move)
        fields = [x.strip() for x in fields[0]]
        return {
            'Descrição da Movimentação': fields[1],
            'Tipo Movimentação': typeLine,
            # 'conta':conta,
            'Data Movimentação': fields[0],
            'Deposito Empresa': float_by_string(fields[2]),
            'Acumulado': float_by_string(fields[3])
        }

    def extract(self):

        for f in os.listdir(self.root):
            if (f[f.rfind("."):] != ".pdf"):
                continue
            reader = PyPDF2.PdfReader(os.path.join(self.root, f))
            fulltext = "\n\n".join([x.extract_text() for x in reader.pages])
            objFGTS = self.header(fulltext)
            movimentations = re.findall('\d{2}/\d{2}/\d{4}\n[^\n]+\nR\$[ -]+\xa0[^\n]+\nR\$ \xa0[^\n]+\n', fulltext)
            companilines = []
            for move in movimentations:
                if ("CREDITO DE JAM" in move):
                    linhaAnterior = self.jamstract(move, objFGTS['Nº DA CONTA'], 'jam', linhaAnterior)
                    companilines.append(linhaAnterior)
                elif ("DEPOSITO" in move):
                    linhaAnterior = self.depositostract(move, objFGTS['Nº DA CONTA'], 'deposito')
                    companilines.append(linhaAnterior)
                elif ("RESULTADO ANO BASE" in move):
                    linhaAnterior = self.depositostract(move, objFGTS['Nº DA CONTA'], 'plr')
                    companilines.append(linhaAnterior)
                elif ("SAQUE" in move):
                    linhaAnterior = self.depositostract(move, objFGTS['Nº DA CONTA'], 'saque')
                    companilines.append(linhaAnterior)
                else:
                    linhaAnterior = self.depositostract(move, objFGTS['Nº DA CONTA'], 'other')
                    companilines.append(linhaAnterior)
            objFGTS["lines"] = companilines
            self.fgtss.append(objFGTS)

            with open(os.path.join(self.root, 'fgts conta_{}.json').format(objFGTS['Nº DA CONTA'].replace("/", "-")),
                      'w') as f:
                json.dump(objFGTS, f)

        with open(os.path.join(self.root, 'fgts.json'), 'w') as f:
            json.dump(self.fgtss, f)

        return self.fgtss


import xlsxwriter
import csv


class writeExcel():

    def __init__(self, root, objFGTS) -> None:

        self.root = root
        self.workbook = xlsxwriter.Workbook(os.path.join(os.path.dirname(root),'fgts.xlsx'))

        self.money = self.workbook.add_format({'num_format': 'R$ #,##0.00;-R$ #,##0.00'})
        self.percent = self.workbook.add_format({'num_format': '0.0000%'})
        self.bold = self.workbook.add_format({'bold': True, 'align': 'right'})
        self.boldMoney = self.workbook.add_format(
            {'bold': True, 'num_format': 'R$ #,##0.00;-R$ #,##0.00', 'align': 'left'})

        self.columns = ['Data Movimentação', 'Descrição da Movimentação', 'Tipo Movimentação', 'Deposito Empresa',
                        'Base de Calculo', 'Juros Aplicado', 'TR da Caixa', 'Resultado dos Juros', 'Acumulado']
        self.calculateColumns = [
            {'name': 'Juros INPC',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.IFNA(VLOOKUP(_xlfn.NUMBERVALUE(_xlfn.IF({fgtsTable}[[#This Row],[Tipo Movimentação]]="jam",_xlfn.CONCAT(_xlfn.YEAR({fgtsTable}[[#This Row],[Data Movimentação]]),_xlfn.RIGHT(_xlfn.CONCAT("0",_xlfn.MONTH({fgtsTable}[[#This Row],[Data Movimentação]])),2)),0)),\'Novos Juros\'!$A$1:$B$400,2,0)/100,"NJAM")',
             'columnType': self.percent},
            {'name': 'INPC + (0,2466%)',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.IF({fgtsTable}[[#This Row],[Juros INPC]] <> "NJAM",{fgtsTable}[[#This Row],[Juros INPC]]+0.002466,"")',
             'columnType': self.percent},
            {'name': 'Movimentações no Período',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.IF({fgtsTable}[[#This Row],[Tipo Movimentação]]<>"jam",{fgtsTable}[[#This Row],[Deposito Empresa]],0)',
             'columnType': self.money},
            {'name': 'Acumulado com INPC',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.TRUNC(_xlfn.IF({fgtsTable}[[#This Row],[Tipo Movimentação]]="jam",{"M" + str(row)}*(1+{fgtsTable}[[#This Row],[INPC + (0,2466%)]]),_xlfn.IFERROR({fgtsTable}[[#This Row],[Movimentações no Período]]+{"M" + str(row)},{fgtsTable}[[#This Row],[Movimentações no Período]])),2)',
             'columnType': self.money},
            {'name': 'Diferenaça entre indices',
             'formula': lambda row,
                               fgtsTable: f'={fgtsTable}[[#This Row],[Acumulado com INPC]]-{fgtsTable}[[#This Row],[Acumulado]]',
             'columnType': self.money},
            {'name': 'Movimentações no Período Sem Saques',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.IF(_xlfn.AND({fgtsTable}[[#This Row],[Tipo Movimentação]]<>"saque",{fgtsTable}[[#This Row],[Movimentações no Período]]>0),{fgtsTable}[[#This Row],[Movimentações no Período]],0)',
             'columnType': self.money},
            {'name': 'Acumulado com INPC Sem Saques',
             'formula': lambda row,
                               fgtsTable: f'=_xlfn.TRUNC(_xlfn.IF({fgtsTable}[[#This Row],[Tipo Movimentação]]="jam",{"P" + str(row)}*(1+{fgtsTable}[[#This Row],[INPC + (0,2466%)]]),_xlfn.IFERROR({fgtsTable}[[#This Row],[Movimentações no Período Sem Saques]]+{"P" + str(row)},{fgtsTable}[[#This Row],[Movimentações no Período Sem Saques]])),2)',
             'columnType': self.money},
        ]
        self.headers = [{'header': x} for x in self.columns]
        self.headers += [{'header': x["name"]} for x in self.calculateColumns]

        self.workbook.add_worksheet("Conclusão")
        for fgts in objFGTS:
            self.workbook.add_worksheet(self.workSheetName(fgts))
        self.workbook.add_worksheet("Novos Juros")

    def novosJuros(self, workbook):
        worksheet = workbook.get_worksheet_by_name("Novos Juros")
        row = 0
        with open(os.path.join(os.path.dirname(__file__),'inpc.csv'), 'r') as data:
            for line in csv.reader(data, delimiter='\t', ):
                if (line[0] == 'Data'):
                    worksheet.write(row, 0, line[0])
                    worksheet.write(row, 1, line[1])
                else:
                    worksheet.write(row, 0, int(line[0]))
                    worksheet.write(row, 1, float_by_string(line[1]))
                row += 1

    def conclusionTotals(self, AllTotals):
        workbook = self.workbook
        worksheet = workbook.get_worksheet_by_name("Conclusão")

        row = 1
        conclusionHeaders = ["EMPREGADOR", 'Nº DA CONTA', "VALOR PARA FINS RECISÓRIOS CAIXA",
                             "VALOR PARA FINS RECISÓRIOS CORRIGIDOS", "CORREÇÂO DE VALORES RESCISSORIOS",
                             "RESIDUAL SAQUE INATIVO"]

        worksheet.add_table(0, 0, len(AllTotals), len(conclusionHeaders) - 1, {
            'name': "conclusao",
            "columns": [{'header': x} for x in conclusionHeaders]
        })
        for line in AllTotals:
            col = 0
            worksheet.write(row, col, line["ObjFGTSEmpregador"]["EMPREGADOR"])
            col += 1
            worksheet.write(row, col, line["ObjFGTSEmpregador"]['Nº DA CONTA'])
            col += 1
            worksheet.write(row, col, float_by_string(line["ObjFGTSEmpregador"]["VALOR PARA FINS RECISÓRIOS"]),
                            self.money)
            col += 1
            worksheet.write_formula(row, col,
                                    f"='{self.workSheetName(line['ObjFGTSEmpregador'])}'!P{str(line['row'] + 1)}",
                                    self.money)
            col += 1
            worksheet.write_formula(row, col,
                                    f"=conclusao[[#This Row],[VALOR PARA FINS RECISÓRIOS CORRIGIDOS]]-conclusao[[#This Row],[VALOR PARA FINS RECISÓRIOS CAIXA]]" if
                                    line['ObjFGTSEmpregador'][
                                        "VALOR PARA FINS RECISÓRIOS"].strip() != "R$ 0,00" else '=0+0', self.money)
            col += 1
            worksheet.write_formula(row, col,
                                    f"='{self.workSheetName(line['ObjFGTSEmpregador'])}'!P{str(line['row'])}" if
                                    line['ObjFGTSEmpregador'][
                                        "VALOR PARA FINS RECISÓRIOS"].strip() == "R$ 0,00" else '=0+0', self.money)
            col += 1
            row += 1

        rowResume = row + 1
        worksheet.write_formula(rowResume, 5, f"=_xlfn.SUM(E2:E{row})", self.boldMoney)
        worksheet.merge_range(rowResume, 0, rowResume, 4, "Valor total Corrigir para Fins Recisórios", self.bold)
        rowResume += 1
        worksheet.write_formula(rowResume, 5, f"=_xlfn.SUM(F2:F{row})", self.boldMoney)
        worksheet.merge_range(rowResume, 0, rowResume, 4, "Valor total para Saque de Contas Inativas", self.bold)
        rowResume += 1
        worksheet.write_formula(rowResume, 5, f"=_xlfn.SUM(E2:F{row})", self.boldMoney)
        worksheet.merge_range(rowResume, 0, rowResume, 4, "Valor da Ação", self.bold)

    def workSheetName(self, fgtsObj) -> str:
        return fgtsObj['EMPREGADOR'][0:31].replace("/", "-")

    def write(self, objFGTS):
        workbook = self.workbook
        AllTotals = []
        for contafgts in objFGTS:
            row = 0
            col = 0
            worksheet = workbook.get_worksheet_by_name(self.workSheetName(contafgts))  # workbook.add_worksheet()

            for headersEmpregado in contafgts.keys():
                if (headersEmpregado == 'lines'):
                    continue

                worksheet.write(row, col, headersEmpregado)
                col += 1
                worksheet.write(row, col, contafgts[headersEmpregado])
                col += 1

            row += 1
            col = 0

            fgtsTable = "".join(contafgts['EMPREGADOR'].replace("/", "-").replace("-", "").split(" "))
            lines = contafgts["lines"]
            worksheet.add_table(1, 0, len(lines) + 1, len(self.headers) - 1, {
                'name': fgtsTable,
                "columns": self.headers
            })
            row += 1
            index=indexCalculate=0
            for line in lines:
                for index, columns in enumerate(self.columns):
                    if (not line.get(columns)):
                        continue;
                    if (columns in ['Deposito Empresa', 'Base de Calculo', 'Resultado dos Juros', 'Acumulado']):
                        worksheet.write(row, index, line[columns], self.money)
                    elif (columns in ['Juros Aplicado', 'TR da Caixa']):
                        worksheet.write(row, index, line[columns], self.percent)

                    else:
                        worksheet.write(row, index, line[columns])

                for indexCalculate, columns in enumerate(self.calculateColumns):
                    worksheet.write_formula(row, index + 1 + indexCalculate, columns["formula"](row, fgtsTable),
                                            columns["columnType"])

                row += 1

            conclusionRow = row + 2
            # worksheet.write(conclusionRow, index+indexCalculate+0,
            # worksheet.write(conclusionRow+1, index+indexCalculate+0,
            worksheet.merge_range('A{row}:O{row}'.format(row=conclusionRow + 1),
                                  "Valor possível Saque ou Correção para fins Recisórios", self.bold)
            worksheet.merge_range('A{row}:O{row}'.format(row=conclusionRow + 2),
                                  "Valor de correção para fins Recisórios", self.bold)

            worksheet.write_formula(conclusionRow, index + indexCalculate + 1, "=N" + str(row), self.boldMoney)
            worksheet.write_formula(conclusionRow + 1, index + indexCalculate + 1, "=P" + str(row) if contafgts[
                                                                                                          "VALOR PARA FINS RECISÓRIOS"].strip() != "R$ 0,00" else '=0+0',
                                    self.boldMoney)

            AllTotals.append({'ObjFGTSEmpregador': contafgts, "row": conclusionRow + 1})

        self.novosJuros(workbook)

        self.conclusionTotals(AllTotals)

        workbook.close()


if __name__ == "__main__":
    root = os.path.join(os.path.dirname(__file__),"PDF_FGTS")
    print(root)
    fgtsObj = readFGTSPDF(root).extract() # json.load(open(os.path.join(root, 'fgts.json')))  # None # 
    writeExcel(root=root, objFGTS=fgtsObj).write(objFGTS=fgtsObj)
