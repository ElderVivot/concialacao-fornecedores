import os
import sys

absPath = os.path.dirname(os.path.abspath(__file__))

import xlsxwriter
from functions import getDateTimeNowInFormatStr, analyzeIfFieldIsValid

class GenerateExcel():
    def __init__(self, lancamentos):
        self._lancamentos = lancamentos
        self._codi_emp = self._lancamentos[0]['codi_emp']

        self._workbook = xlsxwriter.Workbook(self._createWayFileToSave())

        self._addFormatsCells()

    def _createWayFileToSave(self):
        return os.path.join(absPath[:absPath.find('src')], 'razao_pra_conciliacao', f'{self._codi_emp}_{getDateTimeNowInFormatStr()}.xlsx')

    def _addFormatsCells(self):
        self._cell_format_header_yellow = self._workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': 'yellow', 'text_wrap': True})
        self._cell_format_header_blue = self._workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': '#0070c0', 'text_wrap': True})
        self._cell_format_header_green = self._workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': '#a9d08e', 'text_wrap': True})
        self._cell_format_header_red = self._workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': 'red', 'text_wrap': True})
        self._cell_format_money = self._workbook.add_format({'num_format': '##0.00'})
        self._cell_format_money_saldo_anterior = self._workbook.add_format({'num_format': '##0.00"D";##0.00"C"'})
        self._cell_format_money_valor_consolidado = self._workbook.add_format({'num_format': '##0.00"C";##0.00"D"'})
        self._cell_format_date = self._workbook.add_format({'num_format': 'dd/mm/yyyy'})
        
    def createSheet(self):
        sheet = self._workbook.add_worksheet('conciliacao')
        sheet.freeze_panes(1, 0)
        sheet.set_column(7,7,options={'hidden':True})

        sheet.write(0, 0, "Cod. Emp", self._cell_format_header_yellow) #A
        sheet.write(0, 1, "Nome Fornecedor", self._cell_format_header_yellow) #B
        sheet.write(0, 2, "Conta", self._cell_format_header_yellow) #C
        sheet.write(0, 3, "Data", self._cell_format_header_yellow) #D
        sheet.write(0, 4, "Saldo Anterior", self._cell_format_header_yellow) #E
        sheet.write(0, 5, "Valor Cred", self._cell_format_header_yellow) #F
        sheet.write(0, 6, "Valor Deb", self._cell_format_header_yellow) #G
        # sheet.write(0, 7, "Pagamento antes da Provisão", self._cell_format_header_yellow)
        sheet.write(0, 7, "Valor Consolidado", self._cell_format_header_blue) #H
        sheet.write(0, 8, "Valor Provisão Fornecedor", self._cell_format_header_green) #I
        sheet.write(0, 9, "Valor Pago Fornecedor", self._cell_format_header_green) #J
        sheet.write(0, 10, "Saldo Atual Fornecedor", self._cell_format_header_red) #K
        sheet.write(0, 11, "Histórico", self._cell_format_header_yellow) #L
        sheet.write(0, 12, "NF", self._cell_format_header_blue) #M
        sheet.write(0, 13, "Valor Provisão Nota", self._cell_format_header_green) #N
        sheet.write(0, 14, "Valor Pago Nota", self._cell_format_header_green) #O
        sheet.write(0, 15, "Saldo Atual Nota", self._cell_format_header_red) #P
        sheet.write(0, 16, "Cod. Lote", self._cell_format_header_yellow) #Q

        for key, lancamento in enumerate(self._lancamentos):
            row = key+1
            row2 = key+2

            codi_emp = analyzeIfFieldIsValid(lancamento, "codi_emp")
            nome_fornecedor = analyzeIfFieldIsValid(lancamento, "nome_fornecedor")
            conta_fornecedor = analyzeIfFieldIsValid(lancamento, "conta_fornecedor")
            data_lancamento = analyzeIfFieldIsValid(lancamento, "data_lancamento")
            saldo_anterior = analyzeIfFieldIsValid(lancamento, "saldo_anterior")            
            valor_credito = analyzeIfFieldIsValid(lancamento, "valor_credito")
            valor_debito = analyzeIfFieldIsValid(lancamento, "valor_debito")
            formula_valor_consolidado = f'=IF(G{row2}>0,G{row2}*-1,IF(F{row2}=0,E{row2}*(-1),F{row2}))'
            formula_valor_provisao_fornecedor = f'=SUMIFS(F:F,C:C,C{row2})'
            formula_valor_pago_fornecedor = f'=SUMIFS(G:G,C:C,C{row2})'
            formula_valor_consolidado_fornecedor = f'=SUMIFS(H:H,C:C,C{row2})'
            historico = analyzeIfFieldIsValid(lancamento, "historico")
            number_note = analyzeIfFieldIsValid(lancamento, "numberNote")
            formula_valor_provisao_nota = f'=SUMIFS(F:F,C:C,C{row2},M:M,M{row2})'
            formula_valor_pago_nota = f'=SUMIFS(G:G,C:C,C{row2},M:M,M{row2})'
            formula_valor_consolidado_nota = f'=SUMIFS(H:H,C:C,C{row2},M:M,M{row2})'
            codi_lote = analyzeIfFieldIsValid(lancamento, "codi_lote")
            
            sheet.write(row, 0, codi_emp)
            sheet.write(row, 1, nome_fornecedor)
            sheet.write(row, 2, conta_fornecedor)
            sheet.write(row, 3, data_lancamento, self._cell_format_date)
            sheet.write(row, 4, saldo_anterior, self._cell_format_money_saldo_anterior)
            sheet.write(row, 5, valor_credito, self._cell_format_money)
            sheet.write(row, 6, valor_debito, self._cell_format_money)
            sheet.write_formula(row, 7, formula_valor_consolidado, self._cell_format_money)
            sheet.write_formula(row, 8, formula_valor_provisao_fornecedor, self._cell_format_money)
            sheet.write_formula(row, 9, formula_valor_pago_fornecedor, self._cell_format_money)
            sheet.write_formula(row, 10, formula_valor_consolidado_fornecedor, self._cell_format_money_valor_consolidado)
            sheet.write(row, 11, historico)
            sheet.write(row, 12, number_note)
            sheet.write_formula(row, 13, formula_valor_provisao_nota, self._cell_format_money)
            sheet.write_formula(row, 14, formula_valor_pago_nota, self._cell_format_money)
            sheet.write_formula(row, 15, formula_valor_consolidado_nota, self._cell_format_money_valor_consolidado)
            sheet.write(row, 16, codi_lote)

    def closeFile(self):
        self._workbook.close()
        