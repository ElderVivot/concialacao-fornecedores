import os
from functions import readExcel, treatTextFieldInVector, treatNumberFieldInVector, treatDateFieldInVector, treatDecimalFieldInVector

class ReadExcelDominio():
    def __init__(self):
        self._wayRazaoDominio = 'razao_dominio'

    def process(self, wayFile):
        lancamentos = []
        lancamentoDoRazao = {}

        dataExcel = readExcel(wayFile)
        
        for data in dataExcel:
            lancamentoDoRazao['codi_emp'] = treatNumberFieldInVector(data, 1, isInt=True)
            lancamentoDoRazao['nome_fornecedor'] = treatTextFieldInVector(data, 3)
            lancamentoDoRazao['conta_fornecedor'] = treatNumberFieldInVector(data, 4, isInt=True)
            lancamentoDoRazao['data_lancamento'] = treatDateFieldInVector(data, 6)    
            lancamentoDoRazao['saldo_anterior'] = treatDecimalFieldInVector(data, 8)
            lancamentoDoRazao['valor_debito'] = treatDecimalFieldInVector(data, 12)
            lancamentoDoRazao['valor_credito'] = treatDecimalFieldInVector(data, 13)
            lancamentoDoRazao['historico'] = treatTextFieldInVector(data, 18)
            lancamentoDoRazao['codi_lote'] = treatTextFieldInVector(data, 22)
            if lancamentoDoRazao['conta_fornecedor'] > 0 and lancamentoDoRazao['data_lancamento'] is not None:
                lancamentos.append(lancamentoDoRazao.copy())

        return lancamentos