import os
import sys

absPath = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(absPath, '..'))

from operator import itemgetter
from read_excel_dominio import ReadExcelDominio
from read_configuracoes import ReadConfiguracoes
from process_lancamentos import ProcessLancamentos
from generate_excel import GenerateExcel

class MainConciliacaoFornecedores():
    def __init__(self, dtStart: str, dtEnd: str):
        self._dtStart = dtStart
        self._dtEnd = dtEnd
        self._wayRazaoDominio = 'razao_dominio'
        self._readExcelDominio = ReadExcelDominio()
        self._readConfiguracoes = ReadConfiguracoes()
        self._configuracoes = self._readConfiguracoes.process()

        # self._userDominio = input('- Informe seu usuário na Domínio: ').

    def process(self, wayFile):
        lancamentos = self._readExcelDominio.process(wayFile)

        process_lancamentos = ProcessLancamentos(lancamentos, self._configuracoes)
        lancamentos = process_lancamentos.process()
        
        lancamentos = sorted(lancamentos, key=itemgetter('conta_fornecedor', 'numberNote', 'data_lancamento'))

        generate_excel = GenerateExcel(lancamentos)
        generate_excel.createSheet()
        generate_excel.closeFile()

    def processAll(self):
        for root, _, files in os.walk(self._wayRazaoDominio):
            for file in files:
                if file.lower().endswith(('.xls', '.xlsx')):
                    wayFile = os.path.join(root, file)
                    self.process(wayFile)


if __name__ == "__main__":
    main = MainConciliacaoFornecedores('', '')
    main.processAll()