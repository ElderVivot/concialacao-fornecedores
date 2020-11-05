import os
import sys

absPath = os.path.dirname(os.path.abspath(__file__))

from functions import readExcel, removerAcentosECaracteresEspeciais, analyzeIfFieldIsValidMatrix

class ReadConfiguracoes():
    def __init__(self):
        self._wayConfiguracoes = os.path.join(absPath, '..', 'configuracoes.xlsx')

    def process(self):
        keyWords = []

        dataExcel = readExcel(self._wayConfiguracoes)
        for key, data in enumerate(dataExcel):
            if key == 0:
                continue

            keyWord = removerAcentosECaracteresEspeciais(analyzeIfFieldIsValidMatrix(data, 1)).upper()
            if keyWord.strip() != '': 
                keyWords.append(keyWord)
        
        return keyWords

if __name__ == "__main__":
    readConfiguracoes = ReadConfiguracoes()
    readConfiguracoes.process()