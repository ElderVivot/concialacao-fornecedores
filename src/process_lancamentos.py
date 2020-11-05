from functions import treatNumberField

class ProcessLancamentos():
    def __init__(self, lancamentos, configuracoes):
        self._lancamentos = lancamentos
        self._configuracoes = configuracoes

    def getNumberNote(self, lancamento):
        numberNote = ''
        historico: str = lancamento['historico']
        for configuracao in self._configuracoes:
            if configuracao in historico:
                positionNumberNote = historico.find(configuracao) + len(configuracao)
                numberNotePraFrente = historico[positionNumberNote:].strip()
                numberNote = numberNotePraFrente.split(' ')[0].split('/')[0].split('-')[0]
                numberNote = treatNumberField(numberNote, isInt=True)

                return str(numberNote)

        return numberNote

    def process(self):
        lancamentos = []
        for lancamento in self._lancamentos:
            lancamento['numberNote'] = self.getNumberNote(lancamento)
            lancamentos.append(lancamento)

        return lancamentos