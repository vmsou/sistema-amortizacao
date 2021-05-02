from dataclasses import dataclass
from collections import defaultdict
import numpy_financial
import tabela_excel


@dataclass
class Tabela:

    tipo: str  # SAC ou Price
    financiamento: float  # Valor monetario Ex: 300_000
    emprestimo_parcelas: None
    carencia: int  # Ex: 3 anos
    carencia_tipo: str  # Ex: Total
    parcelas: int  # Ex: 5 anuais
    antecipada: int  # 1 - antecipada, 0 - postecipada
    taxa: float  # Ex: 0.12 -> 12%
    limitado: int  # Ex: 5 anos

    _valores = defaultdict(dict)

    def __post_init__(self):
        self._initialize_table()
        self._adjust_carencia()
        self._initialize_type()
        self._check_state()

    def _initialize_table(self):
        if self.limitado:
            for i in range(self.limitado + 1):
                self._valores[i] = {'saldo': 0, 'amort': 0, 'juros': 0, 'prest': 0}
        elif self.carencia > 1:
            for i in range(self.parcelas + self.carencia):
                self._valores[i] = {'saldo': 0, 'amort': 0, 'juros': 0, 'prest': 0}
        else:
            for i in range(self.parcelas + self.carencia + 1):
                self._valores[i] = {'saldo': 0, 'amort': 0, 'juros': 0, 'prest': 0}
        self._valores['total'] = {'saldo': 0, 'amort': 0, 'juros': 0, 'prest': 0}

        self._primeira_linha()

    def _initialize_type(self):
        if self.tipo.lower() == 'sac':
            self._sac()
        elif self.tipo.lower() == 'price':
            self._price()

    def _primeira_linha(self):
        if isinstance(self.emprestimo_parcelas, int):
            valor = self.financiamento / self.emprestimo_parcelas
            self._valores[0]['saldo'] = float(valor)
        elif isinstance(self.emprestimo_parcelas, list):
            valor = self.emprestimo_parcelas[0]
            self._valores[0]['saldo'] = valor
        self._valores[0]['amort'] = 'x'
        self._valores[0]['juros'] = 'x'
        self._valores[0]['prest'] = 'x'

    def _ultima_linha(self):
        amort_total, juros_total, prest_total = 0, 0, 0

        for _, n in self._valores.items():
            amort = n['amort']
            juros = n['juros']
            prest = n['prest']

            if amort != 'x':
                amort_total += amort
            if juros != 'x':
                juros_total += juros
            if prest != 'x':
                prest_total += prest

        self._change_value('total', 'amort', amort_total)
        self._change_value('total', 'juros', juros_total)
        self._change_value('total', 'prest', prest_total)

    def _adjust_carencia(self):
        if self.carencia == 0:
            self.carencia = 1

    def _change_value(self, row, name, value):
        self._valores[row][name] = value

    def _sac(self):
        if self.carencia > 1:
            self._check_carencia()
            self._carencia_sac()
        else:
            if self.carencia_tipo != 'nenhum':
                print(f'There is no carencia, but you choose {self.carencia_tipo}.\nBut no worries we already fixed it!')
                self.carencia_tipo = 'nenhum'

        amort = self._get_amort_sac()

        for n in range(self.carencia, self.parcelas + self.carencia):

            self._change_value(n, 'amort', amort)
            self._get_saldo(n)
            self._get_juros(n)
            self._get_prest_sac(n)

            if n == self.limitado:
                break

        self._ultima_linha()

    def _price(self):

        if self.carencia > 1:
            self._check_carencia()
            self._carencia_price()
        else:
            if self.carencia_tipo != 'nenhum':
                print(f'There is no carencia, but you choose {self.carencia_tipo}.\nBut no worries we already fixed it!')
                self.carencia_tipo = 'nenhum'

        prest = self._get_prest_price()

        for n in range(self.carencia, self.parcelas + self.carencia):
            self._change_value(n, 'prest', prest)
            self._get_juros(n)
            self._get_amort_price(n)
            self._get_saldo(n)

        self._ultima_linha()

    def _get_saldo(self, row):
        self._change_value(row, 'saldo', self._valores[row - 1]['saldo'] - self._valores[row]['amort'])

    def _get_juros(self, row):
        if self.antecipada and row == self.carencia:
            juros = 0
        else:
            juros = self._valores[row - 1]['saldo'] * self.taxa
        self._change_value(row, 'juros', juros)

    def _check_carencia(self):
        if self.carencia_tipo not in ['parcial', 'total']:
            print('You probably forgot to choose carencia_tipo.')
            while self.carencia_tipo not in ['parcial', 'total']:
                tipo = str(input('carencia_tipo(parcial/total):'))
                self.carencia_tipo = tipo

    # SAC
    def _carencia_sac(self):

        if self.carencia_tipo.lower() == 'total':
            for n in range(1, self.carencia):
                if self.emprestimo_parcelas > 1 and n < self.emprestimo_parcelas:
                    valor = self.financiamento / self.emprestimo_parcelas
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + valor + (self._valores[n - 1]['saldo'] * self.taxa))
                else:
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + (self._valores[n - 1]['saldo'] * self.taxa))
                self._change_value(n, 'amort', 'x')
                self._change_value(n, 'juros', 'x')
                self._change_value(n, 'prest', 'x')
        elif self.carencia_tipo.lower() == 'parcial':
            for n in range(1, self.carencia):
                self._change_value(n, 'amort', 'x')
                if self.emprestimo_parcelas > 1 and n < self.emprestimo_parcelas:
                    valor = self.financiamento / self.emprestimo_parcelas
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + valor)
                else:
                    self._change_value(n, 'saldo', self.financiamento)
                self._change_value(n, 'juros', self._valores[n - 1]['saldo'] * self.taxa)
                self._change_value(n, 'prest', self._valores[n]['juros'])
        else:
            print('You probably forgot to choose carencia_tipo.')
            while self.carencia_tipo not in ['parcial', 'total']:
                self.carencia_tipo = str(input('carencia_tipo(parcial/total):'))

    def _get_amort_sac(self):
        if self.limitado:
            return self._valores[self.carencia - 1]['saldo'] / (self.limitado - (self.carencia - 1))
        else:
            return self._valores[self.carencia - 1]['saldo'] / self.parcelas

    def _get_prest_sac(self, row):
        self._change_value(row, 'prest', self._valores[row]['amort'] + self._valores[row]['juros'])

    # Price
    def _carencia_price(self):
        if self.carencia_tipo == 'parcial':
            for n in range(1, self.carencia):
                if self.emprestimo_parcelas > 1 and n < self.emprestimo_parcelas:
                    valor = self.financiamento / self.emprestimo_parcelas
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + valor)
                else:
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'])
                self._change_value(n, 'amort', 'x')
                juros = self._valores[n - 1]['saldo'] * self.taxa
                self._change_value(n, 'juros', juros)
                self._change_value(n, 'prest', juros)
        elif self.carencia_tipo == 'total':
            for n in range(1, self.carencia):
                if self.emprestimo_parcelas > 1 and n < self.emprestimo_parcelas:
                    valor = self.financiamento / self.emprestimo_parcelas
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + valor + (self._valores[n - 1]['saldo'] * self.taxa))
                else:
                    self._change_value(n, 'saldo', self._valores[n - 1]['saldo'] + (self._valores[n - 1]['saldo'] * self.taxa))
                self._change_value(n, 'amort', 'x')
                self._change_value(n, 'juros', 'x')
                self._change_value(n, 'prest', 'x')

    def _get_prest_price(self):
        return -(numpy_financial.pmt(self.taxa, self.parcelas, self._valores[self.carencia - 1]['saldo'], when=self.antecipada))

    def _get_amort_price(self, row):
        prest = self._get_prest_price()
        juros = self._valores[row]['juros']
        amort = prest - juros
        self._change_value(row, 'amort', amort)
        return amort

    def _check_state(self):

        saldo = self._valores[self.carencia - 1]['saldo']
        amort = self._valores['total']['amort']
        juros = self._valores['total']['juros']
        prest = self._valores['total']['prest']

        if (prest - (amort + juros)) > 2:
            print('Something went wrong.')
            print('Total prest:', prest, 'is different than total amort + total juros', amort + juros)
        if (saldo - amort) > 2:
            print('Something went wrong.')
            print(saldo, 'is different than saldo total', self._valores['total']['saldo'])

    def create_table(self):
        return self._valores

    def print_table(self):
        for n, values in self._valores.items():
            print(f"{n}:", end=' ')
            for name, value in values.items():
                print(f"{name.title()}:", end=' ')
                if value != 'x':
                    print(f"{round(value, 2)}", end=' ')
                else:
                    print(value, end=' ')
            print('')

    def create_sheet(self, tabela, row_start=2, column_start=2, filename='Sheet'):
        tabela_excel.run(tabela, row_start, column_start, filename)
