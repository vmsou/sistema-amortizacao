import tabelas


def main():
    dados = {
        'tipo': 'price',
        'financiamento': 300_000,
        'emprestimo_parcelas': 2,  # 1 - sem parcela, 2 - duas vezes, 3 - tres...
        'carencia': 3,  # 0 - nenhum
        'carencia_tipo': 'nenhum',  # nenhum, total, parcial
        'parcelas': 5,
        'antecipada': 0,  # 1 - antecipada, 0 - postecipada
        'taxa': 0.12,
        'limitado': 0,  # 0 - nenhum,
    }

    tabela = tabelas.Tabela(**dados)
    # tabela.print_table()
    tabela.create_sheet(tabela, filename='test')


if __name__ == '__main__':
    main()

