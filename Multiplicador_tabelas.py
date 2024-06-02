import pandas as pd

def expand_records(input_file, output_file, quantity_column):
    # Ler a tabela do Excel
    df = pd.read_excel(input_file)

    # Expandir os registros
    expanded_data = []
    for index, row in df.iterrows():
        quantity = row[quantity_column]
        for _ in range(quantity):
            expanded_data.append(row)

    # Criar um novo DataFrame com os dados expandidos
    expanded_df = pd.DataFrame(expanded_data)

    # Salvar o novo DataFrame em um novo arquivo Excel
    expanded_df.to_excel(output_file, index=False)

def main():
    # Solicitar os detalhes ao usuário
    input_file = input("Digite o caminho do arquivo Excel de entrada inclusive o nome do arquivo.xlsx: ")
    output_file = input("Digite o caminho do arquivo Excel de saída inclusive o nome do arquivo.xlsx: ")
    quantity_column = input("Digite o nome da coluna das quantidades: ")

    # Executar a expansão dos registros
    expand_records(input_file, output_file, quantity_column)
    print(f"Arquivo processado e salvo em {output_file}")

if __name__ == "__main__":
    main()
