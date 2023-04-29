import os
import tabula
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# pasta onde estão os arquivos PDF
pdf_folder = r'C:\\Users\\felip\\AppData\\Local\\Programs\\Python\\Python311\\Scripts\\Recife'

# lista de arquivos PDF na pasta
pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

# função que processa cada arquivo PDF
def process_pdf(pdf_file):
    pdf_path = os.path.join(pdf_folder, pdf_file)
    # extrai a tabela da primeira página usando tabula-py
    table = tabula.read_pdf(pdf_path, pages='1')
    # extrai as colunas da tabela do PDF
    pdf_cols = table[0][table[0].columns[0]]
    pdf_vals = table[0][table[0].columns[1]]
    # cria um dicionário para armazenar as informações do PDF
    pdf_dict = {}
    for i in range(len(pdf_cols)):
        col_name = pdf_cols[i]
        col_val = pdf_vals[i]
        if col_name != "Total Devido Pelo Reclamado":
            pdf_dict[col_name] = col_val
    # adiciona a coluna "Total Devido Pelo Reclamado" ao final
    if "Total Devido Pelo Reclamado" in pdf_cols.tolist():
        pdf_dict["Total Devido Pelo Reclamado"] = pdf_vals[pdf_cols.tolist().index("Total Devido Pelo Reclamado")]
    return (pdf_file, pdf_dict)

# processa os arquivos PDF em paralelo
with ThreadPoolExecutor(max_workers=4) as executor:
    results = executor.map(process_pdf, pdf_files)

# cria um dicionário vazio para armazenar as informações dos PDFs
pdf_dict = {}

# junta os resultados em um dicionário único
for pdf_file, pdf_data in results:
    pdf_dict[pdf_file] = pdf_data

# converte o dicionário para um dataframe
merged_table = pd.DataFrame.from_dict(pdf_dict, orient='index')

# remove a coluna "Total Devido Pelo Reclamado" do dataframe e armazena em uma variável temporária
total_devido = merged_table.pop('Total Devido Pelo Reclamado')

# insere a coluna "Total Devido Pelo Reclamado" na última posição do dataframe
merged_table.insert(len(merged_table.columns), 'Total Devido Pelo Reclamado', total_devido)

# salva a tabela em um arquivo Excel
merged_table.to_excel('teste.xlsx', index_label='Nome do Arquivo')
