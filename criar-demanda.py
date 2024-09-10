import openpyxl
import os
import re

def get_column_index(sheet, column_name):
    for idx, cell in enumerate(sheet[1]):
        if cell.value and cell.value.strip() == column_name:
            return idx
    return None

workbook = openpyxl.load_workbook('C:/Users/amelo/Desktop/relatorio.xlsx')
sheet = workbook.active

print("Nomes das colunas na planilha:")
for idx, cell in enumerate(sheet[1]):
    print(f"Índice: {idx}, Nome da coluna: {cell.value}")

indice_associado = get_column_index(sheet, 'Nome')
indice_placa = get_column_index(sheet, 'Placa')
indice_chassi = get_column_index(sheet, 'Chassi')
indice_modelo = get_column_index(sheet, 'Modelo')
indice_ano = get_column_index(sheet, 'Ano Mod.')
indice_cidade = get_column_index(sheet, 'Cidade Veículo')
indice_bairro = get_column_index(sheet, 'Bairro')
indice_tipo_veiculo = get_column_index(sheet, 'Tipo Veículo')
indice_telefone_1 = get_column_index(sheet, 'Telefone')
indice_telefone_2 = get_column_index(sheet, 'Telefone Celular')
indice_fipe = get_column_index(sheet, 'Valor FIPE Veiculo')
indice_tipo_adesao = get_column_index(sheet, 'Tipo Adesão')

indices = [indice_associado, indice_placa, indice_chassi, indice_modelo, indice_ano, indice_cidade, indice_bairro,
           indice_telefone_1, indice_telefone_2, indice_fipe, indice_tipo_adesao, indice_tipo_veiculo]
if any(idx is None for idx in indices):
    raise ValueError("Um ou mais índices de coluna não foram encontrados. Verifique os nomes das colunas.")

caminho_pasta = "C:/Users/amelo/Downloads/"
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)

caminho_arquivo_combinado = os.path.join(caminho_pasta, "ordens_servico.txt")

caminhos_individuais = []

def telefone_valido(telefone):
    if telefone is None:
        return False, "", ""
    telefone_limpo = re.sub(r'\D', '', telefone)
    if re.search(r'000000|999999', telefone_limpo):
        return False, "", ""
    if telefone_limpo in ["", "0000", "99999999", "999999999", "000000000"]:
        return False, "", ""
    return True, telefone_limpo, telefone

def converter_para_numero(valor):
    if valor is None:
        return None
    try:
        return int(valor)
    except ValueError:
        return None

for row in sheet.iter_rows(min_row=2, values_only=True):
    if row[indice_associado] is None or row[indice_associado] == "":
        print("Linha vazia encontrada na primeira coluna. Parando a execução.")
        break

    tipo_adesao = row[indice_tipo_adesao]
    if tipo_adesao and tipo_adesao != "SEM / RASTREADOR":
        continue
    
    associado = row[indice_associado]  
    placa = row[indice_placa] if row[indice_placa] else row[indice_chassi]  
    modelo = row[indice_modelo]  
    ano = row[indice_ano]  
    cidade = row[indice_cidade]  
    bairro = row[indice_bairro]  
    
    telefone_1_valido, telefone_1_limpo, telefone_1_original = telefone_valido(row[indice_telefone_1])
    telefone_2_valido, telefone_2_limpo, telefone_2_original = telefone_valido(row[indice_telefone_2])
    
    telefones = []
    if telefone_1_valido and telefone_1_limpo not in telefones:
        telefones.append(telefone_1_original)
    if telefone_2_valido and telefone_2_limpo != telefone_1_limpo:
        telefones.append(telefone_2_original)
    telefone = " | ".join(telefones)

    valor_fipe = converter_para_numero(row[indice_fipe])
    
    if row[indice_tipo_veiculo] == "MONITORAMENTO (CARTRACKING)":
        obs = "MONITORAMENTO"
    elif valor_fipe and valor_fipe > 150000:
        obs = "2 RASTREADORES"
    else:
        obs = ""
    
    conteudo = (
        f"*INSTALAÇÃO*\n\n"
        f"*ASSOCIADO*: {associado}\n"
        f"*PLACA*: {placa}\n"
        f"*MODELO*: {modelo}\n"
        f"*ANO*: {ano}\n"
        f"*CIDADE*: {cidade}\n"
        f"*BAIRRO*: {bairro}\n"
        f"*TELEFONE*: {telefone}\n\n"
        f"*OBS*: {obs}\n\n"
    )

    nome_arquivo_individual = f"ordem_servico_{placa}.txt"
    caminho_arquivo_individual = os.path.join(caminho_pasta, nome_arquivo_individual)
    
    with open(caminho_arquivo_individual, 'w', encoding='utf-8') as arquivo_individual:
        arquivo_individual.write(conteudo)

    caminhos_individuais.append(caminho_arquivo_individual)

with open(caminho_arquivo_combinado, 'w', encoding='utf-8') as arquivo_combinado:
    for caminho_individual in caminhos_individuais:
        with open(caminho_individual, 'r', encoding='utf-8') as arquivo_individual:
            conteudo = arquivo_individual.read()
            arquivo_combinado.write(conteudo)
            arquivo_combinado.write("\n\n")  

print("Arquivos salvos e combinados com sucesso!")
