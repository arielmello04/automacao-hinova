import openpyxl
import os

workbook = openpyxl.load_workbook('C:/Users/amelo/Desktop/relatorio.xlsx')
sheet = workbook.active

def telefone_valido(telefone):
    if telefone in ["", "()", "0000", "99999999"]:
        return False
    return True

def converter_para_numero(valor):
    try:
        return int(valor)
    except ValueError:
        return None

caminho_pasta = "C:/Users/amelo/Downloads/"
if not os.path.exists(caminho_pasta):
    os.makedirs(caminho_pasta)

caminho_arquivo = os.path.join(caminho_pasta, "ordens_servico.txt")    

with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:

    for row in sheet.iter_rows(min_row=2, values_only=True):
        tipo_adesao = row[16]  
        
        if tipo_adesao and tipo_adesao != "SEM / RASTREADOR":
            continue
        
        associado = row[0]  
        placa = row[1] if row[1] else row[15]  
        modelo = row[13]  
        ano = row[14]  
        cidade = row[8]  
        bairro = row[9] 
        
        telefone_1 = row[5] if telefone_valido(row[5]) else ""
        telefone_2 = row[6] if telefone_valido(row[6]) else ""
        telefone = f"{telefone_1} {telefone_2}".strip()
        
        valor_fipe = converter_para_numero(row[11])
        
        if row[10] == "MONITORAMENTO (CARTRACKING)":
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
            f"*OBS*: {obs}\n\n\n"
        )
        arquivo.write(conteudo)
        arquivo.write("\n\n\n")