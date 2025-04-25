import pandas as pd
import re
from datetime import datetime

print("""
 ミ★ Bem vindo ao comparador de planilhas ★彡
""")

"""======================================================= Aqui ficam as principais variaveis do programa """
#arquivo = "c:/Users/eulr1059/Downloads/teste.xlsx" # Caminho para a planinlha que deve ser validada.
arquivo = input("Digite o caminho da planilha: ")
df = pd.read_excel(arquivo, engine="openpyxl", skiprows=16) # Carregando a planilha na variavel df, ignorando as 16 primeiras linhas
erros = [] # Cria um array vazio para dicionar os erros encontrados.
lista_erros_texto = "Erros encontrados na planilha: " + "\n" # Variavel que vai ser adicionada em forma de texto todos os erros encontrados


"""======================================================= Aqui Vão ficar as principais funções do sistema """
# Função para validar datas no formato dd/mm/aaaa
def data_valida(data):
    try:
        datetime.strptime(str(data), "%d/%m/%Y")
        return True
    except:
        return False

# Função para validar CPF (simples, sem validação de dígito verificador)
def cpf_valido(cpf):
    cpf_str = re.sub(r'\D', '', str(cpf))
    return len(cpf_str) == 11

# Função para criar um arquivo de texto 
def editar_lista_de_erros(erros):
    logDeErros = open("log_erros_planilha.txt", "w")
    logDeErros.write(erros)
    logDeErros.close()



"""======================================================= Aqui é feito um for iterando todos os itens da planilha e fazendo validações para encontrar erros """
# Loop linha por linha
for index, row in df.iterrows():
    linha = index + 17  # Ajuste para refletir a linha real na planilha     

    # TIPO
    if pd.isna(row['TIPO']) or row['TIPO'] != 1:
        erros.append(f"Linha {linha}: TIPO deve ser 1 e não pode estar em branco.")

    # CNPJ/CPF/CEI/CAEPF
    id_doc = str(row['CNPJ/CPF/CEI/CAEPF']).replace('.', '').replace('-', '').replace('/', '')
    if pd.isna(id_doc) or not id_doc.isdigit() or len(id_doc) > 14:
        erros.append(f"Linha {linha}: CNPJ/CPF/CEI/CAEPF inválido.")

    # NOME
    if pd.isna(row['Nome']) or len(str(row['Nome'])) > 50:
        erros.append(f"Linha {linha}: NOME vazio ou ultrapassando 50 caracteres.")

    # NACIONALIDADE
    if row['Nacionalidade'] not in ['BRA', 'ES']:
        erros.append(f"Linha {linha}: Nacionalidade deve ser BRA ou ES.")

    # NASCIMENTO
    if not data_valida(row['Nascimento']):
        erros.append(f"Linha {linha}: Data de Nascimento inválida ou vazia.")

    # SEXO
    if row['Sexo'] not in ['M', 'F']:
        erros.append(f"Linha {linha}: Sexo inválido (deve ser M ou F).")

    # CPF
    if pd.isna(row['CPF']) or not cpf_valido(row['CPF']) or len(str(row['CPF'])) > 14:
        erros.append(f"Linha {linha}: CPF inválido.")

    # Matrícula
    if pd.isna(row['Matrícula']) or len(str(row['Matrícula'])) > 30:
        erros.append(f"Linha {linha}: Matrícula inválida ou com mais de 30 caracteres.")

    # Matrícula RH
    if pd.isna(row['Matrícula RH']) or len(str(row['Matrícula RH'])) > 30:
        erros.append(f"Linha {linha}: Matrícula RH inválida ou com mais de 30 caracteres.")

    # Admissão
    if not data_valida(row['Admissão']):
        erros.append(f"Linha {linha}: Data de Admissão inválida.")

    # Início
    if not data_valida(row['Inicio']):
        erros.append(f"Linha {linha}: Data de Início inválida.")

    # Setor
    if pd.isna(row['Setor']) or len(str(row['Setor'])) > 100:
        erros.append(f"Linha {linha}: Setor inválido ou com mais de 100 caracteres.")

    # Cargo
    if pd.isna(row['Cargo']) or len(str(row['Cargo'])) > 100:
        erros.append(f"Linha {linha}: Cargo inválido ou com mais de 100 caracteres.")

    # CBO
    if pd.isna(row['CBO']):
        erros.append(f"Linha {linha}: CBO não pode estar em branco.")

    # Descrição Sumária
    if pd.isna(row['Descrição Sumária do Cargo']):
        erros.append(f"Linha {linha}: Descrição Sumária do Cargo não pode estar em branco.")


"""======================================================= Aqui é verificado o array de erros, se tiver algum erro é criado o arquivo txt com o log de erros """
# Mostrar os erros encontrados
if erros:
    print("⚠️ Erros encontrados na planilha:")    

    for erro in erros:
        print(erro) # Imprime no console cada erro encontrado
        lista_erros_texto += erro + "\n" # Adiciona uma linha de texto para cada erro encontrado

    editar_lista_de_erros(lista_erros_texto) # Chama a função para criar o arquivo txt com os erros

else:
    print("✅ Nenhum erro encontrado. Planilha válida!") # Se não houver erros, imprime que a planilha está válida


