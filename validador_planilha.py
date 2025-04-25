import pandas as pd
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


def data_valida(data):
    try:
        datetime.strptime(str(data), "%d/%m/%Y")
        return True
    except:
        return False

def cpf_valido(cpf):
    cpf_str = re.sub(r'\D', '', str(cpf))
    return len(cpf_str) == 11

def editar_lista_de_erros(erros):
    with open("log_erros_planilha.txt", "w") as f:
        f.write(erros)

def processar_planilha():
    erros = []
    lista_erros_texto = "Erros encontrados na planilha:\n"

    try:
        arquivo = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Planilhas Excel", "*.xlsx")])
        if not arquivo:
            return

        df = pd.read_excel(arquivo, engine="openpyxl", skiprows=16)

        for index, row in df.iterrows():
            linha = index + 17

            if pd.isna(row['TIPO']) or row['TIPO'] != 1:
                erros.append(f"Linha {linha}: TIPO deve ser 1 e não pode estar em branco.")

            id_doc = str(row['CNPJ/CPF/CEI/CAEPF']).replace('.', '').replace('-', '').replace('/', '')
            if pd.isna(id_doc) or not id_doc.isdigit() or len(id_doc) > 14:
                erros.append(f"Linha {linha}: CNPJ/CPF/CEI/CAEPF inválido.")

            if pd.isna(row['Nome']) or len(str(row['Nome'])) > 50:
                erros.append(f"Linha {linha}: NOME vazio ou ultrapassando 50 caracteres.")

            if row['Nacionalidade'] not in ['BRA', 'ES']:
                erros.append(f"Linha {linha}: Nacionalidade deve ser BRA ou ES.")

            if not data_valida(row['Nascimento']):
                erros.append(f"Linha {linha}: Data de Nascimento inválida ou vazia.")

            if row['Sexo'] not in ['M', 'F']:
                erros.append(f"Linha {linha}: Sexo inválido (deve ser M ou F).")

            if pd.isna(row['CPF']) or not cpf_valido(row['CPF']) or len(str(row['CPF'])) > 14:
                erros.append(f"Linha {linha}: CPF inválido.")

            if pd.isna(row['Matrícula']) or len(str(row['Matrícula'])) > 30:
                erros.append(f"Linha {linha}: Matrícula inválida ou com mais de 30 caracteres.")

            if pd.isna(row['Matrícula RH']) or len(str(row['Matrícula RH'])) > 30:
                erros.append(f"Linha {linha}: Matrícula RH inválida ou com mais de 30 caracteres.")

            if not data_valida(row['Admissão']):
                erros.append(f"Linha {linha}: Data de Admissão inválida.")

            if not data_valida(row['Inicio']):
                erros.append(f"Linha {linha}: Data de Início inválida.")

            if pd.isna(row['Setor']) or len(str(row['Setor'])) > 100:
                erros.append(f"Linha {linha}: Setor inválido ou com mais de 100 caracteres.")

            if pd.isna(row['Cargo']) or len(str(row['Cargo'])) > 100:
                erros.append(f"Linha {linha}: Cargo inválido ou com mais de 100 caracteres.")

            if pd.isna(row['CBO']):
                erros.append(f"Linha {linha}: CBO não pode estar em branco.")

            if pd.isna(row['Descrição Sumária do Cargo']):
                erros.append(f"Linha {linha}: Descrição Sumária do Cargo não pode estar em branco.")

        # Mostrar resultados
        txt_resultado.delete(1.0, tk.END)

        if erros:
            for erro in erros:
                txt_resultado.insert(tk.END, erro + "\n")
                lista_erros_texto += erro + "\n"

            editar_lista_de_erros(lista_erros_texto)
            messagebox.showwarning("Validação concluída", "⚠️ Erros encontrados! Verifique o log.")
        else:
            txt_resultado.insert(tk.END, "✅ Nenhum erro encontrado. Planilha válida!\n")
            messagebox.showinfo("Validação concluída", "✅ Planilha válida!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Janela principal
janela = tk.Tk()
janela.title("Comparador de Planilhas")
janela.geometry("700x500")

lbl_titulo = tk.Label(janela, text="ミ★ Comparador de Planilhas ★彡", font=("Arial", 16))
lbl_titulo.pack(pady=10)

btn_selecionar = tk.Button(janela, text="Selecionar e Validar Planilha", command=processar_planilha, bg="#4CAF50", fg="white", font=("Arial", 12))
btn_selecionar.pack(pady=10)

txt_resultado = scrolledtext.ScrolledText(janela, width=80, height=20)
txt_resultado.pack(pady=10)

janela.mainloop()
