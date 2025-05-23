import pandas as pd
import re
from datetime import datetime
import flet as ft


# Validação de data
def data_valida(data):
    try:
        datetime.strptime(str(data), "%d/%m/%Y")
        return True
    except:
        return False


# Validação de CPF
def cpf_valido(cpf):
    cpf_str = re.sub(r'\D', '', str(cpf))
    return len(cpf_str) == 11


# Cria arquivo com os erros encontrados
def editar_lista_de_erros(erros):
    with open("log_erros_planilha.txt", "w", encoding="utf-8") as f:
        f.write(erros)


def main(page: ft.Page):
    page.title = "Comparador de Planilhas"
    page.scroll = ft.ScrollMode.ALWAYS

    resultado_texto = ft.Text("", selectable=True, color=ft.colors.WHITE)

    # FilePicker
    def on_file_selected(e: ft.FilePickerResultEvent):
        if not e.files:
            resultado_texto.value = "❌ Nenhum arquivo selecionado."
            page.update()
            return

        arquivo = e.files[0].path
        try:
            df = pd.read_excel(arquivo, engine="openpyxl", skiprows=16)
            erros = []
            log_texto = "Erros encontrados na planilha:\n"

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

            if erros:
                log_final = log_texto + "\n".join(erros)
                resultado_texto.value = f"⚠️ {len(erros)} erro(s) encontrados.\n\n" + "\n".join(erros)
                editar_lista_de_erros(log_final)
            else:
                resultado_texto.value = "✅ Nenhum erro encontrado. Planilha válida!"

        except Exception as ex:
            resultado_texto.value = f"❌ Erro ao processar a planilha:\n{str(ex)}"

        page.update()

    file_picker = ft.FilePicker(on_result=on_file_selected)
    page.overlay.append(file_picker)

    titulo = ft.Text("ミ★ Validador de Planilhas ★彡", size=24, weight=ft.FontWeight.BOLD)

    botao = ft.ElevatedButton(
        text="Selecionar e Validar Planilha",
        on_click=lambda _: file_picker.pick_files(allow_multiple=False)
    )

    conteudo = ft.Column(
        controls=[
            ft.Row( [titulo], alignment=ft.MainAxisAlignment.CENTER ),
            ft.Row( [botao], alignment=ft.MainAxisAlignment.CENTER ),
            resultado_texto
        ],
        spacing=20
    )

    page.add(conteudo)


#ft.app(target=main)
ft.app(main, view=ft.AppView.WEB_BROWSER)