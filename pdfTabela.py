import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()

    # Abre a caixa para selecionar o arquivo
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o PDF de Estoque",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return caminho_pdf

def extrair_dados_pdf():
    arquivo_entrada = selecionar_arquivo()

    if not arquivo_entrada:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    arquivo_saida = arquivo_entrada.replace(".pdf", "_CONVERTIDO.xlsx")
    start_time = datetime.now()
    dados_finais = []

    try:
        with pdfplumber.open(arquivo_entrada) as pdf:
            total_paginas = len(pdf.pages)
            print(f"Processando {total_paginas} páginas...")

            for i, pagina in enumerate(pdf.pages):
                # Extração
                tabelas = pagina.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "join_tolerance": 3
                })

                for tabela in tabelas:
                    df_temp = pd.DataFrame(tabela)
                    # Remove linhas e colunas vazias
                    df_temp = df_temp.dropna(how='all').dropna(axis=1, how='all')
                    if not df_temp.empty:
                        # Remove linhas repitidas do cabeçalho
                        dados_finais.append(df_temp)
                # Progress no terminal
                if (i + 1) % 50 == 0:
                    print(f"Progresso: {i + 1}/{total_paginas} páginas concluídas.")
        # Consolida tudo

        if dados_finais:
            df_consolidado = pd.concat(dados_finais, ignore_index=True)

            # Exportação rápida xlsxwriter
            with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:

                df_consolidado.to_excel(writer, index=False, sheet_name='Estoque')

            tempo_total = datetime.now() - start_time
            print(f"\nSucesso! Arquivo salvo em: {arquivo_saida}")
            print(f"Tempo total de execução: {tempo_total}")
            messagebox.showinfo("Concluído", f"Processamento finalizado em {tempo_total}")
        else:
            print("Nenhuma tabela encontrada no PDF.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")


if __name__ == "__main__":
    extrair_dados_pdf()

