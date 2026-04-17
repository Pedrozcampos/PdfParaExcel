import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os

def selecionar_arquivos():
    root = tk.Tk()
    root.withdraw()
    
    # .askopenfilenames (com 's') permite selecionar vários arquivos de uma vez
    caminhos_pdfs = filedialog.askopenfilenames(
        title="Selecione os 12 PDFs de Estoque (Segure Ctrl para selecionar vários)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return caminhos_pdfs

def extrair_dados_pdf_compilado():
    # Seleciona a lista de arquivos
    lista_arquivos = selecionar_arquivos()
    
    if not lista_arquivos:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    # O arquivo final será salvo na pasta do primeiro PDF selecionado
    pasta_destino = os.path.dirname(lista_arquivos[0])
    arquivo_saida = os.path.join(pasta_destino, "ESTOQUE_ANUAL_COMPILADO.xlsx")
    
    start_time = datetime.now()
    dados_finais_anuais = []
    
    # Configurações de extração que você validou
    conf_tabela = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3
    }

    try:
        print(f"Iniciando processamento de {len(lista_arquivos)} arquivos...")

        for idx, caminho_pdf in enumerate(lista_arquivos):
            nome_arquivo = os.path.basename(caminho_pdf)
            print(f"\n[{idx+1}/{len(lista_arquivos)}] Processando: {nome_arquivo}")
            
            with pdfplumber.open(caminho_pdf) as pdf:
                total_paginas = len(pdf.pages)
                
                for i, pagina in enumerate(pdf.pages):
                    # Extração mantendo sua lógica original
                    tabelas = pagina.extract_tables(conf_tabela)
                    
                    for tabela in tabelas:
                        if tabela:
                            df_temp = pd.DataFrame(tabela)
                            # Limpeza de nulos
                            df_temp = df_temp.dropna(how='all').dropna(axis=1, how='all')
                            
                            if not df_temp.empty:
                                # Opcional: Adiciona uma coluna com o nome do arquivo para saber de qual mês é o dado
                                df_temp['Origem'] = nome_arquivo
                                dados_finais_anuais.append(df_temp)
                    
                    # Progresso a cada 50 páginas de cada arquivo
                    if (i + 1) % 50 == 0 or (i + 1) == total_paginas:
                        print(f"   -> Página {i + 1}/{total_paginas} concluída.")

        # Consolidação de todos os arquivos em um único DataFrame
        if dados_finais_anuais:
            print("\nUnificando todos os meses em um único arquivo... Aguarde.")
            df_consolidado = pd.concat(dados_finais_anuais, ignore_index=True)
            
            # Exportação final usando xlsxwriter para suportar o grande volume de dados
            print(f"Salvando Excel em: {arquivo_saida}")
            with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
                df_consolidado.to_excel(writer, index=False, sheet_name='Estoque_Anual')
            
            tempo_total = datetime.now() - start_time
            print(f"\nSUCESSO!")
            print(f"Tempo total de execução: {tempo_total}")
            messagebox.showinfo("Concluído", f"Todos os arquivos foram compilados com sucesso!\n\nTempo: {tempo_total}")
        else:
            print("Nenhuma tabela encontrada nos arquivos selecionados.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro crítico: {e}")

if __name__ == "__main__":
    # Certifique-se de ter instalado: pip install pdfplumber pandas xlsxwriter
    extrair_dados_pdf_compilado()