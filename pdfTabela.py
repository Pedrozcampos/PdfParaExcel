import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os

def extrair_seguro():
    root = tk.Tk()
    root.withdraw()
    arquivos = filedialog.askopenfilenames(title="Selecione os 12 PDFs", filetypes=[("PDF", "*.pdf")])
    if not arquivos: return

    pasta_destino = os.path.dirname(arquivos[0])
    # Mudamos para .csv para garantir que o arquivo não corrompa
    arquivo_saida = os.path.join(pasta_destino, "COMPILADO_ESTOQUE_FINAL.csv")
    
    start_time = datetime.now()
    dados_acumulados = []

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando extração segura...")

    try:
        for caminho in arquivos:
            nome_arq = os.path.basename(caminho)
            doc = fitz.open(caminho)
            print(f">>> Processando: {nome_arq}")

            for i, pagina in enumerate(doc):
                # Extração por palavras para manter a velocidade
                palavras = pagina.get_text("words")
                
                linhas = {}
                for p in palavras:
                    y = round(p[1], 1)
                    texto = p[4]
                    if y not in linhas:
                        linhas[y] = []
                    linhas[y].append(texto)
                
                for y in sorted(linhas.keys()):
                    # Criamos uma linha limpa: [Produto, Data, Valor..., Nome do Arquivo]
                    linha_dados = [" ".join(linhas[y]), nome_arq]
                    dados_acumulados.append(linha_dados)

                if (i + 1) % 500 == 0:
                    print(f"   Progresso: {i + 1}/{len(doc)} páginas...")

        if dados_acumulados:
            print("\nSalvando arquivo CSV (Formato compatível com Excel)...")
            df_final = pd.DataFrame(dados_acumulados)
            
            # Salvando como CSV com separador ';' que o Excel do Brasil adora
            df_final.to_csv(arquivo_saida, index=False, sep=';', encoding='latin1', errors='replace')
            
            tempo_total = datetime.now() - start_time
            print(f"--- FINALIZADO EM {tempo_total} ---")
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso!\nLocal: {arquivo_saida}")
        else:
            print("Nenhum dado encontrado.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    extrair_seguro()