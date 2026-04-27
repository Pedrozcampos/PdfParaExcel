import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def extrair_dados_pdf():
    root = tk.Tk()
    root.withdraw()
    caminho_pdf = filedialog.askopenfilename(title="Selecione o PDF", filetypes=[("PDF", "*.pdf")])
    if not caminho_pdf: return

    start_time = datetime.now()
    arquivo_saida = caminho_pdf.replace(".pdf", "_TABELA_FINAL.xlsx")
    
    # Mapeamento aproximado das colunas baseado no layout do Modelo-03
    # Essas coordenadas (x0) cobrem os campos: Esp, Número, Dia, Mês, Ano, Cod, Quant, Vl Unit, etc.
    limites_colunas = [0, 35, 90, 115, 135, 155, 200, 250, 310, 360, 420, 480, 540, 600, 660, 720, 800]

    dados_finais = []

    try:
        doc = fitz.open(caminho_pdf)
        print(f"Processando {len(doc)} páginas...")

        for pagina in doc:
            palavras = pagina.get_text("words") # Retorna (x0, y0, x1, y1, "texto", ...)
            
            # Agrupar palavras por linha (y0)
            linhas_dict = {}
            for p in palavras:
                y_coord = round(p[1], 1) # y0 é a altura
                if y_coord not in linhas_dict:
                    linhas_dict[y_coord] = []
                linhas_dict[y_coord].append(p)

            for y in sorted(linhas_dict.keys()):
                linha_palavras = linhas_dict[y]
                # Criar uma linha vazia com o número de colunas definido
                linha_formatada = [""] * len(limites_colunas)
                
                eh_linha_util = False
                for p in linha_palavras:
                    x0, texto = p[0], p[4]
                    
                    # Filtro para ignorar cabeçalhos
                    if any(c in texto for c in ["Página", "Série", "Registro", "Controle"]):
                        continue

                    # Identifica em qual coluna o texto se encaixa baseado no X0
                    for i in range(len(limites_colunas)):
                        limite_atual = limites_colunas[i]
                        proximo_limite = limites_colunas[i+1] if i+1 < len(limites_colunas) else 1000
                        
                        if limite_atual <= x0 < proximo_limite:
                            # Se já houver texto na célula (ex: números grandes), concatena
                            linha_formatada[i] = f"{linha_formatada[i]} {texto}".strip()
                            if texto.replace('.','').replace(',','').isdigit() or texto == "NF":
                                eh_linha_util = True
                            break
                
                if eh_linha_util:
                    dados_finais.append(linha_formatada)

        if dados_finais:
            df = pd.DataFrame(dados_finais)
            # Remove colunas que ficaram totalmente vazias
            df = df.dropna(how='all', axis=1)
            
            with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)

            print(f"Sucesso! Tempo: {datetime.now() - start_time}")
            messagebox.showinfo("Sucesso", f"Processado com colunas alinhadas!")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado encontrado.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))
    finally:
        if 'doc' in locals(): doc.close()

if __name__ == "__main__":
    extrair_dados_pdf()