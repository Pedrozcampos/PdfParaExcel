import pandas as pd
import pdfplumber
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import re

class ExtratorBancarioMaster:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Extrator BB High Precision v7.0")
        self.root.geometry("600x400")
        ctk.set_appearance_mode("dark")
        
        self.label = ctk.CTkLabel(self.root, text="Conversor de Extratos Funerária", font=("Arial", 22, "bold"))
        self.label.pack(pady=20)
        
        self.btn_select = ctk.CTkButton(self.root, text="Selecionar PDF e Gerar Excel", 
                                        height=60, width=320, font=("Arial", 16, "bold"),
                                        fg_color="#1f538d", hover_color="#14375e",
                                        command=self.processar_pdf)
        self.btn_select.pack(pady=20)
        
        self.status_label = ctk.CTkLabel(self.root, text="Status: Pronto para processar", text_color="gray")
        self.status_label.pack(pady=10)

    def converter_valor(self, texto):
        """Converte '1.234,56 C' para float 1234.56"""
        if not texto: return None
        # Pega apenas os números, pontos e vírgulas antes do C ou D
        limpo = re.sub(r'[^\d,\.]', '', texto)
        if not limpo: return None
        return float(limpo.replace('.', '').replace(',', '.'))

    def processar_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path: return

        try:
            self.status_label.configure(text="Processando linhas... Alinhando colunas.", text_color="yellow")
            self.root.update()

            dados_finais = []
            
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    # Extraímos o texto mas mantendo a noção de onde as palavras estão (layout=True)
                    texto_pagina = page.extract_text(layout=True)
                    if not texto_pagina: continue

                    linhas = texto_pagina.split('\n')
                    for linha in linhas:
                        # 1. Busca Data (Início da linha)
                        match_data = re.search(r'(\d{2}/\d{2}/\d{4})', linha)
                        if not match_data: continue
                        data = match_data.group(1)

                        # 2. Busca todos os valores terminados em C ou D (Transação e Saldo)
                        # O regex identifica valores como 1.500,00 C ou 45,00 D
                        valores_encontrados = re.findall(r'(\d[\d\.,]*\s*[CD])', linha)
                        
                        if not valores_encontrados: continue

                        # Lógica das colunas:
                        # O primeiro valor com C/D é sempre a movimentação (Débito ou Crédito)
                        # O último valor com C/D é sempre o Saldo daquela linha
                        val_mov_texto = valores_encontrados[0]
                        val_saldo_texto = valores_encontrados[-1]

                        valor_mov = self.converter_valor(val_mov_texto)
                        valor_saldo = self.converter_valor(val_saldo_texto)

                        # Identifica se é Débito ou Crédito
                        debito = valor_mov if 'D' in val_mov_texto.upper() else ""
                        credito = valor_mov if 'C' in val_mov_texto.upper() else ""
                        
                        # Ajusta o sinal do saldo para o Excel (opcional)
                        saldo_final = -valor_saldo if 'D' in val_saldo_texto.upper() else valor_saldo

                        # 3. Extrai o Histórico (o que sobrar no meio)
                        # Removemos a data e os valores financeiros da linha para isolar o texto
                        historico = linha.replace(data, "").replace(val_mov_texto, "")
                        if len(valores_encontrados) > 1:
                            historico = historico.replace(val_saldo_texto, "")
                        
                        # Limpeza de caracteres residuais e espaços
                        historico = re.sub(r'\s+', ' ', historico).strip()
                        # Remove a agência (ex: 8684-3) se ela aparecer no histórico
                        historico = re.sub(r'\d{4}-\d', '', historico).strip()

                        dados_finais.append([data, historico, debito, credito, saldo_final])

            # Criação do DataFrame e limpeza
            df = pd.DataFrame(dados_finais, columns=['Data', 'Histórico', 'Débito', 'Crédito', 'Saldo'])
            
            # Remove o Saldo Anterior para não sujar o relatório e tira duplicatas
            df = df[~df['Histórico'].str.contains("Saldo Anterior", case=False, na=False)]
            df = df.drop_duplicates()

            if df.empty:
                raise Exception("Não foi possível extrair dados. O PDF pode ser uma imagem.")

            # Salva o arquivo
            output_path = os.path.splitext(file_path)[0] + "_CONVERTIDO_BB.xlsx"
            df.to_excel(output_path, index=False)

            self.status_label.configure(text=f"Sucesso! {len(df)} linhas extraídas.", text_color="green")
            messagebox.showinfo("Concluído", f"Arquivo Excel gerado com sucesso!\nSalvo como: {os.path.basename(output_path)}")

        except Exception as e:
            self.status_label.configure(text="Erro no processamento", text_color="red")
            messagebox.showerror("Erro", f"Falha ao converter: {str(e)}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExtratorBancarioMaster()
    app.run()