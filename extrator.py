import pandas as pd
import pdfplumber
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import re

class ExtratorBradescoGV:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Extrator Bradesco GV - Sistema de Precisão")
        self.root.geometry("600x400")
        ctk.set_appearance_mode("dark")
        
        # Interface
        self.label = ctk.CTkLabel(self.root, text="Conversor Bradesco (26 Páginas)", font=("Arial", 22, "bold"))
        self.label.pack(pady=20)
        
        self.btn_select = ctk.CTkButton(self.root, text="Selecionar PDF Bradesco", 
                                        height=60, width=320, font=("Arial", 16, "bold"),
                                        command=self.processar_pdf)
        self.btn_select.pack(pady=20)
        
        self.status_label = ctk.CTkLabel(self.root, text="Status: Aguardando Arquivo", text_color="gray")
        self.status_label.pack(pady=10)

    def converter_valor(self, texto):
        """Converte o formato brasileiro '1.234,56' para float 1234.56"""
        if not texto: return 0.0
        limpo = texto.replace('.', '').replace(',', '.')
        try:
            return float(limpo)
        except:
            return 0.0

    def processar_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path: return

        try:
            self.status_label.configure(text="Processando páginas... Por favor aguarde.", text_color="yellow")
            self.root.update()

            dados_finais = []
            data_atual = ""

            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    # O segredo: extrair texto com x_tolerance para não grudar colunas
                    texto_pag = page.extract_text(x_tolerance=3, y_tolerance=3)
                    if not texto_pag: continue

                    linhas = texto_pag.split('\n')
                    for linha in linhas:
                        # 1. Identificar a Data (DD/MM/AAAA)
                        match_data = re.search(r'(\d{2}/\d{2}/\d{4})', linha)
                        if match_data:
                            data_atual = match_data.group(1)

                        # 2. Identificar Valores (formato 0,00 ou -0,00 no Bradesco)
                        # Buscamos todos os números que terminam em ,XX
                        valores = re.findall(r'(-?[\d\.]+,\d{2})', linha)

                        if valores and data_atual:
                            # No layout do Bradesco:
                            # O ÚLTIMO valor da linha é sempre o SALDO.
                            saldo_str = valores[-1]
                            
                            # O PENÚLTIMO valor (se existir) é o valor da MOVIMENTAÇÃO.
                            mov_str = valores[-2] if len(valores) > 1 else ""

                            # 3. Extrair Histórico
                            # Removemos a data e os valores financeiros da string original
                            historico = linha.replace(data_atual, "").replace(saldo_str, "")
                            if mov_str:
                                historico = historico.replace(mov_str, "")
                            
                            historico = historico.strip()

                            # Ignorar linhas de cabeçalho
                            if "SALDO ANTERIOR" in historico.upper() or "LANÇAMENTO" in historico.upper():
                                if "SALDO ANTERIOR" in historico.upper():
                                    dados_finais.append([data_atual, "SALDO ANTERIOR", "", "", self.converter_valor(saldo_str)])
                                continue

                            # 4. Tratar Crédito/Débito
                            valor_num = self.converter_valor(mov_str)
                            credito = valor_num if valor_num > 0 else ""
                            debito = valor_num if valor_num < 0 else ""
                            saldo_num = self.converter_valor(saldo_str)

                            dados_finais.append([data_atual, historico, credito, debito, saldo_num])

            # Criar DataFrame e Excel
            if not dados_finais:
                raise Exception("Não foram encontrados dados no formato esperado.")

            df = pd.DataFrame(dados_finais, columns=['Data', 'Histórico', 'Crédito', 'Débito', 'Saldo'])
            
            # Limpeza final: remover linhas totalmente duplicadas
            df = df.drop_duplicates()

            output_path = os.path.splitext(file_path)[0] + "_CONVERTIDO.xlsx"
            df.to_excel(output_path, index=False)

            self.status_label.configure(text="Sucesso!", text_color="green")
            messagebox.showinfo("Concluído", f"Arquivo Excel gerado com sucesso!\n{len(df)} linhas processadas.")

        except Exception as e:
            self.status_label.configure(text="Erro no processamento", text_color="red")
            messagebox.showerror("Erro", f"Detalhes: {str(e)}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExtratorBradescoGV()
    app.run()