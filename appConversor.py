import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# Configuração da interface
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class OFXConverterApp:
    def __init__(self, root):
        root.iconbitmap('logo.ico')
        self.root = root
        self.root.title("Conversor Excel para OFX")
        self.root.geometry("600x450")
        self.root.resizable(False,False)

        # Botão para selecionar arquivo Excel
        self.select_button = ctk.CTkButton(root, text="Selecionar Arquivo Excel", command=self.select_file)
        self.select_button.pack(pady=20)

        # Botão para converter para OFX
        self.convert_button = ctk.CTkButton(root, text="Converter para OFX", command=self.convert_to_ofx, state="disabled")
        self.convert_button.pack(pady=20)

        # Status
        self.status_label = ctk.CTkLabel(root, text="Nenhum arquivo selecionado.")
        self.status_label.pack(pady=20)

        # Rodapé
        self.footer_label = ctk.CTkLabel(
            root,
            text="Desenvolvido por: Eric Castro",
            text_color="#8f8fb8",  # Cor azul
            font=("Arial", 9),  # Fonte e tamanho
            fg_color=None,  # Fundo transparente
        )
        self.footer_label.pack(side="bottom", pady=10)  # Posiciona no rodapé

        self.file_path = None

    def select_file(self):
        # Abre a janela para selecionar o arquivo Excel
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.status_label.configure(text=f"Arquivo selecionado: {self.file_path}")
            self.convert_button.configure(state="normal")
        else:
            self.status_label.configure(text="Nenhum arquivo selecionado.")
            self.convert_button.configure(state="disabled")

    def convert_to_ofx(self):
        if not self.file_path:
            self.status_label.configure(text="Nenhum arquivo selecionado.")
            return

        try:
            # Lê o arquivo Excel
            df = pd.read_excel(self.file_path, sheet_name="pagamentos")

            # Verifica se o arquivo tem o formato esperado
            if df.shape[1] < 29:  # Verifica se há pelo menos 29 colunas
                raise ValueError("O arquivo Excel não tem o formato esperado.")

            # Mapeia as colunas manualmente
            df.columns = [
                "data do recebimento", "data original da venda", "data original de vencimento",
                "valor bruto da parcela original", "valor bruto da parcela atualizada", "taxa MDR",
                "valor MDR descontado", "valor líquido da parcela", "negociada", "%", "NSU/CV",
                "TID", "número do pedido", "número da autorização", "resumo de vendas/número do lote",
                "nome do estabelecimento", "estabelecimento", "número do cartão", "indicador de transação tokenizada",
                "código IATA", "modalidade", "bandeira", "número de parcelas", "parcela", "banco",
                "agência", "conta-corrente", "cancelamento/contestação", "data do cancelamento", "status"
            ]

            # Converte as colunas de data para datetime
            df["data do recebimento"] = pd.to_datetime(df["data do recebimento"], errors="coerce")
            df["data original da venda"] = pd.to_datetime(df["data original da venda"], errors="coerce")

            # Usa 'data do recebimento' como fallback para 'data do recebimento' inválida
            df["data_utilizada"] = df["data original da venda"].combine_first(df["data do recebimento"])

            # Verifica se há valores inválidos (NaT) na coluna 'data_utilizada'
            #if df["data_utilizada"].isnull().any():
            #    messagebox.showwarning("Aviso", "Algumas datas na coluna 'data do recebimento' e 'data original da venda' são inválidas e serão ignoradas.")

            # Remove linhas com datas inválidas
            df = df.dropna(subset=["data_utilizada"])

            # Gera o conteúdo do arquivo OFX
            ofx_content = self.generate_ofx(df)

            # Salva o arquivo OFX
            output_path = filedialog.asksaveasfilename(defaultextension=".ofx", filetypes=[("OFX files", "*.ofx")])
            if output_path:
                with open(output_path, "w", encoding="utf-8") as file:
                    file.write(ofx_content)
                self.status_label.configure(text=f"Arquivo OFX salvo em: {output_path}")
                messagebox.showinfo("Sucesso", "Arquivo OFX gerado com sucesso!")
            else:
                self.status_label.configure(text="Conversão cancelada.")
        except Exception as e:
            self.status_label.configure(text=f"Erro: {str(e)}")
            messagebox.showerror("Erro", str(e))

    def generate_ofx(self, df):
        # Cabeçalho do OFX
        current_time = datetime.now().strftime('%Y%m%d%H%M%S')
        start_date = df["data_utilizada"].min().strftime('%Y%m%d')
        end_date = df["data_utilizada"].max().strftime('%Y%m%d')

        ofx_header = f"""OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
    <SIGNONMSGSRSV1>
        <SONRS>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <DTSERVER>{current_time}</DTSERVER>
            <LANGUAGE>POR</LANGUAGE>
        </SONRS>
    </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>1</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
                <CURDEF>BRL</CURDEF>
                <BANKACCTFROM>
                    <BANKID>341</BANKID>  <!-- Código do Itaú -->
                    <ACCTID>{df.iloc[0]['conta-corrente']}</ACCTID>
                    <ACCTTYPE>CHECKING</ACCTTYPE>
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART>{start_date}</DTSTART>
                    <DTEND>{end_date}</DTEND>
"""

        # Transações (valor bruto da parcela original)
        transactions = ""
        for _, row in df.iterrows():
            transaction_date = row["data_utilizada"].strftime('%Y%m%d')
            amount = row["valor bruto da parcela original"]
            memo = f"{row['nome do estabelecimento']} - {row['bandeira']} {row['modalidade']}"

            # Define o valor como positivo (crédito)
            trntype = "CREDIT"
            trnamt = f"{amount:.2f}"  # Valores sempre positivos

            transactions += f"""
                    <STMTTRN>
                        <TRNTYPE>{trntype}</TRNTYPE>
                        <DTPOSTED>{transaction_date}</DTPOSTED>
                        <TRNAMT>{trnamt}</TRNAMT>
                        <MEMO>{memo}</MEMO>
                    </STMTTRN>
"""

            # Transações (valor MDR descontado como débito)
            mdr_amount = row["valor MDR descontado"]
            if pd.notna(mdr_amount) and mdr_amount != 0:  # Ignora valores nulos ou zero
                mdr_memo = f"MDR Descontado - {row['nome do estabelecimento']}"
                transactions += f"""
                    <STMTTRN>
                        <TRNTYPE>DEBIT</TRNTYPE>
                        <DTPOSTED>{transaction_date}</DTPOSTED>
                        <TRNAMT>-{mdr_amount:.2f}</TRNAMT>
                        <MEMO>{mdr_memo}</MEMO>
                    </STMTTRN>
"""

        # Rodapé do OFX
        ofx_footer = f"""
                </BANKTRANLIST>
                <LEDGERBAL>
                    <BALAMT>1000.00</BALAMT>  <!-- Saldo fictício -->
                    <DTASOF>{current_time}</DTASOF>
                </LEDGERBAL>
            </STMTRS>
        </STMTTRNRS>
    </BANKMSGSRSV1>
</OFX>
"""

        return ofx_header + transactions + ofx_footer

# Executa a aplicação
if __name__ == "__main__":
    root = ctk.CTk()
    app = OFXConverterApp(root)
    root.mainloop()