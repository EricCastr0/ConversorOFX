import pandas as pd
from datetime import datetime

def generate_ofx(df):
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