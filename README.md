# ConversorOFX

🔄 Aplicativo para converter relatórios Excel de pagamentos de cartão para formato OFX.

## 🚀 Instalação

```bash
pip install pandas customtkinter openpyxl
```

## 📖 Como Usar

1. Execute: `python appConversor.py`
2. Selecione arquivo Excel com planilha "pagamentos"
3. Clique em "Converter para OFX"
4. Escolha onde salvar o arquivo

## 📋 Estrutura do Excel

O arquivo deve ter uma planilha chamada "pagamentos" com 29 colunas:
- Data do recebimento, Data original da venda, Data original de vencimento
- Valor bruto da parcela original, Valor bruto da parcela atualizada, Taxa MDR
- Valor MDR descontado, Valor líquido da parcela, Negociada, %
- NSU/CV, TID, Número do pedido, Número da autorização, Resumo de vendas/número do lote
- Nome do estabelecimento, Estabelecimento, Número do cartão, Indicador de transação tokenizada
- Código IATA, Modalidade, Bandeira, Número de parcelas, Parcela
- Banco, Agência, Conta-corrente, Cancelamento/contestação, Data do cancelamento, Status

## ⚙️ Configuração

- Banco: Itaú (código 341)
- Moeda: BRL
- Transações: Créditos (pagamentos) e Débitos (MDR)
