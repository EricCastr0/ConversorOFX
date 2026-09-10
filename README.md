# Conversor OFX

Aplicação desktop desenvolvida em Python para converter relatórios de pagamentos em Excel para o formato OFX.

O projeto surgiu da necessidade de transformar relatórios de recebimentos de cartão em um arquivo que pudesse ser importado em sistemas financeiros, evitando o lançamento manual das movimentações.

A aplicação lê os dados da planilha, gera os lançamentos de crédito referentes aos recebimentos e registra as taxas MDR como débitos separados.

## Funcionalidades

* Seleção de arquivo Excel pela interface
* Leitura da planilha `pagamentos`
* Conversão dos recebimentos para lançamentos OFX
* Registro das taxas MDR como débitos
* Tratamento de datas inválidas
* Utilização da data de recebimento quando a data original da venda não estiver disponível
* Geração do arquivo `.ofx`
* Escolha do local onde o arquivo será salvo
* Interface desktop simples

## Tecnologias utilizadas

* Python
* Pandas
* CustomTkinter
* OpenPyXL

## Como funciona

O arquivo Excel é processado pelo Pandas e cada linha válida da planilha é transformada em uma ou mais movimentações no arquivo OFX.

O valor bruto da parcela é registrado como crédito:

```text
CREDIT
```

Quando existe um valor de MDR descontado, a aplicação também gera uma movimentação separada como débito:

```text
DEBIT
```

A descrição do crédito utiliza informações do estabelecimento, bandeira e modalidade da transação.

Exemplo:

```text
Loja Exemplo - VISA CREDITO
```

O débito referente à taxa é identificado separadamente:

```text
MDR Descontado - Loja Exemplo
```

## Estrutura esperada do Excel

O arquivo deve ser do formato `.xlsx` e possuir uma planilha chamada:

```text
pagamentos
```

A aplicação trabalha com um layout específico de **30 colunas**, nesta ordem:

1. Data do recebimento
2. Data original da venda
3. Data original de vencimento
4. Valor bruto da parcela original
5. Valor bruto da parcela atualizada
6. Taxa MDR
7. Valor MDR descontado
8. Valor líquido da parcela
9. Negociada
10. %
11. NSU/CV
12. TID
13. Número do pedido
14. Número da autorização
15. Resumo de vendas/número do lote
16. Nome do estabelecimento
17. Estabelecimento
18. Número do cartão
19. Indicador de transação tokenizada
20. Código IATA
21. Modalidade
22. Bandeira
23. Número de parcelas
24. Parcela
25. Banco
26. Agência
27. Conta-corrente
28. Cancelamento/contestação
29. Data do cancelamento
30. Status

A ordem dessas colunas é importante, pois atualmente o mapeamento é feito com base na posição das informações dentro da planilha.

## Tratamento das datas

Para definir a data do lançamento, a aplicação utiliza primeiro a **data original da venda**.

Quando essa informação não está disponível ou não pode ser convertida para uma data válida, é utilizada a **data do recebimento**.

Linhas em que nenhuma das duas datas pode ser identificada são ignoradas durante a geração do arquivo.

## Instalação

Clone o repositório:

```bash
git clone https://github.com/EricCastr0/ConversorOFX.git
```

Entre na pasta:

```bash
cd ConversorOFX
```

Crie um ambiente virtual:

```bash
python -m venv .venv
```

No Windows:

```bash
.venv\Scripts\activate
```

No Linux ou macOS:

```bash
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

## Executando

Inicie a aplicação com:

```bash
python appConversor.py
```

Na interface:

1. Clique em **Selecionar Arquivo Excel**
2. Escolha o relatório `.xlsx`
3. Clique em **Converter para OFX**
4. Escolha onde o arquivo será salvo

Ao final do processamento, a aplicação informa se o arquivo foi gerado corretamente.

## Estrutura do projeto

```text
ConversorOFX/
├── appConversor.py
├── ofx_generator.py
├── requirements.txt
├── logo.ico
└── README.md
```

### `appConversor.py`

Responsável pela interface, leitura do arquivo Excel, validação dos dados e processo de conversão.

### `ofx_generator.py`

Responsável por montar a estrutura do arquivo OFX e gerar as movimentações de crédito e débito.

## Formato OFX

O arquivo é gerado utilizando OFX 1.02 no formato SGML.

Algumas configurações atuais:

```text
Moeda: BRL
Banco: Itaú
Código bancário: 341
Tipo de conta: CHECKING
```

A conta utilizada no OFX é obtida da coluna `conta-corrente` do relatório.

## Observações

Este projeto foi desenvolvido para trabalhar com um layout específico de relatório. Arquivos com estrutura diferente podem exigir ajustes no mapeamento das colunas.

O código do banco está atualmente definido como `341` (Itaú). Para utilizar o conversor com relatórios de outros bancos, essa configuração precisa ser adaptada.

O campo de saldo (`BALAMT`) do arquivo OFX ainda utiliza um valor fixo e não representa o saldo real da conta. O foco atual do conversor está na geração das movimentações para importação.

Por esse motivo, é recomendado validar o arquivo gerado no sistema de destino antes de utilizá-lo em um fluxo definitivo.

## Autor

**Eric Castro**

[GitHub](https://github.com/EricCastr0) • [LinkedIn](https://www.linkedin.com/in/eric-castro-silva/)
