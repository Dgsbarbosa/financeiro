# Calendário Financeiro Anual Automatizado

Este é um programa em Python que automatiza a criação de um calendário financeiro anual a partir de dados de recebimentos e despesas fornecidos em uma planilha Excel.

## Requisitos

- Python 3.x
- Bibliotecas Python: `calendar`, `datetime`, `locale`, `openpyxl`, `re`, `csv`, `time`, `tkinter`, `os`, `sys`, `tqdm`, `pandas`

## Como Usar

1. Preencha a planilha "Contas Basicas" com as informações de Recebimentos e Despesas.
    a. Se os dias dos vencimentos forem iguais o programaira solicitar que seja alterado 
    b. Se o vencimento nao for preenchido ou for "0", a despesa sera colocada em um dia que nao tenha conta anual ou mensal
2. Execute o script `calendario_financeiro.py`.
3. O programa converterá a planilha Excel para arquivos CSV e realizará análises nos dados.
4. Se necessário, o programa solicitará ajustes nos vencimentos para evitar duplicatas.
5. Um calendário financeiro anual será gerado em uma nova planilha Excel.
6. O usuário terá a opção de salvar a planilha em um local desejado.

## Funcionalidades

- Leitura de dados de uma planilha Excel.
- Conversão para arquivos CSV.
- Verificação e ajuste de vencimentos para evitar duplicatas.
- Criação de um calendário financeiro anual.

## Autor

Douglas Barbosa
## Licença Liberada

