# Extrator de Acórdãos: PDF para Excel

Automação em Python para extrair texto de documentos jurídicos em PDF e consolidar os dados em uma planilha Excel.

## Funcionalidades

- Extrai o texto dos arquivos PDF usando PyMuPDF.
- Gera uma linha por PDF com as colunas: `Título do Arquivo`, `Número do Processo`, `Outros Processos`, `Status` (Sucesso / Criptografado / Sem texto / Erro), `Detalhes` (mensagem de erro) e `Conteúdo`, para rastrear falhas.
- Senha opcional para PDFs criptografados (`-p/--password`).
- Identifica números de processo no padrão CNJ (validados pelos dígitos verificadores, ISO 7064) e os exporta nas colunas `Número do Processo` (primeiro encontrado) e `Outros Processos` (demais).
- `--recursive`: busca PDFs também em subpastas.
- `--dedup`: remove linhas duplicadas do mesmo processo, considerando todos os números CNJ da linha (principal e outros), e mantém a primeira.
- Proteção contra Formula/Injection: células iniciadas com `=`, `+`, `-`, `@`, tab ou carriage return recebem prefixo `'`.
- Remove caracteres inválidos do Excel e normaliza o espaçamento do texto.
- Processamento paralelo automático (até 4 processos); sequencial para lotes de até 2 PDFs.
- Logs periódicos de progresso (a cada 10%, até 50 arquivos) e resumo final.

## Uso

Coloque os arquivos PDF na pasta `./pdfs` e execute:

```bash
python pdf_to_excel.pyw
```

### Comportamento Padrão
- **Entrada**: Busca arquivos `.pdf` na pasta `./pdfs` (não recursivo).
- **Saída**: Grava os dados em `./processos_extraidos.xlsx` (sobrescreve arquivos existentes).
- **Paralelismo**: Usa processos paralelos automaticamente caso existam mais de 2 arquivos.

### Opções de linha de comando

```bash
python pdf_to_excel.pyw [-i DIRETORIO] [-o ARQUIVO.xlsx] [-p SENHA] [--recursive] [--dedup]
```

- `-i, --input`: diretório com os PDFs (padrão: `./pdfs`).
- `-o, --output`: arquivo Excel de saída (padrão: `./processos_extraidos.xlsx`).
- `-p, --password`: senha para abrir PDFs criptografados.
- `--recursive`: busca PDFs também em subpastas.
- `--dedup`: remove linhas duplicadas do mesmo processo, considerando todos os números CNJ da linha (principal e outros). Mantém a primeira.

## Dependências

- Python 3.11+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`) 1.27.2.2
- [pandas](https://pandas.pydata.org/) 3.0.0
- [openpyxl](https://openpyxl.readthedocs.io/) 3.1.5

```bash
pip install -r requirements.txt
```