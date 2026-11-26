# Extrator de Acórdãos: PDF para Excel

Automação em Python para extrair texto de documentos jurídicos em PDF e consolidar os dados em uma planilha Excel.

## Funcionalidades

- Extrai o texto dos arquivos PDF usando PyMuPDF.
- Gera uma linha por PDF com as colunas: `Título do Arquivo`, `Número do Processo`, `Outros Processos`, `Status` (Sucesso / Criptografado / Sem texto / Erro), `Detalhes` (mensagem de erro ou aviso) e `Conteúdo`, para rastrear falhas.
- Textos maiores que o limite de uma célula do Excel (32.767 caracteres) continuam nas colunas `Conteúdo (2)`, `Conteúdo (3)`... — nada é cortado. `Detalhes` informa quando isso acontece.
- Senha opcional para PDFs criptografados (`-p/--password`).
- Identifica números de processo no padrão CNJ (validados pelos dígitos verificadores, ISO 7064) e os exporta nas colunas `Número do Processo` (primeiro encontrado) e `Outros Processos` (demais citados no texto).
- `--recursive`: busca PDFs também em subpastas. `Título do Arquivo` mostra o caminho relativo (ex.: `2024\acordao.pdf`).
- `--dedup`: remove linhas com o mesmo `Número do Processo` e mantém a primeira. Processos apenas citados (`Outros Processos`) não contam.
- `--ocr`: aplica OCR nas páginas sem texto (PDFs escaneados). Requer o [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) instalado com o idioma português (`por`). Se o OCR falhar, o texto das demais páginas é mantido e o motivo aparece em `Detalhes`.
- Proteção contra Formula Injection: todas as células são gravadas como texto; as iniciadas com `=`, `+`, `-`, `@`, tab ou carriage return recebem prefixo de texto (`'`) do Excel.
- Remove caracteres inválidos do Excel e normaliza o espaçamento do texto.
- Processamento paralelo automático (até 8 processos), dimensionado pelo volume: cerca de 1 processo a cada 10 arquivos ou 250 páginas; lotes pequenos rodam em sequência, que é mais rápido que abrir processos. Com `--ocr`, 1 processo por arquivo.
- Verifica antes de processar se o Excel de saída está aberto/bloqueado.
- Logs periódicos de progresso (a cada 10%, até 50 arquivos) e resumo final.
- Código de saída `0` em sucesso (ou pasta sem PDFs) e `1` em falha (útil para agendamentos).

## Uso

Coloque os arquivos PDF na pasta `./pdfs` e execute:

```bash
python pdf_to_excel.pyw
```

### Comportamento Padrão
- **Entrada**: Busca arquivos `.pdf` na pasta `./pdfs` (não recursivo). Se a pasta padrão não existir, ela é criada. Um diretório informado com `-i` que não existe é tratado como erro.
- **Saída**: Grava os dados em `./processos_extraidos.xlsx` (sobrescreve arquivos existentes).
- **Paralelismo**: Usa processos paralelos automaticamente quando o volume compensa (a partir de ~20 arquivos ou ~500 páginas).
- **Janela**: Ao abrir com duplo clique, a janela aguarda ENTER antes de fechar. No terminal, o script termina direto.

### Opções de linha de comando

```bash
python pdf_to_excel.pyw [-i DIRETORIO] [-o ARQUIVO.xlsx] [-p [SENHA]] [--recursive] [--dedup] [--ocr]
```

- `-i, --input`: diretório com os PDFs (padrão: `./pdfs`).
- `-o, --output`: arquivo Excel de saída (padrão: `./processos_extraidos.xlsx`).
- `-p, --password`: senha para abrir PDFs criptografados. Use `-p` sem valor para digitar a senha sem eco (ela não fica no histórico do terminal). Em execuções agendadas, sem teclado, use `-p SENHA`.
- `--recursive`: busca PDFs também em subpastas.
- `--dedup`: remove linhas duplicadas do mesmo `Número do Processo`. Mantém a primeira.
- `--ocr`: OCR nas páginas sem texto (requer Tesseract).

## Dependências

- Python 3.11+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`) 1.27.2.2
- [openpyxl](https://openpyxl.readthedocs.io/) 3.1.5

```bash
pip install -r requirements.txt
```
