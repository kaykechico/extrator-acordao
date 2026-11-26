# Extrator de Acórdãos: PDF para Excel

[![Testes](https://github.com/kaykechico/extrator-acordao/actions/workflows/tests.yml/badge.svg)](https://github.com/kaykechico/extrator-acordao/actions/workflows/tests.yml)

Lê os acórdãos em PDF de uma pasta e gera uma planilha Excel com uma linha por arquivo: número do processo, outros processos citados e o texto completo.

## Como usar

1. Instale o Python 3.11+ e as dependências:

   ```powershell
   pip install -r requirements.txt
   ```

2. Coloque os PDFs na pasta `pdfs`, ao lado do script. Se ela não existir, a primeira execução cria.
3. Dê duplo clique em `pdf_to_excel.pyw` (Windows) ou execute `python3 pdf_to_excel.pyw` (Linux).

A planilha `processos_extraidos.xlsx` aparece na mesma pasta. Se ela estiver aberta no Excel, o script avisa e para antes de processar.

## O que sai na planilha

| Coluna | Conteúdo |
|---|---|
| Título do Arquivo | Nome do PDF |
| Número do Processo | CNJ do nome do arquivo; se o nome não tiver, o primeiro CNJ válido do texto |
| Outros Processos | Demais números CNJ citados no texto |
| Status | Sucesso, Criptografado, Sem texto ou Erro |
| Detalhes | Motivo do erro ou aviso |
| Conteúdo | Texto completo do acórdão |

Todo número CNJ passa pela conferência do dígito verificador. Texto acima de 32.767 caracteres, o limite de uma célula do Excel, continua nas colunas `Conteúdo (2)`, `Conteúdo (3)` e assim por diante.

## Opções

```powershell
python pdf_to_excel.pyw [-i PASTA] [-o ARQUIVO.xlsx] [-p [SENHA]] [--recursive] [--dedup] [--ocr]
```

| Opção | O que faz |
|---|---|
| `-i` | Pasta dos PDFs (padrão: `pdfs`) |
| `-o` | Planilha de saída (padrão: `processos_extraidos.xlsx`) |
| `-p` | Senha de PDFs protegidos. Sem valor, pede a senha sem mostrar o que você digita |
| `--recursive` | Inclui subpastas |
| `--dedup` | Mantém uma linha por número de processo |
| `--ocr` | Lê páginas escaneadas. Requer o [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) com o idioma português |

O script termina com código `0` quando dá certo e `1` quando falha, o que permite usá-lo em tarefas agendadas.

## Licença

O código deste repositório é MIT. A leitura de PDF usa o [PyMuPDF](https://pymupdf.readthedocs.io/), que tem licença AGPL-3.0 (ou licença comercial da Artifex): quem distribuir ou oferecer a ferramenta como serviço também segue as regras da AGPL.

## Testes

```powershell
pip install -r requirements.txt pytest
python -m pytest tests
```
