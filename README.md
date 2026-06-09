# Extrator de Acórdãos: PDF para Excel

Automação em Python para extrair texto de documentos jurídicos em PDF, identificar o número do processo no padrão CNJ e consolidar os dados em uma planilha Excel.

## Funcionalidades

- Extrai texto de arquivos PDF usando PyMuPDF.
- Localiza automaticamente números de processo no padrão CNJ (com ou sem formatação).
- Robustez CNJ: Busca por sequência de 20 dígitos e formata automaticamente no padrão `0000000-00.0000.0.00.0000` caso esteja sem pontos/hífens.
- Nome do arquivo como fallback caso o CNJ não seja encontrado no conteúdo (com limpeza e formatação de dígitos se houver 20 números).
- Remoção automática de caracteres inválidos do Excel.
- Normalização de espaçamento para ganho de performance.
- Liberação ativa de memória RAM (`fitz.tools.shrink_memory()`) a cada arquivo.
- Processamento paralelo automático (ProcessPoolExecutor) para múltiplos PDFs.
- Execução sequencial otimizada para lotes de até 2 PDFs.
- Gravação direta sobrescrevendo a planilha Excel de saída.
- Logs otimizados exibindo progresso periódico (a cada 10% ou 50 arquivos) e resumo final.

## Uso

Coloque os arquivos PDF na pasta `./pdfs` e execute:

```bash
python pdf_to_excel.pyw
```

### Comportamento Padrão
- **Entrada**: Busca arquivos `.pdf` na pasta `./pdfs` (não recursivo).
- **Saída**: Grava os dados em `./processos_extraidos.xlsx` (sobrescreve arquivos existentes).
- **Paralelismo**: Usa processos paralelos automaticamente caso existam mais de 2 arquivos.

## Dependências

- Python 3.8+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install pymupdf pandas openpyxl
```