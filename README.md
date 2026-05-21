# Extrator de Acórdãos: PDF para Excel

Automação em Python para extrair texto de documentos jurídicos em PDF, identificar o número do processo no padrão CNJ e consolidar os dados em uma planilha Excel.

O projeto foi criado para processar arquivos em lote dentro de uma pasta, limpar o texto extraído e gerar um arquivo `.xlsx` organizado com as informações principais.

## Funcionalidades

- Extrai texto de arquivos PDF usando PyMuPDF.
- Localiza automaticamente números de processo no padrão CNJ.
- Usa o nome do arquivo como fallback quando o CNJ não é encontrado no conteúdo.
- Remove caracteres inválidos que podem quebrar a geração do Excel.
- Normaliza espaços e quebras de linha.
- Evita erro no Excel ao limitar textos acima de 32.767 caracteres por célula.
- Processa múltiplos PDFs em paralelo para melhorar a performance.
- Aceita arquivos com extensão `.pdf`, `.PDF` e variações similares.
- Permite busca opcional em subpastas.
- Gera uma planilha Excel com aba chamada `Dados`.
- Exibe uma janela de console com logs mesmo quando executado como `.pyw`.
- Diferencia PDFs criptografados, vazios e com erro no log de processamento.
- Exibe progresso individual `[N/TOTAL]` e resumo detalhado ao final.

## Uso

```bash
python pdf_to_excel.pyw [opções]
```

### Opções

| Flag          | Descrição                                              | Padrão                       |
|---------------|--------------------------------------------------------|------------------------------|
| `--input`     | Diretório contendo os PDFs                             | `./pdfs`                     |
| `--output`    | Caminho do arquivo Excel de saída                      | `./processos_extraidos.xlsx` |
| `--recursive` | Busca PDFs também em subpastas                         | Desativado                   |
| `--workers`   | Número de processos paralelos                          | Automático (nº de CPUs)      |
| `--force`     | Sobrescreve o arquivo Excel de saída caso já exista    | Desativado                   |
| `--verbose`   | Exibe logs detalhados (nível DEBUG)                    | Desativado                   |
| `--quiet`     | Exibe apenas avisos e erros                            | Desativado                   |

### Exemplos

```bash
# Uso básico (primeira execução)
python pdf_to_excel.pyw

# Sobrescrever Excel existente
python pdf_to_excel.pyw --force

# Buscar em subpastas com 4 workers
python pdf_to_excel.pyw --recursive --workers 4 --force

# Modo silencioso, apenas avisos e erros
python pdf_to_excel.pyw --quiet --force

# Diretório e saída personalizados
python pdf_to_excel.pyw --input ./meus_pdfs --output ./resultado.xlsx --force
```

## Dependências

- Python 3.8+
- [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install pymupdf pandas openpyxl
```