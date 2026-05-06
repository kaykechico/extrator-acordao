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