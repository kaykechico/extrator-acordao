# Extrator de Acórdãos: De PDF para Excel

Este projeto é uma solução em Python desenvolvida para automatizar a extração de dados de documentos jurídicos em PDF. O script processa os arquivos em lote, higieniza o texto, localiza o número do processo (padrão CNJ) e consolida as informações em uma planilha Excel de forma rápida e organizada.

## ✨ Principais Funcionalidades

* **Alta Performance:** Utiliza a biblioteca `PyMuPDF` (`fitz`), permitindo o processamento de centenas de páginas em poucos segundos.
* **Busca Inteligente (Fallback):** Localiza automaticamente processos no padrão CNJ dentro do texto. Caso o PDF seja apenas uma imagem (documento escaneado sem OCR), o script tenta extrair a informação a partir do nome do arquivo.
* **Limpeza de Dados:** Remove caracteres inválidos, espaços duplos e marcações que poderiam corromper a estrutura do arquivo Excel.
* **Prevenção de Falhas:** O Excel possui um limite rígido de 32.767 caracteres por célula. O script identifica textos que ultrapassam esse limite, realiza um corte seguro e insere um aviso, evitando que a automação seja interrompida.

## 🛠️ Pré-requisitos

Certifique-se de ter o Python instalado em sua máquina. Para instalar as bibliotecas necessárias, execute o comando abaixo no seu terminal:

```bash
pip install PyMuPDF pandas openpyxl