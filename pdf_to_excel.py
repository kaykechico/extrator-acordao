import re
import logging
import time
from pathlib import Path
from typing import List, Dict, Optional
import fitz
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

CNJ_PATTERN = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")
ILLEGAL_XML_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')

def clean_extracted_text(text: str) -> str:
    text = ILLEGAL_XML_CHARS.sub('', text)
    text = re.sub(r'_{3,}', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    EXCEL_LIMIT = 32600
    if len(text) > EXCEL_LIMIT:
        text = text[:EXCEL_LIMIT] + " ... [AVISO: TEXTO TRUNCADO DEVIDO AO LIMITE DE 32 MIL CARACTERES DO EXCEL]"
        
    return text

def find_cnj_number(text: str) -> str:
    match = CNJ_PATTERN.search(text)
    return match.group(0) if match else "Não localizado"

def extract_text_from_pdf(pdf_path: Path) -> Optional[str]:
    text_content: List[str] = []
    
    try:
        with fitz.open(pdf_path) as doc:
            if doc.is_encrypted:
                logger.error(f"Arquivo protegido por senha ignorado: {pdf_path.name}")
                return None
            
            for page in doc:
                text_content.append(page.get_text())
                
    except fitz.FileDataError:
        logger.error(f"Arquivo corrompido ou de formato inválido ignorado: {pdf_path.name}")
        return None
    except Exception as e:
        logger.error(f"Erro I/O inesperado ao processar '{pdf_path.name}': {e}")
        return None

    raw_text = " ".join(text_content)
    
    if not raw_text.strip():
         logger.warning(f"O arquivo parece estar vazio ou é apenas imagem: {pdf_path.name}")
         return "" 

    return clean_extracted_text(raw_text)

def process_directory(input_dir: Path) -> List[Dict[str, str]]:
    dataset: List[Dict[str, str]] = []
    
    if not input_dir.is_dir():
        logger.error(f"O diretório informado não existe: {input_dir}")
        return dataset

    pdf_files = list(input_dir.glob("*.pdf"))
    logger.info(f"Iniciando processamento. Total de PDFs encontrados: {len(pdf_files)}")

    for pdf_path in pdf_files:
        logger.info(f"Processando: {pdf_path.name}")
        
        text = extract_text_from_pdf(pdf_path)
        
        if text is None:
            continue
            
        cnj_number = find_cnj_number(text)
        if cnj_number == "Não localizado":
            cnj_number = find_cnj_number(pdf_path.name)
        
        dataset.append({
            "Número do Processo": cnj_number,
            "Conteúdo": text,
            "Título do Arquivo": pdf_path.name
        })

    return dataset

def export_to_excel(data: List[Dict[str, str]], output_path: Path) -> None:
    if not data:
        logger.error("Nenhum dado válido para exportar. O job abortou.")
        return

    try:
        df = pd.DataFrame(data)
        
        columns_order = ["Número do Processo", "Conteúdo", "Título do Arquivo"]
        df = df[columns_order]
        
        df.to_excel(output_path, index=False, engine='openpyxl')
        logger.info(f"Arquivo Excel gerado com sucesso em: {output_path.resolve()}")
    except Exception as e:
        logger.error(f"Falha crítica ao gravar o arquivo Excel: {e}")

if __name__ == "__main__":
    INPUT_DIRECTORY = Path("./pdfs")
    OUTPUT_EXCEL = Path("./processos_extraidos.xlsx")
    
    INPUT_DIRECTORY.mkdir(parents=True, exist_ok=True)

    start_time = time.perf_counter()
    logger.info("--- START: Extração de PDFs e Geração de Excel ---")

    extracted_data = process_directory(INPUT_DIRECTORY)
    export_to_excel(extracted_data, OUTPUT_EXCEL)

    end_time = time.perf_counter()
    logger.info(f"--- END: Finalizado. Tempo de execução: {end_time - start_time:.3f}s ---")
    
    input("\nPressione ENTER para fechar a janela...")