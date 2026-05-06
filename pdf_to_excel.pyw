import argparse
import ctypes
import logging
import os
import re
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import freeze_support
from pathlib import Path
from typing import Dict, List, Optional

import fitz
import pandas as pd


CNJ_PATTERN = re.compile(r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b")
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")

EXCEL_CELL_LIMIT = 32767
TRUNCATION_WARNING = " ... [AVISO: TEXTO TRUNCADO DEVIDO AO LIMITE DE 32.767 CARACTERES DO EXCEL]"

logger = logging.getLogger(__name__)


def setup_console() -> None:
    if os.name != "nt":
        return

    kernel32 = ctypes.windll.kernel32

    if kernel32.GetConsoleWindow():
        return

    kernel32.AllocConsole()

    sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    sys.stdin = open("CONIN$", "r", encoding="utf-8")


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)]
    )


def clean_extracted_text(text: str) -> str:
    if not text:
        return ""

    text = ILLEGAL_XML_CHARS.sub("", text)
    text = re.sub(r"_{3,}", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    if len(text) > EXCEL_CELL_LIMIT:
        allowed_length = EXCEL_CELL_LIMIT - len(TRUNCATION_WARNING)
        text = text[:allowed_length] + TRUNCATION_WARNING

    return text


def find_cnj_number(text: str) -> str:
    if not text:
        return "Não localizado"

    match = CNJ_PATTERN.search(text)
    return match.group(0) if match else "Não localizado"


def extract_text_from_pdf(pdf_path: Path) -> Optional[Dict[str, str]]:
    try:
        text_content: List[str] = []

        with fitz.open(pdf_path) as doc:
            if doc.is_encrypted:
                return None

            for page in doc:
                page_text = page.get_text("text")
                if page_text:
                    text_content.append(page_text)

        raw_text = " ".join(text_content)
        cleaned_text = clean_extracted_text(raw_text)

        cnj_number = find_cnj_number(cleaned_text)

        if cnj_number == "Não localizado":
            cnj_number = find_cnj_number(pdf_path.name)

        return {
            "Número do Processo": cnj_number,
            "Conteúdo": cleaned_text,
            "Título do Arquivo": pdf_path.name
        }

    except Exception:
        return None


def get_pdf_files(input_dir: Path, recursive: bool = False) -> List[Path]:
    file_iterator = input_dir.rglob("*") if recursive else input_dir.glob("*")

    return sorted(
        file_path
        for file_path in file_iterator
        if file_path.is_file() and file_path.suffix.lower() == ".pdf"
    )


def process_directory(
    input_dir: Path,
    recursive: bool = False,
    max_workers: Optional[int] = None
) -> List[Dict[str, str]]:
    if not input_dir.is_dir():
        logger.error(f"O diretório informado não existe: {input_dir}")
        return []

    pdf_files = get_pdf_files(input_dir, recursive=recursive)

    logger.info(f"Total de PDFs encontrados: {len(pdf_files)}")

    if not pdf_files:
        return []

    dataset: List[Dict[str, str]] = []

    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_pdf = {
            executor.submit(extract_text_from_pdf, pdf_path): pdf_path
            for pdf_path in pdf_files
        }

        for future in as_completed(future_to_pdf):
            pdf_path = future_to_pdf[future]

            try:
                result = future.result()

                if result is None:
                    logger.warning(f"Ignorado ou sem texto extraível: {pdf_path.name}")
                    continue

                dataset.append(result)
                logger.info(f"Processado com sucesso: {pdf_path.name}")

            except Exception as e:
                logger.error(f"Falha crítica ao processar {pdf_path.name}: {e}")

    dataset.sort(key=lambda item: item["Título do Arquivo"].lower())

    return dataset


def export_to_excel(data: List[Dict[str, str]], output_path: Path) -> None:
    if not data:
        logger.error("Nenhum dado válido para exportar. O job abortou.")
        return

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)

        df = pd.DataFrame(data)

        columns_order = [
            "Número do Processo",
            "Conteúdo",
            "Título do Arquivo"
        ]

        df = df[columns_order]

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Dados", index=False)

        logger.info(f"Arquivo Excel gerado com sucesso em: {output_path.resolve()}")

    except PermissionError:
        logger.error(
            f"Não foi possível gravar o arquivo Excel. "
            f"Verifique se ele está aberto: {output_path}"
        )

    except Exception as e:
        logger.error(f"Falha crítica ao gravar o arquivo Excel: {e}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai texto e número CNJ de arquivos PDF e exporta para Excel."
    )

    parser.add_argument(
        "--input",
        type=Path,
        default=Path("./pdfs"),
        help="Diretório contendo os PDFs."
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=Path("./processos_extraidos.xlsx"),
        help="Caminho do arquivo Excel de saída."
    )

    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Busca PDFs também em subpastas."
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Número de processos paralelos."
    )

    return parser.parse_args()


def main() -> None:
    setup_console()
    setup_logging()

    args = parse_args()

    if args.workers is not None and args.workers < 1:
        logger.error("O parâmetro --workers deve ser maior ou igual a 1.")
        return

    args.input.mkdir(parents=True, exist_ok=True)

    start_time = time.perf_counter()

    logger.info("--- START: Extração de PDFs e Geração de Excel ---")
    logger.info(f"Diretório de entrada: {args.input.resolve()}")
    logger.info(f"Arquivo de saída: {args.output.resolve()}")

    extracted_data = process_directory(
        input_dir=args.input,
        recursive=args.recursive,
        max_workers=args.workers
    )

    export_to_excel(extracted_data, args.output)

    end_time = time.perf_counter()

    logger.info(f"Total exportado: {len(extracted_data)}")
    logger.info(f"--- END: Finalizado. Tempo de execução: {end_time - start_time:.3f}s ---")


if __name__ == "__main__":
    freeze_support()

    try:
        main()
    finally:
        input("\nPressione ENTER para fechar a janela...")