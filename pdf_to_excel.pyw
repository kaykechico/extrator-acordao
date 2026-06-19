import logging
import os
import re
import sys
import time
import warnings
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import fitz
import pandas as pd


FIND_CNJ_PATTERN = re.compile(
    r"\b(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})\b|\b(\d{20})\b"
)
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
UNDERSCORES_PATTERN = re.compile(r"_{3,}")
WS_PATTERN = re.compile(r"\s+")
EXCEL_INJECTION = re.compile(r"^[=+\-@]")

logger = logging.getLogger(__name__)


@dataclass
class ExtractionResult:

    status: str
    data: Optional[Dict[str, str]] = None
    error_message: str = ""
    filename: str = ""


def setup_console() -> None:
    if sys.platform != "win32":
        return

    import ctypes
    kernel32 = ctypes.windll.kernel32

    if kernel32.GetConsoleWindow():
        return

    if not kernel32.AllocConsole():
        return

    try:
        stdout_handle = open("CONOUT$", "w", encoding="utf-8", buffering=1)
        stdin_handle = open("CONIN$", "r", encoding="utf-8")

        sys.stdout = stdout_handle
        sys.stderr = stdout_handle
        sys.stdin = stdin_handle

    except OSError:
        pass


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
    text = UNDERSCORES_PATTERN.sub(" ", text)
    text = WS_PATTERN.sub(" ", text)
    return text.strip()


def find_cnj_number(text: str) -> str:
    if not text:
        return "Não localizado"
    match = FIND_CNJ_PATTERN.search(text)
    if match:
        if match.group(1):
            return match.group(1)
        d = match.group(2)
        return f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}"
    return "Não localizado"


def extract_text_from_pdf(pdf_path: Path) -> ExtractionResult:
    try:
        text_content = []

        with fitz.open(pdf_path) as doc:
            if doc.is_encrypted:
                if not doc.authenticate(""):
                    return ExtractionResult(
                        status="encrypted",
                        filename=pdf_path.name
                    )

            for page in doc:
                page_text = page.get_text("text")
                if page_text:
                    text_content.append(page_text)

        cleaned_text = clean_extracted_text(" ".join(text_content))

        if not cleaned_text:
            return ExtractionResult(
                status="empty",
                filename=pdf_path.name
            )

        cnj_number = find_cnj_number(cleaned_text)
        if cnj_number == "Não localizado":
            cnj_number = find_cnj_number(pdf_path.name)

        return ExtractionResult(
            status="success",
            data={
                "Número do Processo": cnj_number,
                "Conteúdo": cleaned_text,
                "Título do Arquivo": pdf_path.name
            },
            filename=pdf_path.name
        )

    except Exception as e:
        return ExtractionResult(
            status="error",
            error_message=str(e),
            filename=pdf_path.name
        )


def get_pdf_files(input_dir: Path) -> List[Path]:
    return sorted(input_dir.glob("*.pdf"))


def process_directory(input_dir: Path) -> List[Dict[str, str]]:
    if not input_dir.is_dir():
        logger.error(f"O diretório informado não existe: {input_dir}")
        return []

    pdf_files = get_pdf_files(input_dir)
    total = len(pdf_files)

    logger.info(f"Total de PDFs encontrados: {total}")

    if not pdf_files:
        return []

    dataset: List[Dict[str, str]] = []
    counters = {"success": 0, "encrypted": 0, "empty": 0, "error": 0}
    processed = 0

    use_parallel = total > 2

    step = max(5, total // 10) if total >= 10 else 1
    log_interval = min(50, step)

    def log_progress() -> None:
        if processed % log_interval == 0 or processed == total:
            logger.info(f"Progresso: {processed}/{total} arquivos ({processed * 100 // total}%)")

    def handle_result(result: ExtractionResult) -> None:
        nonlocal processed
        processed += 1
        counters[result.status] = counters.get(result.status, 0) + 1

        if result.status == "success":
            dataset.append(result.data)
            logger.debug(f"Processado com sucesso: {result.filename}")
        elif result.status == "encrypted":
            logger.warning(
                f"[{processed}/{total}] PDF criptografado, ignorado: {result.filename}"
            )
        elif result.status == "empty":
            logger.warning(
                f"[{processed}/{total}] Sem texto extraível: {result.filename}"
            )
        elif result.status == "error":
            logger.error(
                f"[{processed}/{total}] Erro ao processar {result.filename}: {result.error_message}"
            )
        else:
            logger.error(
                f"[{processed}/{total}] Status desconhecido '{result.status}' em: {result.filename}"
            )

        log_progress()

    def handle_failure(pdf_path: Path, exc: BaseException) -> None:
        nonlocal processed
        processed += 1
        counters["error"] += 1
        logger.error(
            f"[{processed}/{total}] Falha crítica ao processar {pdf_path.name}: {exc}"
        )
        log_progress()

    if use_parallel:
        with ProcessPoolExecutor(max_workers=min(4, os.cpu_count() or 1)) as executor:
            future_to_pdf = {
                executor.submit(extract_text_from_pdf, pdf_path): pdf_path
                for pdf_path in pdf_files
            }

            for future in as_completed(future_to_pdf):
                pdf_path = future_to_pdf[future]
                try:
                    result = future.result()
                    handle_result(result)
                except Exception as e:
                    handle_failure(pdf_path, e)
    else:
        for pdf_path in pdf_files:
            try:
                result = extract_text_from_pdf(pdf_path)
                handle_result(result)
            except Exception as e:
                handle_failure(pdf_path, e)

    dataset.sort(key=lambda item: item["Título do Arquivo"].lower())

    logger.info("--- Resumo do Processamento ---")
    logger.info(f"  Total encontrados:  {total}")
    logger.info(f"  Sucesso:            {counters['success']}")
    logger.info(f"  Criptografados:     {counters['encrypted']}")
    logger.info(f"  Sem texto:          {counters['empty']}")
    logger.info(f"  Erros:              {counters['error']}")

    return dataset


def export_to_excel(data: List[Dict[str, str]], output_path: Path) -> None:
    if not data:
        logger.error("Nenhum dado válido para exportar.")
        return

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)

        df = pd.DataFrame(data)

        df = df[["Número do Processo", "Conteúdo", "Título do Arquivo"]]

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore",
                    message=r"Cell contents too long.*",
                    category=UserWarning,
                )
                df.to_excel(writer, sheet_name="Dados", index=False)

            worksheet = writer.sheets["Dados"]
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and EXCEL_INJECTION.match(cell.value):
                        cell.quotePrefix = True

        logger.info(f"Arquivo Excel gerado em: {output_path.resolve()}")

    except (PermissionError, OSError) as e:
        logger.exception(
            f"Não foi possível gravar o arquivo Excel. "
            f"Verifique se ele está aberto: {output_path} ({type(e).__name__})"
        )

    except Exception as e:
        logger.exception(f"Falha crítica ao gravar o arquivo Excel: {e}")


def main() -> None:
    setup_console()
    setup_logging()

    base_dir = Path(__file__).resolve().parent
    input_path = base_dir / "pdfs"
    output_path = base_dir / "processos_extraidos.xlsx"

    if not input_path.is_dir():
        input_path.mkdir(parents=True, exist_ok=True)
        logger.warning(
            f"Diretório de entrada criado: {input_path.resolve()}. "
            f"Coloque os PDFs e execute novamente."
        )
        return

    start_time = time.perf_counter()

    logger.debug(f"Diretório de entrada: {input_path.resolve()}")
    logger.debug(f"Arquivo de saída: {output_path.resolve()}")

    extracted_data = process_directory(input_dir=input_path)

    export_to_excel(extracted_data, output_path)

    end_time = time.perf_counter()

    logger.debug(f"Total exportado: {len(extracted_data)}")
    logger.info(f"--- Finalizado. Tempo de execução: {end_time - start_time:.3f}s ---")


if __name__ == "__main__":
    try:
        main()
    finally:
        input("\nPressione ENTER para fechar a janela...")