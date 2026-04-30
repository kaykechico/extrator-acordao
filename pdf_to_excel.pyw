import argparse
import logging
import multiprocessing
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
EXCEL_INJECTION = re.compile(r"^[=+\-@\t\r]")

CNJ_NOT_FOUND = "Não localizado"

STATUS_LABELS = {
    "success": "Sucesso",
    "encrypted": "Criptografado",
    "empty": "Sem texto",
    "error": "Erro",
}

logger = logging.getLogger(__name__)


@dataclass
class ExtractionResult:

    status: str
    data: Optional[Dict[str, str]] = None
    error_message: str = ""
    filename: str = ""

    def to_row(self) -> Dict[str, str]:
        data = self.data or {}
        return {
            "Número do Processo": data.get("Número do Processo", ""),
            "Outros Processos": data.get("Outros Processos", ""),
            "Conteúdo": data.get("Conteúdo", ""),
            "Título do Arquivo": self.filename,
            "Status": STATUS_LABELS.get(self.status, self.status),
            "Detalhes": self.error_message,
        }


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


def validate_cnj(number: str) -> bool:
    digits = re.sub(r"\D", "", number)
    if len(digits) != 20:
        return False
    n, dd, a, j, tr, o = (
        digits[:7],
        digits[7:9],
        digits[9:13],
        digits[13],
        digits[14:16],
        digits[16:],
    )
    op1 = int(n) % 97
    op2 = int(f"{op1}{a}{j}{tr}") % 97
    return int(f"{op2}{o}{dd}") % 97 == 1


def find_cnj_numbers(text: str) -> List[str]:
    if not text:
        return []
    found, seen = [], set()
    for match in FIND_CNJ_PATTERN.finditer(text):
        if match.group(1):
            candidate = match.group(1)
        else:
            d = match.group(2)
            candidate = f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}"
        if candidate in seen:
            continue
        seen.add(candidate)
        if validate_cnj(candidate):
            found.append(candidate)
    return found


def extract_text_from_pdf(
    pdf_path: Path, password: Optional[str] = None
) -> ExtractionResult:
    try:
        text_content = []

        with fitz.open(pdf_path) as doc:
            if doc.is_encrypted:
                if not doc.authenticate(password or ""):
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

        cnj_numbers = find_cnj_numbers(cleaned_text)
        if not cnj_numbers:
            cnj_numbers = find_cnj_numbers(pdf_path.name)

        return ExtractionResult(
            status="success",
            data={
                "Número do Processo": cnj_numbers[0] if cnj_numbers else CNJ_NOT_FOUND,
                "Outros Processos": "; ".join(cnj_numbers[1:]),
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


def get_pdf_files(input_dir: Path, recursive: bool = False) -> List[Path]:
    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted(
        p for p in input_dir.glob(pattern) if p.suffix.lower() == ".pdf"
    )


def process_directory(
    input_dir: Path,
    password: Optional[str] = None,
    recursive: bool = False,
    dedup: bool = False,
) -> List[Dict[str, str]]:
    if not input_dir.is_dir():
        logger.error(f"O diretório informado não existe: {input_dir}")
        return []

    pdf_files = get_pdf_files(input_dir, recursive)
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
    next_milestone = log_interval

    def log_progress() -> None:
        nonlocal processed, next_milestone
        if processed == total:
            logger.info(f"Progresso: {total}/{total} arquivos (100%)")
            next_milestone = total + log_interval
            return
        while processed >= next_milestone:
            logger.info(
                f"Progresso: {processed}/{total} arquivos ({processed * 100 // total}%)"
            )
            next_milestone += log_interval

    def handle_result(result: ExtractionResult) -> None:
        nonlocal processed
        processed += 1
        counters[result.status] = counters.get(result.status, 0) + 1
        dataset.append(result.to_row())

        if result.status == "success":
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

        log_progress()

    def handle_failure(pdf_path: Path, exc: BaseException) -> None:
        nonlocal processed
        processed += 1
        counters["error"] += 1
        dataset.append(
            ExtractionResult(
                status="error",
                error_message=str(exc),
                filename=pdf_path.name,
            ).to_row()
        )
        logger.error(
            f"[{processed}/{total}] Falha crítica ao processar {pdf_path.name}: {exc}"
        )
        log_progress()

    if use_parallel:
        with ProcessPoolExecutor(max_workers=min(4, os.cpu_count() or 1)) as executor:
            future_to_pdf = {
                executor.submit(extract_text_from_pdf, pdf_path, password): pdf_path
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
                result = extract_text_from_pdf(pdf_path, password)
                handle_result(result)
            except Exception as e:
                handle_failure(pdf_path, e)

    dataset.sort(key=lambda item: item["Título do Arquivo"].lower())

    if dedup:
        seen: set = set()
        unique = []
        for row in dataset:
            numbers = [
                part.strip()
                for part in f"{row['Número do Processo']}; {row['Outros Processos']}".split(";")
                if part.strip() and part.strip() != CNJ_NOT_FOUND
            ]
            if not numbers:
                unique.append(row)
                continue
            fresh = [number for number in numbers if number not in seen]
            if fresh:
                seen.update(fresh)
                unique.append(row)
        dataset = unique

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

        df = df[
            [
                "Título do Arquivo",
                "Número do Processo",
                "Outros Processos",
                "Status",
                "Detalhes",
                "Conteúdo",
            ]
        ]

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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai texto de PDFs e consolida em Excel."
    )
    parser.add_argument(
        "-i", "--input",
        help="Diretório com os PDFs (padrão: pasta 'pdfs' ao lado do script).",
    )
    parser.add_argument(
        "-o", "--output",
        help="Arquivo Excel de saída (padrão: 'processos_extraidos.xlsx').",
    )
    parser.add_argument(
        "-p", "--password",
        help="Senha para abrir PDFs criptografados.",
    )
    parser.add_argument(
        "--recursive", action="store_true",
        help="Buscar PDFs também em subpastas.",
    )
    parser.add_argument(
        "--dedup", action="store_true",
        help="Remover linhas duplicadas do mesmo Número do Processo.",
    )
    return parser.parse_args()


def main(args: argparse.Namespace) -> None:
    setup_console()
    setup_logging()

    base_dir = Path(__file__).resolve().parent
    input_path = Path(args.input) if args.input else base_dir / "pdfs"
    output_path = (
        Path(args.output) if args.output else base_dir / "processos_extraidos.xlsx"
    )

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

    extracted_data = process_directory(
        input_dir=input_path,
        password=args.password,
        recursive=args.recursive,
        dedup=args.dedup,
    )

    export_to_excel(extracted_data, output_path)

    end_time = time.perf_counter()

    logger.debug(f"Total exportado: {len(extracted_data)}")
    logger.info(f"--- Finalizado. Tempo de execução: {end_time - start_time:.3f}s ---")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    args = parse_args()
    try:
        main(args)
    finally:
        try:
            input("\nPressione ENTER para fechar a janela...")
        except (EOFError, OSError):
            pass
