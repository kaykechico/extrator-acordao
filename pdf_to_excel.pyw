import argparse
import getpass
import importlib.util
import logging
import multiprocessing
import os
import re
import sys
import threading
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pymupdf as fitz


FIND_CNJ_PATTERN = re.compile(
    r"(?<!\d)\d{7}(?:-\s?\d{2}\.\s?\d{4}\.\s?\d\.\s?\d{2}\.\s?\d{4}|\d{13})(?!\d)"
)
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
UNDERSCORES_PATTERN = re.compile(r"_{3,}")
EXCEL_INJECTION = re.compile(r"^[=+\-@\t\r]")

CNJ_NOT_FOUND = "Não localizado"
EXCEL_CELL_LIMIT = 32767
OCR_LANGUAGE = "por"
ASK_PASSWORD = object()

COLUMNS = [
    "Título do Arquivo",
    "Número do Processo",
    "Outros Processos",
    "Status",
    "Detalhes",
    "Conteúdo",
]

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


def setup_console() -> bool:
    if sys.platform != "win32":
        return False

    import ctypes
    kernel32 = ctypes.windll.kernel32

    if kernel32.GetConsoleWindow():
        return False

    if not kernel32.AllocConsole():
        return False

    try:
        stdout_handle = open("CONOUT$", "w", encoding="utf-8", buffering=1)
        stdin_handle = open("CONIN$", "r", encoding="utf-8")

        sys.stdout = stdout_handle
        sys.stderr = stdout_handle
        sys.stdin = sys.__stdin__ = stdin_handle

    except OSError:
        pass

    return True


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
    return " ".join(text.split())


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
    for candidate in FIND_CNJ_PATTERN.findall(text):
        candidate = "".join(candidate.split())
        if len(candidate) == 20:
            d = candidate
            candidate = f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}"
        if candidate in seen:
            continue
        seen.add(candidate)
        if validate_cnj(candidate):
            found.append(candidate)
    return found


def extract_text_from_pdf(
    pdf_path: Path,
    password: Optional[str] = None,
    display_name: Optional[str] = None,
    ocr: bool = False,
) -> ExtractionResult:
    display_name = display_name or pdf_path.name
    try:
        text_content = []
        ocr_error = ""

        with fitz.open(pdf_path) as doc:
            if doc.is_encrypted:
                if not doc.authenticate(password or ""):
                    return ExtractionResult(
                        status="encrypted",
                        filename=display_name
                    )

            for page in doc:
                page_text = page.get_text("text")
                if ocr and not ocr_error and not page_text.strip():
                    try:
                        textpage = page.get_textpage_ocr(language=OCR_LANGUAGE, full=True)
                        page_text = page.get_text("text", textpage=textpage)
                    except Exception as e:
                        ocr_error = f"OCR falhou: {e}"
                if page_text:
                    text_content.append(page_text)

        cleaned_text = clean_extracted_text(" ".join(text_content))

        if not cleaned_text:
            return ExtractionResult(
                status="error" if ocr_error else "empty",
                error_message=ocr_error,
                filename=display_name
            )

        text_numbers = find_cnj_numbers(cleaned_text)
        name_numbers = find_cnj_numbers(pdf_path.name)
        main_number = (name_numbers or text_numbers or [CNJ_NOT_FOUND])[0]
        other_numbers = [n for n in text_numbers if n != main_number]

        return ExtractionResult(
            status="success",
            data={
                "Número do Processo": main_number,
                "Outros Processos": "; ".join(other_numbers),
                "Conteúdo": cleaned_text,
            },
            error_message=ocr_error,
            filename=display_name
        )

    except Exception as e:
        return ExtractionResult(
            status="error",
            error_message=str(e),
            filename=display_name
        )


def get_pdf_files(input_dir: Path, recursive: bool = False) -> List[Path]:
    pattern = "**/*" if recursive else "*"
    return sorted(
        p for p in input_dir.glob(pattern) if p.suffix.lower() == ".pdf" and p.is_file()
    )


def count_pages(paths: List[Path]) -> int:
    pages = 0
    for path in paths:
        try:
            with fitz.open(path) as doc:
                pages += doc.page_count
        except Exception:
            pass
    return pages


def dedup_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    seen = set()
    unique = []
    for row in rows:
        number = row["Número do Processo"]
        if number and number != CNJ_NOT_FOUND:
            if number in seen:
                continue
            seen.add(number)
        unique.append(row)
    return unique


def process_directory(
    input_dir: Path,
    password: Optional[str] = None,
    recursive: bool = False,
    dedup: bool = False,
    ocr: bool = False,
) -> List[Dict[str, str]]:
    pdf_files = get_pdf_files(input_dir, recursive)
    total = len(pdf_files)

    logger.info(f"Total de PDFs encontrados: {total}")

    if not pdf_files:
        return []

    dataset: List[Dict[str, str]] = []
    counters = {"success": 0, "encrypted": 0, "empty": 0, "error": 0}
    processed = 0

    max_workers = min(8, os.cpu_count() or 1)
    if ocr:
        workers = min(max_workers, total)
    else:
        workers = min(max_workers, total // 10)
        if workers < max_workers:
            workers = min(max_workers, max(workers, count_pages(pdf_files) // 250))

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
        counters[result.status] += 1
        dataset.append(result.to_row())

        if result.status == "success":
            logger.debug(f"Processado com sucesso: {result.filename}")
            if result.error_message:
                logger.warning(
                    f"[{processed}/{total}] {result.filename}: {result.error_message}"
                )
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

    def handle_failure(display_name: str, exc: BaseException) -> None:
        nonlocal processed
        processed += 1
        counters["error"] += 1
        dataset.append(
            ExtractionResult(
                status="error",
                error_message=str(exc),
                filename=display_name,
            ).to_row()
        )
        logger.error(
            f"[{processed}/{total}] Falha crítica ao processar {display_name}: {exc}"
        )
        log_progress()

    jobs = [(p, str(p.relative_to(input_dir))) for p in pdf_files]

    if workers > 1:
        with ProcessPoolExecutor(
            max_workers=workers, mp_context=multiprocessing.get_context("spawn")
        ) as executor:
            future_to_name = {
                executor.submit(extract_text_from_pdf, pdf_path, password, name, ocr): name
                for pdf_path, name in jobs
            }

            for future in as_completed(future_to_name):
                name = future_to_name[future]
                try:
                    handle_result(future.result())
                except Exception as e:
                    handle_failure(name, e)
    else:
        for pdf_path, name in jobs:
            try:
                handle_result(extract_text_from_pdf(pdf_path, password, name, ocr))
            except Exception as e:
                handle_failure(name, e)

    dataset.sort(key=lambda item: item["Título do Arquivo"].lower())

    if dedup:
        before = len(dataset)
        dataset = dedup_rows(dataset)
        logger.info(f"Linhas duplicadas removidas: {before - len(dataset)}")

    logger.info("--- Resumo do Processamento ---")
    logger.info(f"  Total encontrados:  {total}")
    logger.info(f"  Sucesso:            {counters['success']}")
    logger.info(f"  Criptografados:     {counters['encrypted']}")
    logger.info(f"  Sem texto:          {counters['empty']}")
    logger.info(f"  Erros:              {counters['error']}")

    return dataset


def split_content(text: str) -> List[str]:
    if len(text.encode("utf-16-le", "surrogatepass")) == 2 * len(text):
        return [
            text[i:i + EXCEL_CELL_LIMIT] for i in range(0, len(text), EXCEL_CELL_LIMIT)
        ] or [""]
    chunks, start, units = [], 0, 0
    for i, char in enumerate(text):
        size = 2 if ord(char) > 0xFFFF else 1
        if units + size > EXCEL_CELL_LIMIT:
            chunks.append(text[start:i])
            start, units = i, 0
        units += size
    chunks.append(text[start:])
    return chunks


def row_values(row: Dict[str, str], chunks: List[str]) -> List[str]:
    details = row["Detalhes"]
    if len(chunks) > 1:
        note = f"Conteúdo dividido em {len(chunks)} colunas ({len(row['Conteúdo'])} caracteres)."
        details = f"{details} {note}".strip()
    return [
        row["Título do Arquivo"],
        row["Número do Processo"],
        row["Outros Processos"],
        row["Status"],
        details,
    ] + chunks


def export_to_excel(data: List[Dict[str, str]], output_path: Path) -> bool:
    if not data:
        logger.error("Nenhum dado válido para exportar.")
        return False

    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)

        content_chunks = [split_content(row["Conteúdo"]) for row in data]
        extra_columns = max(len(chunks) for chunks in content_chunks) - 1
        header = COLUMNS + [f"Conteúdo ({n})" for n in range(2, extra_columns + 2)]

        workbook = Workbook(write_only=True)
        worksheet = workbook.create_sheet("Dados")

        def text_cell(value: str) -> WriteOnlyCell:
            cell = WriteOnlyCell(worksheet, value)
            cell.data_type = "s"
            if EXCEL_INJECTION.match(value):
                cell.quotePrefix = True
            return cell

        worksheet.append([text_cell(name) for name in header])
        for row, chunks in zip(data, content_chunks):
            worksheet.append([text_cell(value) for value in row_values(row, chunks)])

        tmp_path = output_path.with_name(output_path.name + ".tmp")
        try:
            workbook.save(tmp_path)
            os.replace(tmp_path, output_path)
        finally:
            tmp_path.unlink(missing_ok=True)

        logger.info(f"Arquivo Excel gerado em: {output_path.resolve()}")
        return True

    except (PermissionError, OSError) as e:
        logger.exception(
            f"Não foi possível gravar o arquivo Excel. "
            f"Verifique se ele está aberto: {output_path} ({type(e).__name__})"
        )

    except Exception as e:
        logger.exception(f"Falha crítica ao gravar o arquivo Excel: {e}")

    return False


def check_output_writable(output_path: Path) -> bool:
    if output_path.is_dir():
        logger.error(f"O caminho de saída é uma pasta, informe um arquivo .xlsx: {output_path}")
        return False
    if not output_path.exists():
        return True
    try:
        with open(output_path, "ab"):
            pass
        return True
    except OSError as e:
        logger.error(
            f"Não é possível gravar em {output_path} ({type(e).__name__}). "
            f"Feche o arquivo no Excel e execute novamente."
        )
        return False


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
        nargs="?", const=ASK_PASSWORD,
        help="Senha para abrir PDFs criptografados. Sem valor, a senha é pedida "
             "sem eco (não fica no histórico do terminal).",
    )
    parser.add_argument(
        "--recursive", action="store_true",
        help="Buscar PDFs também em subpastas.",
    )
    parser.add_argument(
        "--dedup", action="store_true",
        help="Remover linhas duplicadas do mesmo Número do Processo.",
    )
    parser.add_argument(
        "--ocr", action="store_true",
        help="Aplicar OCR (Tesseract) nas páginas sem texto, como PDFs escaneados.",
    )
    return parser.parse_args()


def main(args: argparse.Namespace) -> int:
    setup_logging()

    base_dir = Path(__file__).resolve().parent
    input_path = Path(args.input) if args.input else base_dir / "pdfs"
    output_path = (
        Path(args.output) if args.output else base_dir / "processos_extraidos.xlsx"
    )

    if not input_path.is_dir():
        if args.input:
            logger.error(f"O diretório informado não existe: {input_path.resolve()}")
            return 1
        input_path.mkdir(parents=True, exist_ok=True)
        logger.warning(
            f"Diretório de entrada criado: {input_path.resolve()}. "
            f"Coloque os PDFs e execute novamente."
        )
        return 0

    if not check_output_writable(output_path):
        return 1

    if importlib.util.find_spec("openpyxl") is None:
        logger.error("Dependência ausente: openpyxl. Execute: pip install -r requirements.txt")
        return 1

    password = args.password
    if password is ASK_PASSWORD:
        password = getpass.getpass("Senha dos PDFs: ")

    start_time = time.perf_counter()
    threading.Thread(target=__import__, args=("openpyxl",)).start()

    logger.debug(f"Diretório de entrada: {input_path.resolve()}")
    logger.debug(f"Arquivo de saída: {output_path.resolve()}")

    extracted_data = process_directory(
        input_dir=input_path,
        password=password,
        recursive=args.recursive,
        dedup=args.dedup,
        ocr=args.ocr,
    )

    if not extracted_data:
        logger.warning("Nenhum PDF encontrado. Nenhum arquivo Excel foi gerado.")
        return 0

    exported = export_to_excel(extracted_data, output_path)

    end_time = time.perf_counter()

    logger.debug(f"Total exportado: {len(extracted_data)}")
    logger.info(f"--- Finalizado. Tempo de execução: {end_time - start_time:.3f}s ---")
    return 0 if exported else 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    console_created = setup_console()
    args = parse_args()
    exit_code = 1
    try:
        exit_code = main(args)
    except Exception:
        logger.exception("Erro inesperado")
    finally:
        if console_created:
            try:
                input("\nPressione ENTER para fechar a janela...")
            except (EOFError, OSError):
                pass
    sys.exit(exit_code)
