import importlib.util
from importlib.machinery import SourceFileLoader
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parent.parent / "pdf_to_excel.pyw"
# .pyw is not importable by name, so load it explicitly.
_loader = SourceFileLoader("pdf_to_excel", str(SCRIPT))
_spec = importlib.util.spec_from_loader("pdf_to_excel", _loader)
mod = importlib.util.module_from_spec(_spec)
_loader.exec_module(mod)

VALID_A = "0001234-14.2020.8.19.0001"
VALID_B = "1000000-95.2023.5.01.0010"
WRONG_DV = "0001234-15.2020.8.19.0001"


@pytest.mark.parametrize("number", [VALID_A, VALID_B, "00012341420208190001"])
def test_validate_cnj_accepts_valid(number):
    assert mod.validate_cnj(number)


@pytest.mark.parametrize("number", [WRONG_DV, "123", "", "0001234-14.2020.8.19.000"])
def test_validate_cnj_rejects_invalid(number):
    assert not mod.validate_cnj(number)


def test_find_cnj_numbers_dedups_formats_and_filters():
    text = (
        f"Processo {VALID_A}, repetido {VALID_A}, sem máscara 10000009520235010010, "
        f"com espaço 0001234-14. 2020.8.19.0001 e inválido {WRONG_DV}."
    )
    assert mod.find_cnj_numbers(text) == [VALID_A, VALID_B]


def test_find_cnj_numbers_ignores_longer_digit_runs():
    assert mod.find_cnj_numbers("9" + "00012341420208190001") == []


def test_clean_extracted_text():
    assert mod.clean_extracted_text("a\x00b ____ c\n\n d") == "ab c d"
    assert mod.clean_extracted_text("") == ""


def test_split_content_respects_excel_limit():
    limit = mod.EXCEL_CELL_LIMIT
    assert mod.split_content("") == [""]
    assert mod.split_content("x" * (limit + 1)) == ["x" * limit, "x"]
    # Emoji counts as 2 UTF-16 units in Excel.
    chunks = mod.split_content("😀" * limit)
    assert "".join(chunks) == "😀" * limit
    assert all(len(c.encode("utf-16-le")) // 2 <= limit for c in chunks)


def test_dedup_rows_keeps_first_and_all_not_found():
    rows = [
        {"Número do Processo": VALID_A, "id": 1},
        {"Número do Processo": VALID_A, "id": 2},
        {"Número do Processo": mod.CNJ_NOT_FOUND, "id": 3},
        {"Número do Processo": mod.CNJ_NOT_FOUND, "id": 4},
    ]
    assert [r["id"] for r in mod.dedup_rows(rows)] == [1, 3, 4]


def test_export_to_excel_splits_long_content_and_blocks_formulas(tmp_path):
    from openpyxl import load_workbook

    row = mod.ExtractionResult(
        status="success",
        filename="=cmd.pdf",
        data={"Número do Processo": VALID_A, "Outros Processos": "",
              "Conteúdo": "y" * (mod.EXCEL_CELL_LIMIT + 5)},
    ).to_row()
    out = tmp_path / "out.xlsx"
    assert mod.export_to_excel([row], out)

    ws = load_workbook(out).active
    header = [c.value for c in ws[1]]
    assert header[-2:] == ["Conteúdo", "Conteúdo (2)"]
    values = [c.value for c in ws[2]]
    assert values[0] == "=cmd.pdf"
    assert ws.cell(2, 1).data_type == "s"
    assert len(values[-1]) == 5
