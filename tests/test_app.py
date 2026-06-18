from io import BytesIO
from math import nan
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("flask")

from openpyxl import Workbook, load_workbook

from app import calculate_score, is_blank, is_numeric, process_workbook


def test_blank_detection_matches_rules():
    assert is_blank(None)
    assert is_blank("")
    assert is_blank("   ")
    assert is_blank(nan)
    assert not is_blank("text")


def test_numeric_detection_is_strict_and_deterministic():
    assert is_numeric(123)
    assert is_numeric(45.6)
    assert is_numeric("123")
    assert is_numeric("45.6")
    assert not is_numeric("45,6")
    assert not is_numeric("abc")
    assert not is_numeric("")


def test_calculate_score_reference_example():
    assert calculate_score(None, None, "сдан", "сдан") == 0.78


def test_calculate_score_text_in_b_and_m_is_not_an_event():
    assert calculate_score("уволен", "перевод", " СДАН ", " сдан ") == 0.78


def test_calculate_score_with_dismissal_and_transfer():
    assert calculate_score(123, "456", "не сдан", "сдан") == 0.42


def test_calculate_score_with_blank_m_is_zero():
    assert calculate_score(123, None, "не сдан", "не сдан") == 0


def test_process_workbook_inserts_score_after_m_and_preserves_following_columns():
    workbook = Workbook()
    sheet = workbook.active
    headers = [f"H{i}" for i in range(1, 16)]
    for column, value in enumerate(headers, start=1):
        sheet.cell(row=1, column=column, value=value)

    sheet.cell(row=2, column=2, value=None)
    sheet.cell(row=2, column=5, value="сдан")
    sheet.cell(row=2, column=8, value="сдан")
    sheet.cell(row=2, column=13, value=100)
    sheet.cell(row=2, column=14, value="original N")

    source = BytesIO()
    workbook.save(source)
    source.seek(0)

    result = process_workbook(source)
    processed = load_workbook(result)
    processed_sheet = processed.active

    assert processed_sheet.cell(row=1, column=14).value == "Балл"
    assert processed_sheet.cell(row=2, column=14).value == 1.2
    assert processed_sheet.cell(row=1, column=15).value == "H14"
    assert processed_sheet.cell(row=2, column=15).value == "original N"


def test_process_workbook_replaces_existing_score_column():
    workbook = Workbook()
    sheet = workbook.active
    sheet.cell(row=1, column=14, value="Балл")
    sheet.cell(row=2, column=14, value=999)

    source = BytesIO()
    workbook.save(source)
    source.seek(0)

    result = process_workbook(source)
    processed = load_workbook(result)

    assert processed.active.cell(row=1, column=14).value == "Балл"
    assert processed.active.cell(row=2, column=14).value == 0.4


def test_calculate_route_accepts_xlsx_content():
    workbook = Workbook()
    sheet = workbook.active
    sheet.cell(row=2, column=5, value="сдан")

    source = BytesIO()
    workbook.save(source)
    source.seek(0)

    from app import app

    app.config.update(TESTING=True)
    with app.test_client() as client:
        response = client.post(
            "/calculate",
            data={"file": (source, "rating.xlsx")},
            content_type="multipart/form-data",
        )

    assert response.status_code == 200
    assert response.headers["Content-Disposition"].startswith(
        "attachment; filename=rating_with_scores.xlsx"
    )

    processed = load_workbook(BytesIO(response.data))
    assert processed.active.cell(row=1, column=14).value == "Балл"
    assert processed.active.cell(row=2, column=14).value == 0.65
