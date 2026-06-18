from io import BytesIO

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("flask")

from openpyxl import Workbook, load_workbook

from app import calculate_score, process_workbook


def test_calculate_score_example():
    assert calculate_score(None, None, "сдан", "сдан") == 0.78


def test_calculate_score_with_dismissal_and_transfer():
    assert calculate_score(123, 456, "не сдан", "не сдан") == 0.35


def test_process_workbook_inserts_score_after_m_and_preserves_following_columns():
    workbook = Workbook()
    sheet = workbook.active
    headers = [f"H{i}" for i in range(1, 16)]
    for column, value in enumerate(headers, start=1):
        sheet.cell(row=1, column=column, value=value)

    sheet.cell(row=2, column=2, value=None)
    sheet.cell(row=2, column=5, value="сдан")
    sheet.cell(row=2, column=8, value="сдан")
    sheet.cell(row=2, column=13, value=None)
    sheet.cell(row=2, column=14, value="original N")

    source = BytesIO()
    workbook.save(source)
    source.seek(0)

    result = process_workbook(source)
    processed = load_workbook(result)
    processed_sheet = processed.active

    assert processed_sheet.cell(row=1, column=14).value == "Балл"
    assert processed_sheet.cell(row=2, column=14).value == 0.78
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


def test_calculate_route_accepts_excel_content_with_any_extension():
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
            data={"file": (source, "rating.anything")},
            content_type="multipart/form-data",
        )

    assert response.status_code == 200
    assert response.headers["Content-Disposition"].startswith(
        "attachment; filename=rating_with_scores.xlsx"
    )

    processed = load_workbook(BytesIO(response.data))
    assert processed.active.cell(row=1, column=14).value == "Балл"
    assert processed.active.cell(row=2, column=14).value == 0.65
