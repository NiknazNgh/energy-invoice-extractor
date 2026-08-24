from datetime import date

import pytest

import invoice_extractor as ie


def base_row():
    row = {c: None for c in ie.INTERNAL_COLUMNS}
    row.update(
        {
            "Energy Charge": 100.0,
            "Transmission Charges": 50.0,
            "Taxes & PUC Assessment Charge": 10.0,
            "Prior Period Pass Through Charge": 5.0,
            "Other Taxes": 0.0,
            "Bill Total": 165.0,
            "Firm Fuel Supply Service": 0.0,
            "Firm Fuel Supply Service - Backbill": 0.0,
            "Ancilliary Service Obligation Adjustment": 0.0,
        }
    )
    return row


def test_output_column_order():
    assert ie.OUTPUT_COLUMNS[0] == "Production Month"
    assert ie.OUTPUT_COLUMNS[-1] == "Bill Total"
    assert len(ie.OUTPUT_COLUMNS) == 27
    assert ie.FINAL_OUTPUT_COLUMNS[-2:] == ["Review Status", "Review Notes"]


def test_production_month_uses_dominant_month():
    result = ie.production_month_from_range(date(2026, 7, 25), date(2026, 8, 23))
    assert result == date(2026, 8, 1)


def test_percent_conversion():
    assert ie.percent_text_to_decimal("95.2%") == pytest.approx(0.952)
    assert ie.percent_text_to_decimal("95.2") == pytest.approx(0.952)


def test_reconciliation_ok():
    status, notes = ie.build_review(base_row(), "", [])
    assert status == "OK"
    assert "reconciles" in notes.lower()


def test_reconciliation_finds_specific_source_line():
    row = base_row()
    row["Bill Total"] = 351.20
    text = "Special Capacity Adjustment 186.20\n"
    status, notes = ie.build_review(row, text, [])
    assert status == "NEEDS REVIEW"
    assert "$186.20" in notes
    assert "Special Capacity Adjustment 186.20" in notes


def test_no_double_counting_of_nested_4cp():
    row = base_row()
    row["4CP Charges ($)"] = 25.0
    status, _ = ie.build_review(row, "", [])
    assert status == "OK"
