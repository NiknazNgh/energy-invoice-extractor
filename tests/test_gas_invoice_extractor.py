from datetime import date

import pytest

import gas_invoice_extractor as gas


def sample_item(description, amount=1.0, rate=0.0, mcf=0.0, mmbtu=0.0, line_no=1):
    return {
        "Line #": line_no,
        "Prod Date": "May-26",
        "Service Account": "",
        "Customer Name": "",
        "Description": description,
        "MCF": mcf,
        "MMBtu": mmbtu,
        "Rate": rate,
        "Amount": amount,
        "Amount Text": f"${amount:,.2f}",
    }


def test_money_to_float_handles_parentheses():
    assert gas.money_to_float("($403.02)") == pytest.approx(-403.02)
    assert gas.money_to_float("$1,250.50") == pytest.approx(1250.50)


def test_month_dates():
    month, first_day, last_day = gas.get_month_dates("Feb-24")
    assert month == "Feb-24"
    assert first_day == date(2024, 2, 1)
    assert last_day == date(2024, 2, 29)


def test_sales_tax_is_retained_as_extra_charge():
    items = [
        sample_item("Customer Charge", 100.0, line_no=1),
        sample_item("Sales Tax", 8.25, line_no=2),
    ]
    extras = gas.find_extra_items(items)
    assert len(extras) == 1
    assert extras[0]["Description"] == "Sales Tax"


def test_extra_charge_marks_invoice_for_review():
    extras = [sample_item("Sales Tax", 8.25, line_no=9)]
    status, notes = gas.build_review_status_and_notes(extras)
    assert status == "NEEDS REVIEW"
    assert "Sales Tax" in notes


def test_standard_charges_are_ok():
    items = [
        sample_item("Customer Charge"),
        sample_item("Reimbursement of MGRT", line_no=2),
        sample_item("Street & Alley Fee", line_no=3),
        sample_item("Pipeline Safety Fee", line_no=4),
        sample_item("Gas Cost Recovery", line_no=5),
        sample_item("Sales Service Rate", line_no=6),
    ]
    extras = gas.find_extra_items(items)
    status, notes = gas.build_review_status_and_notes(extras)
    assert extras == []
    assert status == "OK"
    assert notes == ""


def test_gas_summary_schema_ends_with_review_fields():
    assert gas.SUMMARY_COLUMNS[-2:] == ["Review Status", "Review Notes"]
    assert "Current Charges" in gas.SUMMARY_COLUMNS
    assert "Billed MMBtu" in gas.SUMMARY_COLUMNS
