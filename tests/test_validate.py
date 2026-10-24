"""Tests for each validation rule on a small hand-built batch, plus the sample batch."""
from datetime import date

import pandas as pd
import pytest

from src import config as cfg, validate, workbook
from src.validate import normalise_invoice_number

SETTINGS = dict(cfg.SETTINGS, batch_date=date(2026, 10, 23))


def clean_data():
    """One vendor, one PO with two lines, one invoice that passes every check."""
    return {
        "vendors": pd.DataFrame([
            {"vendor_id": "V1", "vendor_name": "Test Supplies Pty Ltd", "abn": "51 824 753 556", "category": "x",
             "gst_registered": "Y", "status": "Active", "payment_terms_days": 30, "bsb": "063-000",
             "account": "12345678"},
            {"vendor_id": "V2", "vendor_name": "Small Trader", "abn": "51 824 753 556", "category": "x",
             "gst_registered": "N", "status": "Active", "payment_terms_days": 14, "bsb": "063-001",
             "account": "87654321"},
        ]),
        "po_lines": pd.DataFrame([
            {"po_number": "PO1", "po_line": 10, "vendor_id": "V1", "item": "Widget", "qty_ordered": 10,
             "unit_price_ex_gst": 50.0, "gst_code": "GST", "po_status": "Open"},
            {"po_number": "PO1", "po_line": 20, "vendor_id": "V1", "item": "Flour", "qty_ordered": 5,
             "unit_price_ex_gst": 20.0, "gst_code": "FRE", "po_status": "Open"},
            {"po_number": "PO2", "po_line": 10, "vendor_id": "V2", "item": "Labour", "qty_ordered": 4,
             "unit_price_ex_gst": 100.0, "gst_code": "FRE", "po_status": "Open"},
        ]),
        "receipts": pd.DataFrame([
            {"po_number": "PO1", "po_line": 10, "qty_received": 10},
            {"po_number": "PO1", "po_line": 20, "qty_received": 5},
            {"po_number": "PO2", "po_line": 10, "qty_received": 4},
        ]),
        "invoices": pd.DataFrame([{
            "invoice_key": "B1", "vendor_id": "V1", "invoice_number": "INV-1043",
            "invoice_date": pd.Timestamp("2026-10-15"), "received_date": pd.Timestamp("2026-10-16"),
            "abn_on_invoice": "51824753556", "po_number": "PO1", "total_ex_gst": 600.0, "gst_total": 50.0,
            "total_inc_gst": 650.0, "bsb_on_invoice": "063-000", "account_on_invoice": "12345678"}]),
        "invoice_lines": pd.DataFrame([
            {"invoice_key": "B1", "line_no": 1, "po_line": 10, "item": "Widget", "qty": 10,
             "unit_price_ex_gst": 50.0, "line_ex_gst": 500.0, "line_gst": 50.0},
            {"invoice_key": "B1", "line_no": 2, "po_line": 20, "item": "Flour", "qty": 5,
             "unit_price_ex_gst": 20.0, "line_ex_gst": 100.0, "line_gst": 0.0},
        ]),
        "posted": pd.DataFrame([{"vendor_id": "V1", "invoice_number": "INV-0999", "total_inc_gst": 99.0,
                                 "posted_date": pd.Timestamp("2026-09-01")}]),
    }


def run(data):
    results, exc = validate.validate(data, SETTINGS)
    return results.iloc[0]["status"], exc


def codes(exc):
    return sorted(set(exc["code"]))


def set_invoice(data, **values):
    for k, v in values.items():
        data["invoices"].loc[0, k] = v


def set_line(data, i, **values):
    for k, v in values.items():
        data["invoice_lines"].loc[i, k] = v


@pytest.mark.parametrize("raw, expected", [
    ("INV-1043", "1043"), ("inv001043", "1043"), ("1043", "1043"), (" 10/43 ", "1043"), ("TI1043", "TI1043")])
def test_normalise_invoice_number(raw, expected):
    assert normalise_invoice_number(raw) == expected


def test_clean_invoice_passes():
    status, exc = run(clean_data())
    assert status == "PASS" and exc.empty


def test_v01_inactive_vendor():
    d = clean_data()
    d["vendors"].loc[0, "status"] = "Inactive"
    status, exc = run(d)
    assert status == "HOLD" and codes(exc) == ["V01"]


@pytest.mark.parametrize("abn, fragment", [("", "No ABN"), ("51 824 753 557", "checksum"),
                                           ("53 004 085 616", "does not match")])
def test_v02_abn(abn, fragment):
    d = clean_data()
    set_invoice(d, abn_on_invoice=abn)
    status, exc = run(d)
    assert codes(exc) == ["V02"] and fragment in exc.iloc[0]["detail"]


def test_v03_bank_details_changed_but_blank_is_fine():
    d = clean_data()
    set_invoice(d, account_on_invoice="99999999")
    assert codes(run(d)[1]) == ["V03"]
    d = clean_data()
    set_invoice(d, bsb_on_invoice="", account_on_invoice="")
    assert run(d)[0] == "PASS"


def test_v04_already_posted_with_reformatted_number():
    d = clean_data()
    d["posted"].loc[0, "invoice_number"] = "INV001043"
    status, exc = run(d)
    assert codes(exc) == ["V04"] and "Already posted on 01/09/2026" in exc.iloc[0]["detail"]


def test_v04_twice_in_batch_flags_second_only():
    d = clean_data()
    second = d["invoices"].iloc[0].copy()
    second["invoice_key"], second["invoice_number"] = "B2", "1043"
    d["invoices"] = pd.concat([d["invoices"], second.to_frame().T], ignore_index=True)
    lines = d["invoice_lines"].copy()
    lines["invoice_key"] = "B2"
    d["invoice_lines"] = pd.concat([d["invoice_lines"], lines], ignore_index=True)
    results, exc = validate.validate(d, SETTINGS)
    assert results.set_index("invoice_key")["status"].to_dict() == {"B1": "PASS", "B2": "HOLD"}
    assert exc.iloc[0]["detail"] == "Same invoice as B1 in this batch"


@pytest.mark.parametrize("po, fragment", [("", "No PO"), ("PO9", "not found"), ("PO2", "belongs to vendor V2")])
def test_v05_po_header(po, fragment):
    d = clean_data()
    set_invoice(d, po_number=po)
    status, exc = run(d)
    assert codes(exc) == ["V05"] and fragment in exc.iloc[0]["detail"]


def test_v05_closed_po_and_missing_line():
    d = clean_data()
    d["po_lines"].loc[d["po_lines"]["po_number"] == "PO1", "po_status"] = "Closed"
    assert "is closed" in run(d)[1].iloc[0]["detail"]
    d = clean_data()
    set_line(d, 1, po_line=30)
    exc = run(d)[1]
    assert codes(exc) == ["V05"] and exc.iloc[0]["line_no"] == 2


def test_v06_price_tolerance_boundary():
    d = clean_data()
    set_line(d, 0, unit_price_ex_gst=51.0, line_ex_gst=510.0, line_gst=51.0)   # exactly +2%: allowed
    set_invoice(d, total_ex_gst=610.0, gst_total=51.0, total_inc_gst=661.0)
    assert run(d)[0] == "PASS"
    set_line(d, 0, unit_price_ex_gst=51.5, line_ex_gst=515.0, line_gst=51.5)   # +3%: held
    set_invoice(d, total_ex_gst=615.0, gst_total=51.5, total_inc_gst=666.5)
    exc = run(d)[1]
    assert codes(exc) == ["V06"] and "3.0% above" in exc.iloc[0]["detail"]


def test_v07_billed_before_received():
    d = clean_data()
    d["receipts"].loc[0, "qty_received"] = 6
    exc = run(d)[1]
    assert codes(exc) == ["V07"] and exc.iloc[0]["detail"] == "Qty 10 invoiced but only 6 received"


def test_v08_line_and_header_arithmetic():
    d = clean_data()
    set_line(d, 0, line_ex_gst=510.0, line_gst=51.0)
    set_invoice(d, total_ex_gst=610.0, gst_total=51.0, total_inc_gst=661.0)
    exc = run(d)[1]
    assert codes(exc) == ["V08"] and len(exc) == 1
    d = clean_data()
    set_invoice(d, total_inc_gst=750.0)
    exc = run(d)[1]
    assert codes(exc) == ["V08"] and "does not equal ex-GST plus GST (650.00)" in exc.iloc[0]["detail"]


def test_v09_gst_rules():
    d = clean_data()
    set_line(d, 1, line_gst=10.0)                       # GST on a GST-free line
    set_invoice(d, gst_total=60.0, total_inc_gst=660.0)
    exc = run(d)[1]
    assert codes(exc) == ["V09"] and "GST-free" in exc.iloc[0]["detail"]
    d = clean_data()
    set_line(d, 0, line_gst=50.04)                      # within 5 cents: fine
    set_invoice(d, gst_total=50.04, total_inc_gst=650.04)
    assert run(d)[0] == "PASS"


def test_v09_unregistered_vendor_charging_gst():
    d = clean_data()
    set_invoice(d, vendor_id="V2", po_number="PO2", abn_on_invoice="51 824 753 556", bsb_on_invoice="",
                account_on_invoice="", total_ex_gst=400.0, gst_total=40.0, total_inc_gst=440.0)
    d["invoice_lines"] = pd.DataFrame([{"invoice_key": "B1", "line_no": 1, "po_line": 10, "item": "Labour",
                                        "qty": 4, "unit_price_ex_gst": 100.0, "line_ex_gst": 400.0,
                                        "line_gst": 40.0}])
    exc = run(d)[1]
    assert list(exc["code"]) == ["V09", "V09"]
    assert exc.iloc[1]["detail"] == "Vendor is not registered for GST but charged 40.00 GST"


def test_v10_future_is_hold_stale_is_review():
    d = clean_data()
    set_invoice(d, invoice_date=pd.Timestamp("2026-11-01"))
    assert run(d)[0] == "HOLD"
    set_invoice(d, invoice_date=pd.Timestamp("2025-09-01"))
    status, exc = run(d)
    assert status == "REVIEW" and exc.iloc[0]["action"] == "REVIEW"


def test_a01_approval_limit_is_review_only():
    d = clean_data()
    set_line(d, 0, qty=10, unit_price_ex_gst=50.0)
    d["po_lines"].loc[0, "unit_price_ex_gst"] = 1000.0
    set_line(d, 0, unit_price_ex_gst=1000.0, line_ex_gst=10000.0, line_gst=1000.0)
    set_invoice(d, total_ex_gst=10100.0, gst_total=1000.0, total_inc_gst=11100.0)
    status, exc = run(d)
    assert status == "REVIEW" and codes(exc) == ["A01"]


def test_sample_batch_catches_every_planted_issue_and_nothing_else():
    data, settings = workbook.read_inputs(cfg.TOOL_WORKBOOK)
    results, exc = validate.validate(data, settings)
    seeded = pd.read_csv(cfg.SAMPLE_DIR / "seeded_issues.csv")
    found = set(zip(exc["invoice_key"], exc["code"]))
    assert all((k, c) in found for k, c in zip(seeded["invoice_key"], seeded["code"]))
    planted = set(zip(seeded["invoice_key"], seeded["code"]))
    extra = [(k, c) for k, c in found if c != "A01" and (k, c) not in planted]
    assert extra == []


def test_workbook_settings_round_trip():
    _, settings = workbook.read_inputs(cfg.TOOL_WORKBOOK)
    assert settings["batch_date"] == cfg.SETTINGS["batch_date"]
    for key in ("price_tolerance", "approval_limit", "stale_days"):
        assert settings[key] == pytest.approx(cfg.SETTINGS[key])
