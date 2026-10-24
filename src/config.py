"""Paths and default validation settings.

The same settings are written to the Settings sheet of the Excel tool, and both the
Python validator and the VBA macro read them from there, so the two always agree.
"""
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DIR = ROOT / "data" / "sample_batch"
EXCEL_DIR = ROOT / "excel"
OUTPUT_DIR = ROOT / "outputs"
IMAGES_DIR = ROOT / "images"
TOOL_WORKBOOK = EXCEL_DIR / "Invoice_Validation_Tool.xlsx"

RANDOM_SEED = 11

SETTINGS = {
    "batch_date": date(2026, 10, 23),   # the day the batch is validated
    "price_tolerance": 0.02,            # invoice unit price may exceed PO price by 2%
    "line_tolerance": 0.01,             # qty x price vs line total, dollars
    "total_tolerance": 0.02,            # line totals vs header totals, dollars
    "gst_tolerance": 0.05,              # GST vs 10% of the ex-GST amount, dollars
    "gst_rate": 0.10,
    "approval_limit": 10_000,           # inc GST; at or above needs a second approver
    "stale_days": 365,                  # invoices older than this are queried
}

SETTING_NOTES = {
    "batch_date": "Date the batch is validated (future-dated / stale checks)",
    "price_tolerance": "Unit price may exceed the PO price by this share",
    "line_tolerance": "Allowed difference between qty x price and the line total ($)",
    "total_tolerance": "Allowed difference between line totals and invoice totals ($)",
    "gst_tolerance": "Allowed difference between GST and 10% of the ex-GST amount ($)",
    "gst_rate": "GST rate",
    "approval_limit": "Invoices at or above this (inc GST) need a second approver",
    "stale_days": "Invoices older than this many days are queried",
}

INPUT_SHEETS = {
    "vendors": "Vendors",
    "po_lines": "PO_Lines",
    "receipts": "Receipts",
    "invoices": "Invoices",
    "invoice_lines": "Invoice_Lines",
    "posted": "Posted_History",
}

# Exception codes in the order they are reported
CODES = {
    "V01": "Vendor not set up or inactive",
    "V02": "ABN missing, invalid or different from vendor master",
    "V03": "Bank details differ from vendor master",
    "V04": "Duplicate invoice",
    "V05": "PO missing, not found, closed, other vendor or bad line",
    "V06": "Unit price above PO price",
    "V07": "Quantity invoiced above quantity received",
    "V08": "Line or invoice totals do not add up",
    "V09": "GST incorrect",
    "V10": "Invoice date in the future or older than 12 months",
    "A01": "Needs second approval",
}
HOLD_CODES = {"V01", "V02", "V03", "V04", "V05", "V06", "V07", "V08", "V09"}
