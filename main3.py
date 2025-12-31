import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from decimal import Decimal, InvalidOperation

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# --------------------
# Helpers & regex
# --------------------

def normalize_line(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\u00A0", " ")  # non-breaking space
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def clean_client_name(s: str) -> str:
    """
    Removes page markers that sometimes get captured as part of the client name,
    e.g. "7 of 16 Hulst ..." -> "Hulst ..."
    """
    s = normalize_line(s)
    s = re.sub(r"^\d+\s+of\s+\d+\s+", "", s, flags=re.IGNORECASE)
    return s.strip()


DATE_RE = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2}, \d{4}"
PERIOD_RE = re.compile(fr"({DATE_RE})\s*-\s*({DATE_RE})", re.IGNORECASE)

TOTAL_RE = re.compile(r"Total\s*-\s*EUR\s*€?\s*([-\d.,\(\)]+)", re.IGNORECASE)

# Domain-like: "alsemconsultancy.nl", "droitsdurgence.org", "hq-time.com"
DOMAIN_RE = re.compile(r"^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

# Money-like amounts require decimals (this prevents capturing years like 2025)
MONEY_TOKEN_RE = re.compile(r"-?\(?\d[\d.,]*[.,]\d{2}\)?")

# Total Current Charges Detail line: extract last money token from the line
TCCD_RE = re.compile(r"^Total Current Charges Detail", re.IGNORECASE)


def extract_period(text: str) -> str | None:
    m = PERIOD_RE.search(text)
    if m:
        return f"{m.group(1)} - {m.group(2)}"
    return None


def to_float(amount_str: str | None) -> float | None:
    """
    Convert strings like:
      "156.98", "(1.65)", "1,085.18", "1.085,18"
    into float.
    """
    if not amount_str:
        return None

    s = normalize_line(amount_str)
    s = s.replace("€", "").replace("EUR", "").replace(" ", "")
    s = s.replace("(", "-").replace(")", "")

    # Handle thousand/decimal separators
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # "1.085,18" -> "1085.18"
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # "1,085.18" -> "1085.18"
            s = s.replace(",", "")
    elif "," in s and "." not in s:
        # "156,98" -> "156.98"
        s = s.replace(",", ".")

    try:
        return float(Decimal(s))
    except (InvalidOperation, ValueError):
        return None


def extract_last_money_token(line: str) -> str | None:
    """
    Returns the last token that looks like a money amount with cents (2 decimals).
    This avoids false captures like the year "2025".
    """
    l = normalize_line(line)
    tokens = MONEY_TOKEN_RE.findall(l)
    if not tokens:
        return None
    return tokens[-1]


def extract_license_fields(line: str):
    """
    Robust extraction for license lines. Uses:
      - term: Monthly/Annual
      - qty: first integer token after term
      - subtotal: last MONEY token (must include decimals)
    """
    l = normalize_line(line)

    mterm = re.search(r"\b(Monthly|Annual)\b", l, flags=re.IGNORECASE)
    if not mterm:
        return None, None, None

    term = mterm.group(1).capitalize()

    # qty: first pure integer token after term
    qty = None
    term_pos = mterm.start()
    for mtok in re.finditer(r"-?\(?\d[\d.,]*\)?", l):
        if mtok.start() < term_pos:
            continue
        tok = mtok.group(0)
        raw = tok.replace("(", "").replace(")", "").replace("-", "")
        # integer only (no '.' or ',')
        if re.fullmatch(r"\d+", raw):
            try:
                qty = int(raw)
                break
            except ValueError:
                pass

    # subtotal: last money token (must have decimals)
    subtotal = extract_last_money_token(l)

    return term, qty, subtotal


def is_product_header_line(line: str) -> bool:
    """
    Real product header is:
      "<product name> - <domain>"
    where <domain> has no spaces and looks domain-like.
    """
    l = normalize_line(line)
    if " - " not in l:
        return False
    if "Total - EUR" in l or "Charges" in l:
        return False

    left, right = l.split(" - ", 1)
    right = right.strip()

    if " " in right:
        return False
    if not DOMAIN_RE.match(right):
        return False

    return bool(left.strip())


# --------------------
# Block parsing
# --------------------

def parse_block(client: str, default_period: str | None, block_lines: list[str]) -> dict:
    """
    Parse one product/adjustment block into a row.

    Key behavior:
      - For license blocks, extract term/qty/subtotal from license line.
      - For adjustments blocks, extract period + subtotal from a line that contains a date range
        and a money amount (decimals), never a year.
      - Total from "Total - EUR € ..."
    """
    if not block_lines:
        return {}

    description = normalize_line(block_lines[0])

    term = None
    qty = None
    subtotal = None

    term_text = None
    period_from_dates = None
    subtotal_from_dates = None

    # Heuristic: adjustment block if it contains "Adjustments" or date-subtotal header
    is_adjustment = any(
        normalize_line(x).lower() in ("adjustments", "description date subtotal")
        for x in block_lines
    )

    # First pass
    for raw in block_lines[1:]:
        line = normalize_line(raw)
        if not line:
            continue

        # License line
        if term is None or qty is None or subtotal is None:
            t, q, st = extract_license_fields(line)
            if t and term is None:
                term = t
            if q is not None and qty is None:
                qty = q
            if st is not None and subtotal is None:
                subtotal = st

        # Fallback term detection
        if term_text is None:
            if re.search(r"\bAnnual\b", line, re.IGNORECASE):
                term_text = "Annual"
            elif re.search(r"\bMonthly\b", line, re.IGNORECASE):
                term_text = "Monthly"

        # Adjustment/date line with amount:
        # only accept an amount token with decimals (prevents capturing 2025)
        if "Charges" not in line:
            p = extract_period(line)
            if p:
                period_from_dates = p
                money = extract_last_money_token(line)
                if money:
                    subtotal_from_dates = money

    period = period_from_dates if period_from_dates else default_period

    # Total
    total = None
    for raw in block_lines:
        line = normalize_line(raw)
        m = TOTAL_RE.search(line)
        if m:
            total = m.group(1)
            break

    # Prefer adjustment subtotal if no license subtotal OR if this looks like an adjustment block
    if (subtotal is None and subtotal_from_dates is not None) or (is_adjustment and subtotal_from_dates is not None):
        subtotal = subtotal_from_dates

    # For adjustments, often subtotal == total; if subtotal missing but total exists, copy it
    if is_adjustment and subtotal is None and total is not None:
        subtotal = total

    # Term fallback
    if term is None and term_text is not None:
        term = term_text

    # Qty cleanup: bogus 365 (from "Microsoft 365") should become empty
    if qty == 365:
        qty = None

    return {
        "Client": client,
        "Description": description,
        "Term": term,
        "Period": period,
        "Qty": qty,
        "Subtotal": to_float(subtotal),
        "Total": to_float(total),
        "IsAdjustment": bool(is_adjustment),
    }


# --------------------
# Invoice extraction
# --------------------

def extract_invoice_rows(pdf_path: str) -> tuple[list[dict], dict[str, float]]:
    """
    Returns:
      - rows: list of dict rows (one per block)
      - client_tccd: dict mapping client -> Total Current Charges Detail amount
    """
    rows: list[dict] = []
    client_tccd: dict[str, float] = {}

    current_client: str | None = None
    current_period: str | None = None

    context: str | None = None
    client_lines: list[str] = []
    block_lines: list[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw_line in text.split("\n"):
                line = normalize_line(raw_line)
                if not line:
                    continue

                # Ignore page markers like "7 of 16"
                if re.match(r"^\d+\s+of\s+\d+$", line, re.IGNORECASE):
                    continue

                # boilerplate
                if line.startswith("Phone:"):
                    continue
                if line.startswith("AppRiver, AG"):
                    continue
                if line.startswith("Invoice Summary") or line.startswith("Invoice #"):
                    continue
                if line.startswith("Invoice Date") or line.startswith("Total Due"):
                    continue
                if line.startswith("Subscription Terms"):
                    continue
                if line.startswith("Account Details") or line.startswith("Payment Information"):
                    continue

                # Capture Total Current Charges Detail per client (do NOT skip anymore)
                if TCCD_RE.match(line):
                    # extract last money token with decimals
                    amt_token = extract_last_money_token(line)
                    amt_val = to_float(amt_token) if amt_token else None
                    if amt_val is not None:
                        key = clean_client_name(current_client or "UNKNOWN")
                        client_tccd[key] = amt_val
                    continue

                # update current_period from Charges line
                if "Charges" in line and "-" in line:
                    p = extract_period(line)
                    if p:
                        current_period = p
                    continue

                # section headers
                if "Subscriptions of Utilis" in line:
                    if block_lines:
                        rows.append(parse_block(clean_client_name(current_client or "Utilis"), current_period, block_lines))
                        block_lines = []
                    context = "subscriptions"
                    current_client = "Utilis"
                    client_lines = []
                    continue

                if "Customers of Utilis" in line:
                    if block_lines:
                        rows.append(parse_block(clean_client_name(current_client or "UNKNOWN"), current_period, block_lines))
                        block_lines = []
                    context = "customers"
                    current_client = None
                    client_lines = []
                    continue

                # table headers
                if line.startswith("Description Term Qty UnitPrice Subtotal"):
                    continue
                if line.startswith("Description Date Subtotal"):
                    # keep a marker in the block when inside it; skip as a standalone line
                    if block_lines:
                        block_lines.append("Description Date Subtotal")
                    continue
                if line == "Adjustments":
                    if block_lines:
                        block_lines.append("Adjustments")
                    continue

                # product header detection
                is_header = is_product_header_line(line)

                # in customers context, collect client lines until first product header
                if context == "customers" and not block_lines and not is_header:
                    client_lines.append(line)
                    continue

                # start new block
                if is_header:
                    if block_lines:
                        rows.append(parse_block(clean_client_name(current_client or "UNKNOWN"), current_period, block_lines))
                        block_lines = []

                    if context == "customers":
                        if client_lines:
                            current_client = clean_client_name(" ".join(client_lines))
                            client_lines = []
                        elif current_client is None:
                            current_client = "UNKNOWN"
                    else:
                        current_client = "Utilis"

                    block_lines = [line]
                    continue

                # inside block
                if block_lines:
                    block_lines.append(line)
                    if "Total - EUR" in line:
                        rows.append(parse_block(clean_client_name(current_client or "UNKNOWN"), current_period, block_lines))
                        block_lines = []
                    continue

    if block_lines:
        rows.append(parse_block(clean_client_name(current_client or "UNKNOWN"), current_period, block_lines))

    return rows, client_tccd


# --------------------
# Excel output
# --------------------

HEADERS = [
    "Client",
    "Description",
    "Term",
    "Period",
    "Qty",
    "Subtotal",
    "Total",
    "Total Current Charges Detail",
]


def autosize_columns(ws):
    for col_idx, header in enumerate(HEADERS, start=1):
        max_len = len(header)
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            v = row[0].value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 80)


def save_to_excel(rows: list[dict], client_tccd: dict[str, float], output_path: str):
    # Excel Euro currency format (works across locales)
    EURO_FMT = u'[$€-x-euro2] #,##0.00'
    RED_FONT = Font(color="00FF0000")  # ARGB

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Write rows
    for r in rows:
        client = clean_client_name(r.get("Client") or "UNKNOWN")
        qty_val = r.get("Qty")
        qty_cell_value = "-" if qty_val is None else qty_val

        tccd_val = client_tccd.get(client)

        ws.append([
            client,
            r.get("Description"),
            r.get("Term"),
            r.get("Period"),
            qty_cell_value,
            r.get("Subtotal"),
            r.get("Total"),
            tccd_val,
        ])

        # If adjustments row, color the entire row red
        if r.get("IsAdjustment"):
            current_row = ws.max_row
            for col in range(1, len(HEADERS) + 1):
                ws.cell(row=current_row, column=col).font = RED_FONT

    # Formatting: Qty, currency columns
    for row in ws.iter_rows(min_row=2):
        # Qty column (E)
        if row[4].value not in (None, "-"):
            try:
                row[4].value = int(row[4].value)
            except Exception:
                pass

        # Currency columns: Subtotal (F), Total (G), TCCD (H)
        for idx in (5, 6, 7):
            cell = row[idx]
            if isinstance(cell.value, (int, float)):
                cell.number_format = EURO_FMT
                cell.alignment = Alignment(horizontal="right")

    autosize_columns(ws)
    wb.save(output_path)


# --------------------
# Conversion + GUI
# --------------------

def convert_pdf_to_excel(pdf_path: str) -> str:
    rows, client_tccd = extract_invoice_rows(pdf_path)
    output_path = os.path.splitext(pdf_path)[0] + ".xlsx"
    save_to_excel(rows, client_tccd, output_path)
    return output_path


def on_convert():
    pdf_path = filedialog.askopenfilename(
        title="Select AppRiver Invoice PDF",
        filetypes=[("PDF files", "*.pdf")],
    )
    if not pdf_path:
        return

    try:
        out = convert_pdf_to_excel(pdf_path)
        messagebox.showinfo("Done", f"Excel created:\n{out}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    root.title("PDF → Excel Converter")
    root.geometry("420x160")
    root.resizable(False, False)

    tk.Label(root, text="Convert an AppRiver invoice PDF to Excel", pady=18).pack()
    tk.Button(root, text="Select PDF and Convert", command=on_convert, width=28, height=2).pack()

    root.mainloop()


if __name__ == "__main__":
    main()
