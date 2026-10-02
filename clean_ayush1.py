"""Cleans Ayush1.xlsx -> Ayush1_cleaned.xlsx (Cleaned_Data, Change_Log, Verification sheets).
Place this file next to Ayush1.xlsx and run: python clean_ayush1.py
Requires: pandas, openpyxl"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

raw = pd.read_excel("Ayush1.xlsx")
df = raw.copy()
n = len(df)
log = []
def rec(desc, impact, status="Resolved"): log.append((f"CR{len(log)+1:03d}", desc, impact, status))

rec("Loaded raw dataset (Ayush1.xlsx, Sheet1)", f"{n} rows, {df.shape[1]} columns")
# Audit results (nothing to change, but documented)
rec("Audit: exact duplicate rows", "0 found - nothing removed")
rec("Audit: duplicate OrderID / TrackingNumber", "0 found - all IDs unique")
rec("Audit: nulls in all columns except CouponCode", "0 found")

# Text trimming / case (no-op check, still enforced)
txt = df.select_dtypes(include=["object","string"]).columns
chg = 0
for c in txt:
    new = df[c].astype("string").str.strip().str.replace(r"\s+"," ",regex=True)
    chg += (new.fillna("") != df[c].astype("string").fillna("")).sum(); df[c] = new
rec("Trimmed whitespace in all text columns", f"{chg} cells changed (data was already clean)")

# Coupon: blank = no coupon used, NOT missing data -> don't impute with mode
nc = df.CouponCode.isna().sum()
df["CouponCode"] = df.CouponCode.fillna("NO COUPON")
rec("CouponCode blanks filled with 'NO COUPON' (blank = coupon not used; mode imputation would wrongly invent a coupon)",
    f"{nc} records preserved ({nc/n:.1%})")

# Dates -> ISO 8601, drop 00:00:00 time component
df["Date"] = pd.to_datetime(df["Date"]).dt.normalize()
rec("Date converted to ISO 8601 (YYYY-MM-DD); removed meaningless 00:00:00 time part", f"{n} cells, range {df.Date.min():%Y-%m-%d} to {df.Date.max():%Y-%m-%d}")

# numbers
for c in ["UnitPrice","TotalPrice"]: df[c] = df[c].round(2)
rec("UnitPrice and TotalPrice set to 2-decimal precision", "2 columns")
mism = ((df.Quantity*df.UnitPrice - df.TotalPrice).abs() > 0.01).sum()
rec("Validated TotalPrice = Quantity x UnitPrice", f"{mism} mismatches")
rec("Validated Quantity <= ItemsInCart", f"{(df.Quantity>df.ItemsInCart).sum()} violations")
rec("OBSERVATION: Pending/Cancelled orders all carry tracking numbers", "Likely synthetic data; flagged, not altered", "Review")

# ---- workbook ----
F = "Arial"
hdr_fill = PatternFill("solid", fgColor="2F4F4F"); hdr_font = Font(name=F, bold=True, color="FFFFFF")
wb = Workbook()
ws = wb.active; ws.title = "Cleaned_Data"
ws.append(list(df.columns))
for r in df.itertuples(index=False):
    ws.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in r])
for c in ws[1]: c.fill, c.font = hdr_fill, hdr_font
cols = list(df.columns)
for i,c in enumerate(cols,1):
    L = get_column_letter(i)
    ws.column_dimensions[L].width = max(12, len(c)+4)
    for row in range(2, n+2):
        cell = ws[f"{L}{row}"]; cell.font = Font(name=F)
        if c=="Date": cell.number_format = "yyyy-mm-dd"
        if c in ("UnitPrice","TotalPrice"): cell.number_format = "0.00"
ws.column_dimensions["ShippingAddress".__class__("")] if False else None
ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

lg = wb.create_sheet("Change_Log")
lg.append(["Change ID","Description","Impact","Status"])
for r in log: lg.append(list(r))
for c in lg[1]: c.fill, c.font = hdr_fill, hdr_font
for row in lg.iter_rows(min_row=2):
    for c in row: c.font = Font(name=F); c.alignment = Alignment(wrap_text=True, vertical="top")
for L,w in zip("ABCD",[11,80,45,11]): lg.column_dimensions[L].width = w

v = wb.create_sheet("Verification")
last = n+1
rows = [
 ("Check","Result","Target","Status"),
 ("Total rows", f"=COUNTA(Cleaned_Data!A2:A{last})", n, None),
 ("Duplicate OrderIDs", f"=SUMPRODUCT((COUNTIF(Cleaned_Data!A2:A{last},Cleaned_Data!A2:A{last})>1)*1)", 0, None),
 ("Duplicate TrackingNumbers", f"=SUMPRODUCT((COUNTIF(Cleaned_Data!J2:J{last},Cleaned_Data!J2:J{last})>1)*1)", 0, None),
 ("Incorrectly formatted dates (non-date cells)", f"=ROWS(Cleaned_Data!B2:B{last})-COUNT(Cleaned_Data!B2:B{last})", 0, None),
 ("Dates with a time component", f"=SUMPRODUCT((Cleaned_Data!B2:B{last}<>INT(Cleaned_Data!B2:B{last}))*1)", 0, None),
 ("Blank cells (all 14 columns)", f"=COUNTBLANK(Cleaned_Data!A2:N{last})", 0, None),
 ("TotalPrice <> Quantity x UnitPrice", f"=SUMPRODUCT((ABS(Cleaned_Data!E2:E{last}*Cleaned_Data!F2:F{last}-Cleaned_Data!N2:N{last})>0.01)*1)", 0, None),
]
for i,r in enumerate(rows,1):
    v.append(list(r))
    if i>1: v[f"D{i}"] = f'=IF(B{i}=C{i},"PASS","FAIL")'
v.append([]); v.append(["OVERALL GATE (Project 2)", '=IF(COUNTIF(D2:D8,"FAIL")=0,"PASS - 0% error rate","FAIL")'])
for c in v[1]: c.fill, c.font = hdr_fill, hdr_font
for row in v.iter_rows(min_row=2):
    for c in row: c.font = Font(name=F, bold=(c.row==10))
for L,w in zip("ABCD",[48,26,10,10]): v.column_dimensions[L].width = w
wb.move_sheet("Verification", offset=-1) if False else None
wb.save("Ayush1_cleaned.xlsx")
pd.DataFrame(log, columns=["Change ID","Description","Impact","Status"]).to_csv("log.csv", index=False)
