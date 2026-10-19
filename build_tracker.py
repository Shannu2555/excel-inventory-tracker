"""Builds inventory_tracker.xlsx (Products, Stock In, Sales, Dashboard, How-To)."""
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference

NP, NT = 200, 2000          # product rows, transaction rows
BLUE, LIGHT = "1F4E78", "DDEBF7"
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor=BLUE)
INPUT = PatternFill("solid", fgColor="FFF2CC")   # yellow = you type here
thin = Side(style="thin", color="BFBFBF")

def header(ws, cols, widths):
    ws.append(cols)
    for c, w in zip(ws[1], widths):
        c.font, c.fill = H, HF
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"

wb = Workbook()
how = wb.active; how.title = "How-To"
prod = wb.create_sheet("Products"); sin = wb.create_sheet("Stock In")
sales = wb.create_sheet("Sales"); dash = wb.create_sheet("Dashboard")

# ---------- Products ----------
header(prod, ["SKU", "Product Name", "Category", "Unit Cost", "Selling Price", "Opening Stock",
              "Reorder Level", "Received", "Incoming (on order)", "Sold", "Current Stock",
              "Stock Value", "Status", "Units Sold (30d)", "Days of Stock Left"],
       [12, 28, 14, 11, 12, 12, 12, 11, 13, 9, 13, 13, 13, 13, 13])
import json; _R=json.load(open("data/real_data.json"))
samples = [tuple(x) for x in _R["products"]]
_old = [("SKU-001", "Wireless Earbuds", "Electronics", 8.5, 24.99, 120, 30),
           ("SKU-002", "Phone Case - Black", "Accessories", 1.2, 9.99, 300, 60),
           ("SKU-003", "USB-C Cable 1m", "Accessories", 0.9, 7.49, 80, 50),
           ("SKU-004", "Laptop Stand", "Office", 6.0, 29.99, 24, 15),
           ("SKU-005", "LED Desk Lamp", "Office", 7.5, 34.99, 22, 10),
           ("SKU-006", "Yoga Mat", "Fitness", 4.0, 19.99, 60, 20),
           ("SKU-007", "Water Bottle 750ml", "Fitness", 2.1, 12.99, 45, 25),
           ("SKU-008", "Mini Bluetooth Speaker", "Electronics", 9.0, 39.99, 35, 10)]
for r in range(2, NP + 2):
    if r - 2 < len(samples):
        for col, v in zip("ABCDEFG", samples[r - 2]):
            prod[f"{col}{r}"] = v
    for col in "ABCDEFG":
        prod[f"{col}{r}"].fill = INPUT
    S = "'Stock In'"
    prod[f"H{r}"] = f'=IF($A{r}="","",SUMIFS({S}!$C:$C,{S}!$B:$B,$A{r},{S}!$E:$E,"Received"))'
    prod[f"I{r}"] = f'=IF($A{r}="","",SUMIFS({S}!$C:$C,{S}!$B:$B,$A{r},{S}!$E:$E,"Incoming"))'
    prod[f"J{r}"] = f'=IF($A{r}="","",SUMIFS(Sales!$C:$C,Sales!$B:$B,$A{r}))'
    prod[f"K{r}"] = f'=IF($A{r}="","",MAX(0,N($F{r})+N($H{r})-N($J{r})))'
    prod[f"L{r}"] = f'=IF($A{r}="","",$K{r}*N($D{r}))'
    prod[f"M{r}"] = (f'=IF($A{r}="","",IF($K{r}<=0,"OUT OF STOCK",'
                     f'IF($K{r}<=N($G{r}),IF(N($I{r})>0,"LOW - ON ORDER","REORDER NOW"),"OK")))')
    prod[f"N{r}"] = (f'=IF($A{r}="","",SUMIFS(Sales!$C:$C,Sales!$B:$B,$A{r},'
                     f'Sales!$A:$A,">="&(TODAY()-30)))')
    prod[f"O{r}"] = f'=IF($A{r}="","",IF(N($N{r})=0,"no recent sales",MAX(0,ROUND($K{r}/($N{r}/30),0))))'
    for col in "DEL":
        prod[f"{col}{r}"].number_format = '#,##0.00'
rng = f"A2:O{NP+1}"
prod.conditional_formatting.add(rng, FormulaRule(formula=['AND($A2<>"",$K2<=0)'],
                                fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(bold=True, color="9C0006")))
prod.conditional_formatting.add(rng, FormulaRule(formula=['AND($A2<>"",$K2<=$G2)'],
                                fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C5700")))
prod.auto_filter.ref = f"A1:O{NP+1}"

# SKU dropdown for transaction sheets
dv_sku = lambda: DataValidation(type="list", formula1=f"=Products!$A$2:$A${NP+1}", allow_blank=True,
                                showErrorMessage=True, errorTitle="Unknown SKU",
                                error="Pick a SKU that exists on the Products sheet.")

# ---------- Stock In ----------
header(sin, ["Date", "SKU", "Quantity", "Supplier", "Status", "PO / Reference", "Product (auto)"],
       [12, 12, 10, 20, 12, 16, 28])
today = date.today()
stock_rows = []
_old2 = [(today - timedelta(days=20), "SKU-001", 50, "ShenZhen Audio", "Received", "PO-1001"),
              (today - timedelta(days=15), "SKU-003", 100, "CableCo", "Received", "PO-1002"),
              (today - timedelta(days=5), "SKU-007", 60, "BottleWorks", "Incoming", "PO-1003"),
              (today - timedelta(days=3), "SKU-002", 200, "CaseMart", "Received", "PO-1004")]
for i, row in enumerate(stock_rows, 2):
    for col, v in zip("ABCDEF", row):
        sin[f"{col}{i}"] = v
for r in range(2, NT + 2):
    sin[f"A{r}"].number_format = "dd-mmm-yyyy"
    sin[f"G{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(Products!$B:$B,MATCH(B{r},Products!$A:$A,0)),"SKU not found"))'
    for col in "ABCDEF":
        sin[f"{col}{r}"].fill = INPUT
d = dv_sku(); sin.add_data_validation(d); d.add(f"B2:B{NT+1}")
st = DataValidation(type="list", formula1='"Received,Incoming"', allow_blank=True)
sin.add_data_validation(st); st.add(f"E2:E{NT+1}")

# ---------- Sales ----------
header(sales, ["Date", "SKU", "Quantity", "Channel", "Order ID", "Unit Price (auto)", "Revenue (auto)",
               "Product (auto)"], [12, 12, 10, 14, 16, 14, 14, 28])
sale_rows = [(date.fromisoformat(a),b_,c,d_,e) for a,b_,c,d_,e in _R["sales"]]
sale_rows.sort()
for i, row in enumerate(sale_rows, 2):
    for col, v in zip("ABCDE", row):
        sales[f"{col}{i}"] = v
for r in range(2, NT + 2):
    sales[f"A{r}"].number_format = "dd-mmm-yyyy"
    sales[f"F{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(Products!$E:$E,MATCH(B{r},Products!$A:$A,0)),0))'
    sales[f"G{r}"] = f'=IF(B{r}="","",N(C{r})*N(F{r}))'
    sales[f"H{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(Products!$B:$B,MATCH(B{r},Products!$A:$A,0)),"SKU not found"))'
    for col in "FG":
        sales[f"{col}{r}"].number_format = '#,##0.00'
    for col in "ABCDE":
        sales[f"{col}{r}"].fill = INPUT
d = dv_sku(); sales.add_data_validation(d); d.add(f"B2:B{NT+1}")
ch = DataValidation(type="list", formula1='"UK,Ireland,France,Germany,Other EU/Intl"', allow_blank=True)
sales.add_data_validation(ch); ch.add(f"D2:D{NT+1}")

# ---------- Dashboard ----------
dash.column_dimensions["A"].width = 30; dash.column_dimensions["B"].width = 16
for c in "DEFGH": dash.column_dimensions[c].width = 16
dash["A1"] = "Inventory Dashboard"; dash["A1"].font = Font(bold=True, size=16, color=BLUE)
dash["A2"] = '="Updated: "&TEXT(TODAY(),"dd-mmm-yyyy")'
P = f"Products!$A$2:$A${NP+1}"
kpis = [("Active SKUs", f'=COUNTIF({P},"?*")', "0"),
        ("Total units in stock", f"=SUM(Products!$K$2:$K${NP+1})", "#,##0"),
        ("Stock value (at cost)", f"=SUM(Products!$L$2:$L${NP+1})", "#,##0.00"),
        ("Units on order (incoming)", f"=SUM(Products!$I$2:$I${NP+1})", "#,##0"),
        ("SKUs to reorder now", f'=COUNTIF(Products!$M$2:$M${NP+1},"REORDER NOW")', "0"),
        ("SKUs low but on order", f'=COUNTIF(Products!$M$2:$M${NP+1},"LOW - ON ORDER")', "0"),
        ("SKUs out of stock", f'=COUNTIF(Products!$M$2:$M${NP+1},"OUT OF STOCK")', "0"),
        ("Units sold (all time)", f"=SUM(Sales!$C$2:$C${NT+1})", "#,##0"),
        ("Revenue (all time)", f"=SUM(Sales!$G$2:$G${NT+1})", "#,##0.00"),
        ("Revenue (last 30 days)", f'=SUMIFS(Sales!$G$2:$G${NT+1},Sales!$A$2:$A${NT+1},">="&(TODAY()-30))', "#,##0.00")]
for i, (lab, f, fmt) in enumerate(kpis, 4):
    dash[f"A{i}"] = lab; dash[f"B{i}"] = f; dash[f"B{i}"].number_format = fmt
    dash[f"A{i}"].fill = PatternFill("solid", fgColor=LIGHT); dash[f"B{i}"].font = Font(bold=True)
    for c in (dash[f"A{i}"], dash[f"B{i}"]): c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
for r in (8, 10):
    dash.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"],
                                    fill=PatternFill("solid", fgColor="F8CBAD")))

# Sales by channel
dash["A16"] = "Sales by market (country)"; dash["A16"].font = Font(bold=True, size=12, color=BLUE)
dash["A17"], dash["B17"] = "Market", "Revenue"
for c in (dash["A17"], dash["B17"]): c.font, c.fill = H, HF
for i, cname in enumerate(["UK", "Ireland", "France", "Germany", "Other EU/Intl"], 18):
    dash[f"A{i}"] = cname
    dash[f"B{i}"] = f'=SUMIFS(Sales!$G$2:$G${NT+1},Sales!$D$2:$D${NT+1},A{i})'
    dash[f"B{i}"].number_format = "#,##0.00"

# Per-SKU stock table (mirrors Products, top 30 rows) + reorder highlight
dash["D3"] = "Stock by SKU"; dash["D3"].font = Font(bold=True, size=12, color=BLUE)
for col, t in zip("DEFGH", ["SKU", "Product", "Current Stock", "Reorder Level", "Status"]):
    dash[f"{col}4"] = t; dash[f"{col}4"].font, dash[f"{col}4"].fill = H, HF
dash.column_dimensions["E"].width = 26
for k in range(30):
    r, pr = 5 + k, 2 + k
    dash[f"D{r}"] = f'=IF(Products!A{pr}="","",Products!A{pr})'
    dash[f"E{r}"] = f'=IF(Products!A{pr}="","",Products!B{pr})'
    dash[f"F{r}"] = f'=IF(Products!A{pr}="","",Products!K{pr})'
    dash[f"G{r}"] = f'=IF(Products!A{pr}="","",Products!G{pr})'
    dash[f"H{r}"] = f'=IF(Products!A{pr}="","",Products!M{pr})'
dash.conditional_formatting.add("D5:H34", FormulaRule(formula=['$H5="OUT OF STOCK"'],
                                fill=PatternFill("solid", fgColor="F8CBAD")))
dash.conditional_formatting.add("D5:H34", FormulaRule(formula=['OR($H5="REORDER NOW",$H5="LOW - ON ORDER")'],
                                fill=PatternFill("solid", fgColor="FFEB9C")))
chart = BarChart(); chart.title = "Current stock vs reorder level"; chart.height, chart.width = 8, 18
chart.add_data(Reference(dash, min_col=6, max_col=7, min_row=4, max_row=4 + len(samples)), titles_from_data=True)
chart.set_categories(Reference(dash, min_col=4, min_row=5, max_row=4 + len(samples)))
dash.add_chart(chart, "J4")

# ---------- How-To ----------
how.column_dimensions["A"].width = 110
lines = [("How to use this Inventory Tracker", True),
         ("Yellow cells are for you to type in. White cells are formulas, so don't type over them.", False),
         ("", False),
         ("1. PRODUCTS: add one row per SKU: SKU, name, category, unit cost, selling price, opening stock and reorder level.", False),
         ("   Received, Incoming, Sold, Current Stock, Stock Value, Status and Days of Stock Left calculate automatically.", False),
         ("   Rows turn YELLOW when stock is at or below the reorder level, and RED when stock is 0 or below.", False),
         ("2. STOCK IN: log every purchase order or delivery. Pick the SKU from the dropdown.", False),
         ("   Status = 'Incoming' while it's on the way (counted as Incoming); change it to 'Received' when it arrives (added to stock).", False),
         ("3. SALES: log sales (or paste a daily export from Amazon / Shopify): Date, SKU, Quantity, Channel, Order ID.", False),
         ("   Unit price and revenue are looked up from Products automatically. 'SKU not found' means a typo or a missing product.", False),
         ("4. DASHBOARD: totals, items to reorder, out-of-stock count, revenue by channel and a stock chart. Nothing to edit.", False),
         ("", False),
         ("Status meanings", True),
         ("   OK = above reorder level | REORDER NOW = at/below reorder level, nothing on order | "
          "LOW - ON ORDER = low but a PO is incoming | OUT OF STOCK = 0 or less", False),
         ("Days of Stock Left = current stock / average daily sales over the last 30 days.", False),
         ("", False),
         ("Tips", True),
         ("   Never delete formula columns. To remove a product, clear its yellow cells and the formulas go blank by themselves.", False),
         (f"   Capacity: {NP} products and {NT} rows each for Stock In and Sales. Copy a formula row down to extend it.", False),
         ("   Use the filter arrows on Products to sort by Status or Days of Stock Left.", False),
         ("Data: UCI Machine Learning Repository, Online Retail (Daqing Chen, 2015), CC BY 4.0, doi:10.24432/C5BW33. 15 most-ordered products; Sales = real units sold 1 Nov-9 Dec 2011, grouped per day and country.", False),
         ("Assumptions (dataset has no stock or cost data): opening stock on 1 Nov 2011 = 60 days of avg daily demand (Dec 2010-Oct 2011); reorder level = 14 days; unit cost = 60% of avg selling price; stock cannot go below 0.", False)]
for i, (t, bold) in enumerate(lines, 1):
    how[f"A{i}"] = t
    if bold: how[f"A{i}"].font = Font(bold=True, size=13 if i == 1 else 11, color=BLUE)

wb.save("inventory_tracker.xlsx")
print("saved inventory_tracker.xlsx")
