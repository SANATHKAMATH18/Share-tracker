"""
Capital Gains Classifier Script
================================
Reads the original CapitalGain CSV file and generates a classified Excel report
with three sections:
  1. Short Term  - transactions with non-zero ShortTermProfit
  2. Long Term   - transactions with non-zero LongTermProfit
  3. Squaring    - transactions with non-zero SquringProfit (intraday)

The output matches the format of the reference file in the 'modified' folder.

Usage:
    python classify_capital_gains.py [input_csv] [output_xlsx]

Defaults:
    input_csv  = orginal/CapitalGain_20242025.csv
    output_xlsx = output/CapitalGain_Classified.xlsx
"""

import csv
import os
import sys
from collections import defaultdict
from datetime import datetime
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def parse_date(date_str):
    if not date_str:
        return None
    date_str = str(date_str).strip()
    if not date_str or date_str.lower() in ('nan', 'none', ''):
        return None
    # Try explicit formats
    for fmt in (
        '%d-%b-%y', '%d-%b-%Y', '%d %b %Y', '%d %b %y', 
        '%Y-%m-%d', '%d-%m-%Y', '%d-%m-%y', '%d/%m/%Y', '%d/%m/%y', 
        '%b %Y', '%d %B %Y',
        '%Y-%m-%d %H:%M:%S', '%d-%b-%y %H:%M:%S', '%d-%b-%Y %H:%M:%S',
        '%d/%m/%Y %H:%M:%S', '%d/%m/%y %H:%M:%S', '%d-%m-%Y %H:%M:%S',
        '%d-%m-%y %H:%M:%S'
    ):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass
    # Fallback parsing logic
    parts = re.split(r'[-/\s]+', date_str)
    year = None
    for p in parts:
        if p.isdigit() and len(p) == 4:
            year = int(p)
            break
    if year:
        months = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                  'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
        month = 6
        for p in parts:
            p_low = p.lower()[:3]
            if p_low in months:
                month = months[p_low]
                break
            elif p.isdigit() and 1 <= int(p) <= 12:
                month = int(p)
        try:
            return datetime(year, month, 1)
        except Exception:
            pass
    return None


def read_and_classify(csv_path):
    """Read the CSV or Excel and aggregate data into Short Term, Long Term, and Squaring buckets."""

    short_term = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})
    long_term = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})
    squaring = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})

    # Extract client code and name from the first data row
    cl_code = None
    cl_name = None
    fy_counts = defaultdict(int)

    # Check file type and read accordingly
    if csv_path.lower().endswith('.xlsx'):
        import pandas as pd
        import io
        df = pd.read_excel(csv_path)
        csv_file = io.StringIO(df.to_csv(index=False))
        reader = csv.DictReader(csv_file)
        f = None
    else:
        f = open(csv_path, 'r', encoding='utf-8-sig')
        reader = csv.DictReader(f)

    try:
        for row in reader:
            # Capture client info from first row
            if cl_code is None:
                cl_code = row.get('ClCode', '').strip()
                cl_name = row.get('ClName', '').strip()

            name = row['scrip_name'].strip()
            stp = float(row['ShortTermProfit'])
            ltp = float(row['LongTermProfit'])
            sqp = float(row['SquringProfit'])

            buy_qty = float(row['BuyQty'])
            buy_val = float(row['BuyValue'])
            sell_qty = float(row['sellQty'])
            sell_val = float(row['SellValue'])

            # Determine transaction date and track FY
            dt = parse_date(row.get('SellDate')) or parse_date(row.get('BuyDate'))
            if dt:
                fy = f"{dt.year}-{str(dt.year+1)[-2:]}" if dt.month >= 4 else f"{dt.year-1}-{str(dt.year)[-2:]}"
                fy_counts[fy] += 1

            if stp != 0:
                short_term[name]['BuyQty'] += buy_qty
                short_term[name]['BuyValue'] += buy_val
                short_term[name]['SellQty'] += sell_qty
                short_term[name]['SellValue'] += sell_val
                short_term[name]['Profit'] += stp

            if ltp != 0:
                long_term[name]['BuyQty'] += buy_qty
                long_term[name]['BuyValue'] += buy_val
                long_term[name]['SellQty'] += sell_qty
                long_term[name]['SellValue'] += sell_val
                long_term[name]['Profit'] += ltp

            if sqp != 0:
                squaring[name]['BuyQty'] += buy_qty
                squaring[name]['BuyValue'] += buy_val
                squaring[name]['SellQty'] += sell_qty
                squaring[name]['SellValue'] += sell_val
                squaring[name]['Profit'] += sqp
    finally:
        if f is not None:
            f.close()

    # Derive financial year from transaction dates if possible
    if fy_counts:
        fy_year = max(fy_counts, key=fy_counts.get)
    else:
        # Fallback to CSV filename
        basename = os.path.basename(csv_path)
        if '20242025' in basename:
            fy_year = '2024-25'
        elif '20232024' in basename:
            fy_year = '2023-24'
        else:
            fy_year = '2024-25'  # default

    return short_term, long_term, squaring, cl_code, cl_name, fy_year


def compute_grand_total(data, profit_key='Profit'):
    """Compute grand total across all scrips."""
    total = {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0}
    for vals in data.values():
        total['BuyQty'] += vals['BuyQty']
        total['BuyValue'] += vals['BuyValue']
        total['SellQty'] += vals['SellQty']
        total['SellValue'] += vals['SellValue']
        total['Profit'] += vals['Profit']
    return total


def write_excel(short_term, long_term, squaring, cl_code, cl_name, fy_year, output_path):
    """Write the classified data to an Excel file matching the reference format."""

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # ── Styles ──
    header_font = Font(name='Calibri', bold=True, size=11)
    section_font = Font(name='Calibri', bold=True, size=12, color='1F4E79')
    title_font = Font(name='Calibri', bold=True, size=11, color='1F4E79')
    data_font = Font(name='Calibri', size=11)
    total_font = Font(name='Calibri', bold=True, size=11)
    number_format = '#,##0.00'
    qty_format = '#,##0'

    header_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    total_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ── Column widths ──
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 22

    current_row = 1

    # ── Row 1: blank ──
    current_row += 1

    # ── Row 2: Client Code ──
    ws.cell(row=current_row, column=2, value=cl_code).font = title_font
    current_row += 1

    # ── Row 3: Client Name + FY Year ──
    ws.cell(row=current_row, column=2, value=cl_name).font = title_font
    ws.cell(row=current_row, column=7, value=fy_year).font = title_font
    current_row += 1

    def write_section(section_name, data, profit_label, start_row):
        """Write a section (Short Term / Long Term / Squaring) and return the next row."""
        row = start_row

        # ── Section Title ──
        ws.cell(row=row, column=2, value=section_name).font = section_font
        row += 1

        # ── Column Headers ──
        headers = ['Scrip Name', 'Sum of BuyQty', 'Sum of BuyValue',
                    'Sum of sellQty', 'Sum of SellValue', profit_label]
        for col_idx, header in enumerate(headers, start=2):
            cell = ws.cell(row=row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        row += 1

        # ── Data Rows ──
        for name in sorted(data.keys()):
            vals = data[name]
            ws.cell(row=row, column=2, value=name).font = data_font
            ws.cell(row=row, column=2).border = thin_border

            qty_cell = ws.cell(row=row, column=3, value=round(vals['BuyQty']))
            qty_cell.font = data_font
            qty_cell.number_format = qty_format
            qty_cell.border = thin_border

            val_cell = ws.cell(row=row, column=4, value=round(vals['BuyValue'], 2))
            val_cell.font = data_font
            val_cell.number_format = number_format
            val_cell.border = thin_border

            sqty_cell = ws.cell(row=row, column=5, value=round(vals['SellQty']))
            sqty_cell.font = data_font
            sqty_cell.number_format = qty_format
            sqty_cell.border = thin_border

            sval_cell = ws.cell(row=row, column=6, value=round(vals['SellValue'], 2))
            sval_cell.font = data_font
            sval_cell.number_format = number_format
            sval_cell.border = thin_border

            profit_cell = ws.cell(row=row, column=7, value=round(vals['Profit'], 2))
            profit_cell.font = data_font
            profit_cell.number_format = number_format
            profit_cell.border = thin_border
            # Color code: green for profit, red for loss
            if vals['Profit'] >= 0:
                profit_cell.font = Font(name='Calibri', size=11, color='006100')
            else:
                profit_cell.font = Font(name='Calibri', size=11, color='9C0006')

            row += 1

        # ── Grand Total Row ──
        totals = compute_grand_total(data)
        ws.cell(row=row, column=2, value='Grand Total').font = total_font
        ws.cell(row=row, column=2).fill = total_fill
        ws.cell(row=row, column=2).border = thin_border

        gt_bqty = ws.cell(row=row, column=3, value=round(totals['BuyQty']))
        gt_bqty.font = total_font
        gt_bqty.number_format = qty_format
        gt_bqty.fill = total_fill
        gt_bqty.border = thin_border

        gt_bval = ws.cell(row=row, column=4, value=round(totals['BuyValue'], 2))
        gt_bval.font = total_font
        gt_bval.number_format = number_format
        gt_bval.fill = total_fill
        gt_bval.border = thin_border

        gt_sqty = ws.cell(row=row, column=5, value=round(totals['SellQty']))
        gt_sqty.font = total_font
        gt_sqty.number_format = qty_format
        gt_sqty.fill = total_fill
        gt_sqty.border = thin_border

        gt_sval = ws.cell(row=row, column=6, value=round(totals['SellValue'], 2))
        gt_sval.font = total_font
        gt_sval.number_format = number_format
        gt_sval.fill = total_fill
        gt_sval.border = thin_border

        gt_profit = ws.cell(row=row, column=7, value=round(totals['Profit'], 2))
        gt_profit.font = total_font
        gt_profit.number_format = number_format
        gt_profit.fill = total_fill
        gt_profit.border = thin_border

        row += 1

        # ── Blank separator row ──
        row += 1

        return row

    # ── Write Short Term Section ──
    current_row = write_section('Short term', short_term, 'Sum of ShortTermProfit', current_row)

    # ── Write Long Term Section ──
    current_row = write_section('Long Term', long_term, 'Sum of LongTermProfit', current_row)

    # ── Write Squaring Section ──
    current_row = write_section('Squaring ', squaring, 'Sum of SquringProfit', current_row)

    # ── Save ──
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"\n[OK] Report saved to: {output_path}")
    print(f"  Client: {cl_code} - {cl_name}")
    print(f"  FY: {fy_year}")
    print(f"  Short Term entries: {len(short_term)}")
    print(f"  Long Term entries:  {len(long_term)}")
    print(f"  Squaring entries:   {len(squaring)}")

    # Print summary
    st_total = compute_grand_total(short_term)
    lt_total = compute_grand_total(long_term)
    sq_total = compute_grand_total(squaring)

    print(f"\n  Short Term Profit/Loss: {st_total['Profit']:>15,.2f}")
    print(f"  Long Term Profit/Loss:  {lt_total['Profit']:>15,.2f}")
    print(f"  Squaring Profit/Loss:   {sq_total['Profit']:>15,.2f}")
    print(f"  ---------------------------------------")
    print(f"  Total P&L:              {st_total['Profit'] + lt_total['Profit'] + sq_total['Profit']:>15,.2f}")


def main():
    # Default paths
    input_csv = 'orginal/CapitalGain_20242025.csv'

    # Allow override of input CSV via command line
    if len(sys.argv) >= 2:
        input_csv = sys.argv[1]

    if not os.path.exists(input_csv):
        print(f"Error: Input file not found: {input_csv}")
        sys.exit(1)

    print(f"Reading: {input_csv}")
    short_term, long_term, squaring, cl_code, cl_name, fy_year = read_and_classify(input_csv)

    # Determine output XLSX path (dynamic or command-line override)
    if len(sys.argv) >= 3:
        output_xlsx = sys.argv[2]
    else:
        output_xlsx = f'output/CapitalGain_{fy_year}_Classified.xlsx'

    write_excel(short_term, long_term, squaring, cl_code, cl_name, fy_year, output_xlsx)


if __name__ == '__main__':
    main()
