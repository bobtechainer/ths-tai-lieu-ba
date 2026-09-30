import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl

wb = openpyxl.load_workbook(os.path.join(os.path.dirname(__file__), 'Rà soát tính năng trùng lặp VNTIT.xlsx'))
print('ALL SHEETS:', wb.sheetnames)

# Focus on THCS,THPT sheet first
for name in wb.sheetnames:
    ws = wb[name]
    print(f'\n{"="*80}')
    print(f'SHEET: {name} (Rows: {ws.max_row}, Cols: {ws.max_column})')
    print(f'{"="*80}')
    
    # Print header row
    header = [c.value for c in list(ws.iter_rows(min_row=1, max_row=1))[0]]
    while header and header[-1] is None:
        header.pop()
    print(f'HEADER: {header}')
    print()
    
    # Print only non-empty rows, first 100 rows max
    count = 0
    for idx, row_data in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200), values_only=True), 1):
        if all(v is None for v in row_data):
            continue
        vals = list(row_data)
        while vals and vals[-1] is None:
            vals.pop()
        print(f'  R{idx}: {vals}')
        count += 1
    print(f'\n  Total non-empty rows: {count}')
