"""Debug the actual generate_report_for_date call."""
import sys
sys.path.insert(0, r'C:\Users\Admin\workspace')
import json
import tempfile
from pathlib import Path
from datetime import datetime

# Import what we need
from email_analysis_automation import load_config, get_latest_excel, extract_forecast, get_target_dates, excel_serial_to_date
from email_analysis_automation import generate_report_for_date

config = load_config()

# Find the latest Excel
email_info, content, filename = get_latest_excel(config)
if not content:
    print("NO FILE FOUND")
    sys.exit(1)

print(f"File: {filename}")

# Parse it
import pandas as pd
suffix = Path(filename).suffix.lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(content)
    tmp_path = tmp.name

sheets_data = {}
# Use openpyxl
xls = pd.ExcelFile(tmp_path, engine="openpyxl")
for sheet_name in xls.sheet_names:
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, header=None)
        sheets_data[sheet_name] = df.fillna("").values.tolist()
    except:
        pass

forecast_data = extract_forecast(sheets_data)
granula_cur, granula_prog, polu_cur, polu_prog, peregon = forecast_data
print(f"granula_cur={granula_cur}, granula_prog={granula_prog}")
print(f"polu_cur={polu_cur}, polu_prog={polu_prog}")
print(f"peregon={peregon}")

# Now replicate lw/nw calc
lw = max(len('материал'), len('Гранула'), len('П/фабрикат'), len('Перегон'))
nw = max(len("текущее"), len("прогноз"),
         max(len(str(x)) for x in [granula_cur, granula_prog, polu_cur, polu_prog, peregon]))
pw = nw
print(f"lw={lw} ('материал'={len('материал')}, 'Гранула'={len('Гранула')}, 'П/фабрикат'={len('П/фабрикат')}, 'Перегон'={len('Перегон')})")
print(f"nw={nw}, pw={pw}")
print(f"Number lengths: {[len(str(x)) for x in [granula_cur, granula_prog, polu_cur, polu_prog, peregon]]}")

target_dates = get_target_dates(sheets_data)
report = generate_report_for_date(sheets_data, target_dates[-1], forecast_data)
# Find the prognoz lines
for line in report.split('\n'):
    if 'материал' in line and 'текущее' in line:
        print(f"\nLINE: {repr(line)}")
        print(f"LEN: {len(line)}")
        start = line.index('>') + 1
        end = line.rindex('<')
        inside = line[start:end]
        print(f"INSIDE: {repr(inside)}")
        print(f"INSIDE LEN: {len(inside)}")
        break
