"""Check the ACTUAL prognoz section generation."""
import sys
sys.path.insert(0, r'C:\Users\Admin\workspace')

# Import the script's functions
from email_analysis_automation import generate_report_for_date

# Simulate typical values
granula_cur = 980880
granula_prog = 1600383
polu_cur = 924181
polu_prog = 1507874
peregon = 0

# Reproduce lw/nw calc
lw = max(len('материал'), len('Гранула'), len('П/фабрикат'), len('Перегон'))
nw = max(len("текущее"), len("прогноз"),
         max(len(str(x)) for x in [granula_cur, granula_prog, polu_cur, polu_prog, peregon]))
pw = nw
print(f"lw={lw}, pw={pw}")

hdr = '<code>' + 'материал'.ljust(lw) + 'текущее'.rjust(pw) + ' ' + 'прогноз'.rjust(pw) + '</code>'
print(f"HEADER: {repr(hdr)}")
print(f"LEN: {len(hdr)}")

# Count inside <code>
inside = 'материал'.ljust(lw) + 'текущее'.rjust(pw) + ' ' + 'прогноз'.rjust(pw)
print(f"INSIDE: {repr(inside)}")
print(f"INSIDE LEN: {len(inside)}")

# Check saved file
import glob
reports = sorted(glob.glob(r'C:\Users\Admin\workspace\reports\report_*.txt'))
print(f"\nLast file: {reports[-1]}")
with open(reports[-1], 'r', encoding='utf-8') as f:
    for line in f:
        if 'материал' in line and 'текущее' in line and 'прогноз' in line:
            start = line.index('>') + 1
            end = line.rindex('<')
            inside_saved = line[start:end]
            print(f"SAVED INSIDE: {repr(inside_saved)}")
            print(f"SAVED INSIDE LEN: {len(inside_saved)}")
