"""Check EXACT content of saved report files."""
with open(r'C:\Users\Admin\workspace\reports\report_20260717_200234.txt', 'r', encoding='utf-8') as f:
    text = f.read()

# Find prognoz section
idx = text.find('📊')
if idx > 0:
    section = text[idx:]
    for line in section.split('\n')[:6]:
        print(repr(line))
        print(f'  len={len(line)}')
        # Extract inside code tags
        if '<code>' in line:
            start = line.index('>') + 1
            end = line.rindex('<')
            inside = line[start:end]
            print(f'  inside code: {repr(inside)} len={len(inside)}')
            # Count spaces
            space_count = inside.count(' ')
            print(f'  spaces: {space_count}')
