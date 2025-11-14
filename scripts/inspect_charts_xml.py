import os, glob, zipfile

temp_dir = os.environ.get('TEMP') or os.environ.get('TMP') or r'C:\Windows\Temp'
pattern = os.path.join(temp_dir, 'Motor_Performance_Report_*.xlsx')
files = glob.glob(pattern)
if not files:
    print('No report files found')
    raise SystemExit(1)
latest = max(files, key=os.path.getmtime)
print('Latest report:', latest)

with zipfile.ZipFile(latest, 'r') as z:
    chart_files = [f for f in z.namelist() if f.startswith('xl/charts/') and f.endswith('.xml')]
    if not chart_files:
        print('No chart xml files found')
    for cf in chart_files:
        print('\n---', cf, '---')
        data = z.read(cf).decode('utf-8', errors='replace')
        # print a short snippet containing sheet names or sheet references
        lines = data.splitlines()
        for i, line in enumerate(lines):
            if 'sheet' in line.lower() or 'sheetname' in line.lower() or 'v=' in line.lower() or 'ref' in line.lower() or 'F!$' in line:
                snippet = line.strip()
                print(i+1, snippet)
        # Print a small excerpt:
        excerpt = '\n'.join(lines[:120])
        print('\nExcerpt start:\n', excerpt[:2000])
print('\nDone')
