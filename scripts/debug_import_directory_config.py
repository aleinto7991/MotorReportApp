import traceback
try:
    import src.directory_config as dc
    print('Imported src.directory_config OK')
    print('Members:', [m for m in dir(dc) if not m.startswith('_')][:50])
except Exception:
    traceback.print_exc()
