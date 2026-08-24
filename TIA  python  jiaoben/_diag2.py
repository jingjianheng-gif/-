import sys, sqlite3
sys.stdout.reconfigure(encoding='utf-8')

VCI = r'C:\Users\Administrator\Documents\Project\JSB-25-081B（瑞源橡塑）\JSB-25--081B(瑞源橡塑）TPV包纱管1.0\Vci\Vci.db'
conn = sqlite3.connect(VCI)
cur = conn.cursor()
tables = cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
print('Tables:', [t[0] for t in tables])
for t in tables:
    tn = t[0]
    cols = [c[1] for c in cur.execute(f'PRAGMA table_info({tn})').fetchall()]
    cnt = cur.execute(f'SELECT COUNT(*) FROM {tn}').fetchone()[0]
    print(f'\n{tn}: cols={cols} rows={cnt}')
    rows = cur.execute(f'SELECT * FROM {tn} LIMIT 5').fetchall()
    for r in rows:
        vals = [str(v)[:60] if isinstance(v, bytes) else str(v)[:40] for v in r]
        print(f'  {vals}')
conn.close()
