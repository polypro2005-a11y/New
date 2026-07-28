import sqlite3, os
db = os.path.expanduser('~/AppData/Local/hermes/state.db')
conn = sqlite3.connect(db)
cur = conn.cursor()
cur.execute("SELECT id, message_count, created_at FROM sessions WHERE ended_at IS NULL ORDER BY message_count DESC LIMIT 5")
for r in cur.fetchall():
    sid, cnt, created = r
    print(f'Session: {sid[:24]}... | msgs: {cnt} | created: {created}')
conn.close()
