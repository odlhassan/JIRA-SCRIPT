import sqlite3
conn = sqlite3.connect(r'E:\JIRA SCRIPT\assignee_hours_capacity.db')
rows = conn.execute('SELECT DISTINCT issue_type, COUNT(*) as cnt FROM canonical_issues GROUP BY issue_type ORDER BY cnt DESC').fetchall()
for r in rows:
    print(r)
conn.close()
