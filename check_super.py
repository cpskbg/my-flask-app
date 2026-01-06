import sqlite3
import os

db_path = 'stock_management.db'
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT email, is_super_admin FROM users WHERE is_super_admin = 1")
    rows = cursor.fetchall()
    if not rows:
        print("AUCUN SUPER ADMIN TROUVE")
    for row in rows:
        print(f"Super Admin: {row[0]}")
    conn.close()
