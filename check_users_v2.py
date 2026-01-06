import sqlite3
import os

db_path = 'stock_management.db'
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT id, email, role, is_super_admin FROM users")
    for row in cursor.fetchall():
        print(f"ID: {row[0]} | Email: {row[1]} | Role: {row[2]} | SuperAdmin: {row[3]}")
    conn.close()
