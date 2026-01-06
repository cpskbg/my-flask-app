import sqlite3
import os

db_path = 'stock_management.db'
if not os.path.exists(db_path):
    print(f"Base de données introuvable à {db_path}")
else:
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    print("=== Table 'users' ===")
    cursor.execute("PRAGMA table_info(users)")
    for col in cursor.fetchall():
        print(f"  {col[1]} ({col[2]})")
    
    print("\n=== Contenu de 'users' ===")
    cursor.execute("SELECT id, email, name, role, is_super_admin, is_active, is_deleted FROM users")
    for row in cursor.fetchall():
        print(row)
    conn.close()
