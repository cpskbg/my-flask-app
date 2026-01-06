import sqlite3

conn = sqlite3.connect('stock_management.db')
cursor = conn.cursor()

print("=== Structure des tables ===")

# Vérifier la table items
print("\nTable 'items':")
cursor.execute('PRAGMA table_info(items)')
columns = cursor.fetchall()
for col in columns:
    print(f"  {col[1]} {col[2]} {'NOT NULL' if not col[3] else 'NULL'} {'PRIMARY KEY' if col[5] == 1 else ''}")

# Vérifier la table reception_stocks
print("\nTable 'reception_stocks':")
cursor.execute('PRAGMA table_info(reception_stocks)')
columns = cursor.fetchall()
for col in columns:
    print(f"  {col[1]} {col[2]} {'NOT NULL' if not col[3] else 'NULL'} {'PRIMARY KEY' if col[5] == 1 else ''}")

# Vérifier la table dotation_items
print("\nTable 'dotation_items':")
cursor.execute('PRAGMA table_info(dotation_items)')
columns = cursor.fetchall()
for col in columns:
    print(f"  {col[1]} {col[2]} {'NOT NULL' if not col[3] else 'NULL'} {'PRIMARY KEY' if col[5] == 1 else ''}")

conn.close()
