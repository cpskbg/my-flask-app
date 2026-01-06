import sqlite3

conn = sqlite3.connect('stock_management.db')
cursor = conn.cursor()

# Rechercher l'article
cursor.execute('SELECT name, sku, quantity FROM items WHERE name LIKE "%CLAVIER%" OR name LIKE "%SOURIS%" OR name LIKE "%LOGITECH%" OR name LIKE "%MK120%"')
results = cursor.fetchall()

print("Articles correspondants trouvés:")
for row in results:
    print(f'Nom: {row[0]}, SKU: {row[1]}, Quantité: {row[2]}')

# Si aucun résultat, afficher tous les articles pour diagnostic
if not results:
    print("\nAucun article trouvé. Voici tous les articles:")
    cursor.execute('SELECT name, sku, quantity FROM items ORDER BY name LIMIT 20')
    all_items = cursor.fetchall()
    for row in all_items:
        print(f'Nom: {row[0]}, SKU: {row[1]}, Quantité: {row[2]}')

conn.close()
