import sqlite3

conn = sqlite3.connect('stock_management.db')
cursor = conn.cursor()

# Corriger l'incohérence - mettre le stock à 0 pour correspondre au FIFO
cursor.execute('UPDATE items SET quantity = 0 WHERE sku = "SKU-007"')
conn.commit()

print('✅ Stock corrigé pour CLAVIER+SOURIS OPTIQUE LOGITECH FILAIRE MK120 AZERTY')
print('   Stock mis à jour: 0 unités (cohérent avec FIFO)')
print('   Pour livrer 2 unités, vous devez d\'abord ajouter du stock via réception')

# Vérifier la correction
cursor.execute('SELECT name, sku, quantity FROM items WHERE sku = "SKU-007"')
item = cursor.fetchone()
if item:
    print(f'   Vérification: {item[0]} - Stock actuel: {item[2]} unités')

conn.close()
