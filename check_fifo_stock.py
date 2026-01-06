import sqlite3

conn = sqlite3.connect('stock_management.db')
cursor = conn.cursor()

# Vérifier le stock FIFO pour l'article SKU-007
print("=== État du stock FIFO pour CLAVIER+SOURIS OPTIQUE LOGITECH FILAIRE MK120 AZERTY ===")
print()

# Informations de l'article
cursor.execute('SELECT name, sku, quantity FROM items WHERE sku = "SKU-007"')
item = cursor.fetchone()
if item:
    print(f"Article: {item[0]}")
    print(f"SKU: {item[1]}")
    print(f"Stock total actuel: {item[2]} unités")
    print()

# Vérifier les réceptions FIFO disponibles
cursor.execute('''
    SELECT id, quantity, date_reception, prix_unitaire_ht, prix_unitaire_ttc, 
           (quantity - COALESCE((
               SELECT SUM(d.quantite_dotee) 
               FROM dotation_items di 
               JOIN dotations d ON di.dotation_id = d.id 
               WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
           ), 0)) as quantite_disponible
    FROM reception_stocks rs 
    WHERE item_id = (SELECT id FROM items WHERE sku = "SKU-007")
    ORDER BY date_reception ASC
''')

receptions = cursor.fetchall()
print("Réceptions FIFO disponibles:")
for i, rec in enumerate(receptions, 1):
    print(f"{i}. Réception #{rec[0]} - Date: {rec[2]}")
    print(f"   Quantité reçue: {rec[1]} unités")
    print(f"   Quantité disponible: {rec[5]} unités")
    print(f"   Prix HT: {rec[3]:.2f} MAD")
    print(f"   Prix TTC: {rec[4]:.2f} MAD")
    print()

# Calculer le total disponible
total_disponible = sum(rec[5] for rec in receptions)
print(f"Total disponible selon FIFO: {total_disponible} unités")
print(f"Stock total dans items: {item[2] if item else 0} unités")

if total_disponible != (item[2] if item else 0):
    print("⚠️  INCOHÉRENCE DÉTECTÉE entre le stock total et le stock FIFO!")
else:
    print("✅ Stock cohérent")

conn.close()
