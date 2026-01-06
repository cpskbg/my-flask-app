import sqlite3

conn = sqlite3.connect('stock_management.db')
cursor = conn.cursor()

# Informations de l'article
cursor.execute('SELECT id, name, sku, quantity FROM items WHERE sku = "SKU-007"')
item = cursor.fetchone()

if not item:
    print("❌ Article SKU-007 non trouvé!")
    conn.close()
    exit()

print("=== DIAGNOSTIC DU STOCK FIFO ===")
print()
print(f"Article: {item[1]}")
print(f"SKU: {item[2]}")
print(f"Stock total actuel: {item[3]} unités")
print()

# Vérifier les réceptions FIFO disponibles
cursor.execute('''
    SELECT rs.id, rs.quantite_initiale, rs.quantite_restante, rs.date_reception, 
           rs.prix_unitaire_ht, rs.prix_unitaire_ttc,
           (rs.quantite_restante - COALESCE((
               SELECT SUM(di.quantite_dotee) 
               FROM dotation_items di 
               JOIN dotations d ON di.dotation_id = d.id 
               WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
           ), 0)) as quantite_disponible
    FROM reception_stocks rs 
    WHERE rs.item_id = ?
    ORDER BY rs.date_reception ASC
''', (item[0],))

receptions = cursor.fetchall()
print("Réceptions FIFO disponibles:")
total_disponible = 0

for i, rec in enumerate(receptions, 1):
    print(f"{i}. Réception #{rec[0]} - Date: {rec[3]}")
    print(f"   Quantité initiale: {rec[1]} unités")
    print(f"   Quantité restante: {rec[2]} unités")
    print(f"   Quantité disponible: {rec[6]} unités")
    print(f"   Prix HT: {rec[4]:.2f} MAD")
    print(f"   Prix TTC: {rec[5]:.2f} MAD")
    total_disponible += rec[6]
    print()

print(f"Total disponible selon FIFO: {total_disponible} unités")
print(f"Stock total dans items: {item[3]} unités")

# Vérifier les dotations en cours qui utilisent cet article
cursor.execute('''
    SELECT d.id, d.numero_dotation, d.statut, di.quantite_dotee
    FROM dotations d
    JOIN dotation_items di ON d.id = di.dotation_id
    WHERE di.item_id = ? AND d.statut != 'livree'
    ORDER BY d.date_dotation DESC
''', (item[0],))

dotations_en_cours = cursor.fetchall()
if dotations_en_cours:
    print(f"\nDotations en cours utilisant cet article:")
    for dotation in dotations_en_cours:
        print(f"  • {dotation[1]} - Statut: {dotation[2]} - Quantité: {dotation[3]} unités")

# Vérifier la cohérence
if total_disponible != item[3]:
    print(f"\n⚠️  INCOHÉRENCE DÉTECTÉE!")
    print(f"   Stock items: {item[3]} unités")
    print(f"   Stock FIFO: {total_disponible} unités")
    print(f"   Différence: {item[3] - total_disponible} unités")
else:
    print(f"\n✅ Stock cohérent")

# Solution recommandée
print(f"\n=== SOLUTION RECOMMANDÉE ===")
if total_disponible < 2:
    print("❌ Stock insuffisant pour livrer 2 unités")
    print("Options:")
    print("1. Ajouter du stock via réception")
    print("2. Réduire la quantité dans la dotation à 1 unité")
    print("3. Mettre à jour manuellement le stock:")
    print(f"   UPDATE items SET quantity = 2 WHERE sku = 'SKU-007';")
else:
    print("✅ Stock suffisant disponible")

conn.close()
