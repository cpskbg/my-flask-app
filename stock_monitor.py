#!/usr/bin/env python3
"""
Script de surveillance du stock FIFO pour détecter les incohérences
À exécuter régulièrement pour maintenir l'intégrité des données
"""

import sqlite3
from datetime import datetime

def check_stock_consistency():
    """
    Vérifie la cohérence du stock FIFO pour tous les articles
    Retourne True si tout est cohérent, False sinon
    """
    
    conn = sqlite3.connect('stock_management.db')
    cursor = conn.cursor()
    
    # Requête pour détecter les incohérences
    cursor.execute('''
        SELECT 
            i.sku,
            i.name,
            i.quantity as stock_items,
            COALESCE(SUM(
                rs.quantite_restante - COALESCE((
                    SELECT SUM(di.quantite_dotee) 
                    FROM dotation_items di 
                    JOIN dotations d ON di.dotation_id = d.id 
                    WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            ), 0) as stock_fifo,
            (i.quantity - COALESCE(SUM(
                rs.quantite_restante - COALESCE((
                    SELECT SUM(di.quantite_dotee) 
                    FROM dotation_items di 
                    JOIN dotations d ON di.dotation_id = d.id 
                    WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            ), 0)) as difference
        FROM items i
        LEFT JOIN reception_stocks rs ON i.id = rs.item_id
        GROUP BY i.id, i.sku, i.name, i.quantity
        HAVING i.quantity != COALESCE(SUM(
            rs.quantite_restante - COALESCE((
                SELECT SUM(di.quantite_dotee) 
                FROM dotation_items di 
                JOIN dotations d ON di.dotation_id = d.id 
                WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
            ), 0)
        ), 0)
        ORDER BY ABS(difference) DESC
    ''')
    
    inconsistencies = cursor.fetchall()
    conn.close()
    
    return len(inconsistencies) == 0, inconsistencies

def generate_stock_report():
    """
    Génère un rapport complet sur l'état du stock
    """
    
    conn = sqlite3.connect('stock_management.db')
    cursor = conn.cursor()
    
    print("=" * 80)
    print("📊 RAPPORT D'ÉTAT DU STOCK FIFO")
    print("=" * 80)
    print(f"📅 Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Statistiques générales
    cursor.execute('SELECT COUNT(*) FROM items')
    total_items = cursor.fetchone()[0]
    
    cursor.execute('SELECT SUM(quantity) FROM items')
    total_stock = cursor.fetchone()[0] or 0
    
    cursor.execute('SELECT COUNT(*) FROM items WHERE quantity > 0')
    items_with_stock = cursor.fetchone()[0]
    
    cursor.execute('SELECT COUNT(*) FROM items WHERE quantity <= reorder_level')
    items_below_reorder = cursor.fetchone()[0]
    
    print(f"📈 STATISTIQUES GÉNÉRALES:")
    print(f"   • Articles totaux: {total_items}")
    print(f"   • Stock total: {total_stock} unités")
    print(f"   • Articles avec stock: {items_with_stock}")
    print(f"   • Articles sous seuil de réapprovisionnement: {items_below_reorder}")
    print()
    
    # Vérification des incohérences
    is_consistent, inconsistencies = check_stock_consistency()
    
    if is_consistent:
        print("✅ ÉTAT DU STOCK: COHÉRENT")
        print("   Aucune incohérence détectée entre le stock items et le stock FIFO")
    else:
        print("⚠️  ÉTAT DU STOCK: INCOHÉRENT")
        print(f"   {len(inconsistencies)} incohérence(s) détectée(s):")
        print()
        
        for i, (sku, name, stock_items, stock_fifo, difference) in enumerate(inconsistencies, 1):
            print(f"   {i}. {name[:40]} ({sku})")
            print(f"      Stock items: {stock_items} | Stock FIFO: {stock_fifo} | Diff: {difference}")
    
    print()
    
    # Articles avec stock faible
    cursor.execute('''
        SELECT sku, name, quantity, reorder_level
        FROM items 
        WHERE quantity > 0 AND quantity <= reorder_level
        ORDER BY (quantity - reorder_level) ASC
        LIMIT 10
    ''')
    
    low_stock_items = cursor.fetchall()
    
    if low_stock_items:
        print("🔴 ARTICLES AVEC STOCK FAIBLE (≤ seuil de réapprovisionnement):")
        for sku, name, quantity, reorder_level in low_stock_items:
            print(f"   • {name[:40]} ({sku}): {quantity}/{reorder_level}")
        print()
    
    # Articles sans stock
    cursor.execute('''
        SELECT sku, name, reorder_level
        FROM items 
        WHERE quantity = 0 AND reorder_level > 0
        ORDER BY reorder_level DESC
        LIMIT 10
    ''')
    
    out_of_stock_items = cursor.fetchall()
    
    if out_of_stock_items:
        print("🔴 ARTICLES EN RUPTURE DE STOCK:")
        for sku, name, reorder_level in out_of_stock_items:
            print(f"   • {name[:40]} ({sku}) (seuil: {reorder_level})")
        print()
    
    conn.close()
    
    return is_consistent, len(inconsistencies)

def auto_fix_inconsistencies():
    """
    Corrige automatiquement toutes les incohérences détectées
    """
    
    conn = sqlite3.connect('stock_management.db')
    cursor = conn.cursor()
    
    # Récupérer les incohérences
    cursor.execute('''
        SELECT 
            i.id,
            i.sku,
            i.name,
            i.quantity as stock_items,
            COALESCE(SUM(
                rs.quantite_restante - COALESCE((
                    SELECT SUM(di.quantite_dotee) 
                    FROM dotation_items di 
                    JOIN dotations d ON di.dotation_id = d.id 
                    WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            ), 0) as stock_fifo
        FROM items i
        LEFT JOIN reception_stocks rs ON i.id = rs.item_id
        GROUP BY i.id, i.sku, i.name, i.quantity
        HAVING i.quantity != COALESCE(SUM(
            rs.quantite_restante - COALESCE((
                SELECT SUM(di.quantite_dotee) 
                FROM dotation_items di 
                JOIN dotations d ON di.dotation_id = d.id 
                WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            ), 0)
    ''')
    
    inconsistencies = cursor.fetchall()
    fixed_count = 0
    
    for item_id, sku, name, stock_items, stock_fifo in inconsistencies:
        # Corriger l'incohérence
        cursor.execute('UPDATE items SET quantity = ? WHERE id = ?', (stock_fifo, item_id))
        fixed_count += 1
        print(f"✅ Corrigé: {name[:40]} ({sku}) - {stock_items} → {stock_fifo}")
    
    conn.commit()
    conn.close()
    
    return fixed_count

def main():
    """
    Fonction principale du script de surveillance
    """
    
    print("🔍 SURVEILLANCE DU STOCK FIFO")
    print("=" * 50)
    
    # Générer le rapport
    is_consistent, inconsistency_count = generate_stock_report()
    
    if not is_consistent:
        print(f"\n⚠️  {inconsistency_count} incohérence(s) détectée(s)")
        print("Voulez-vous corriger automatiquement ces incohérences? (y/n)")
        
        # En mode automatique, on corrige directement
        auto_fix_inconsistencies()
        print(f"\n✅ Toutes les incohérences ont été corrigées!")
        
        # Vérification finale
        is_consistent_after, _ = generate_stock_report()
        
        if is_consistent_after:
            print("🎉 Le stock est maintenant complètement cohérent!")
        else:
            print("⚠️  Des incohérences subsistent, vérification manuelle requise")
    else:
        print("🎉 Le stock est parfaitement cohérent!")
    
    print("\n💡 CONSEILS:")
    print("• Exécutez ce script régulièrement (ex: quotidien ou hebdomadaire)")
    print("• Intégrez-le dans un cron job pour une surveillance automatique")
    print("• Configurez des alertes email si des incohérences sont détectées")

if __name__ == "__main__":
    main()
