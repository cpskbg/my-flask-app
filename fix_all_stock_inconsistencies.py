#!/usr/bin/env python3
"""
Script de diagnostic et correction des incohérences de stock FIFO pour tous les articles
Auteur: Système de Gestion Stock
Date: 2025-11-23
"""

import sqlite3

def diagnose_and_fix_all_stock():
    """
    Diagnostic et correction des incohérences de stock pour tous les articles
    """
    
    conn = sqlite3.connect('stock_management.db')
    cursor = conn.cursor()
    
    print("=" * 80)
    print("🔍 DIAGNOSTIC COMPLET DES INCOHÉRENCES DE STOCK FIFO")
    print("=" * 80)
    print(f"📅 Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Récupérer tous les articles
    cursor.execute('''
        SELECT i.id, i.sku, i.name, i.quantity as stock_items
        FROM items i
        ORDER BY i.name
    ''')
    
    all_items = cursor.fetchall()
    total_articles = len(all_items)
    inconsistencies_found = 0
    inconsistencies_fixed = 0
    
    print(f"📊 Analyse de {total_articles} articles...")
    print()
    
    for i, (item_id, sku, name, stock_items) in enumerate(all_items, 1):
        print(f"[{i:3d}/{total_articles}] {name[:50]} ({sku})")
        
        # Calculer le stock FIFO disponible
        cursor.execute('''
            SELECT COALESCE(SUM(
                rs.quantite_restante - COALESCE((
                    SELECT SUM(di.quantite_dotee) 
                    FROM dotation_items di 
                    JOIN dotations d ON di.dotation_id = d.id 
                    WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            ), 0) as stock_fifo
            FROM reception_stocks rs 
            WHERE rs.item_id = ?
        ''', (item_id,))
        
        result = cursor.fetchone()
        stock_fifo = result[0] if result else 0
        
        # Comparer les stocks
        if stock_items != stock_fifo:
            inconsistencies_found += 1
            print(f"    ❌ INCOHÉRENCE: Items={stock_items} vs FIFO={stock_fifo}")
            
            # Correction automatique
            cursor.execute('UPDATE items SET quantity = ? WHERE id = ?', (stock_fifo, item_id))
            inconsistencies_fixed += 1
            print(f"    ✅ CORRIGÉ: Stock mis à jour à {stock_fifo} unités")
        else:
            print(f"    ✅ OK: Stock cohérent ({stock_items} unités)")
        
        print()
    
    # Validation des modifications
    conn.commit()
    
    # Afficher le résumé
    print("=" * 80)
    print("📋 RÉSUMÉ DU DIAGNOSTIC")
    print("=" * 80)
    print(f"📊 Articles analysés: {total_articles}")
    print(f"❌ Incohérences trouvées: {inconsistencies_found}")
    print(f"✅ Incohérences corrigées: {inconsistencies_fixed}")
    print(f"📈 Articles restants cohérents: {total_articles - inconsistencies_found}")
    
    if inconsistencies_found > 0:
        print(f"\n🎉 Toutes les incohérences ont été corrigées automatiquement!")
        print(f"   Le stock de chaque article correspond maintenant au calcul FIFO.")
    else:
        print(f"\n🎉 Parfait! Tous les articles ont déjà des stocks cohérents.")
    
    # Vérification finale
    print(f"\n🔍 VÉRIFICATION FINALE...")
    cursor.execute('''
        SELECT COUNT(*) 
        FROM items i 
        WHERE i.quantity != COALESCE((
            SELECT SUM(
                rs.quantite_restante - COALESCE((
                    SELECT SUM(di.quantite_dotee) 
                    FROM dotation_items di 
                    JOIN dotations d ON di.dotation_id = d.id 
                    WHERE di.reception_stock_id = rs.id AND d.statut = 'livree'
                ), 0)
            )
            FROM reception_stocks rs 
            WHERE rs.item_id = i.id
        ), 0)
    ''')
    
    remaining_inconsistencies = cursor.fetchone()[0]
    
    if remaining_inconsistencies == 0:
        print(f"✅ Succès: Plus aucune incohérence détectée!")
    else:
        print(f"⚠️  Attention: {remaining_inconsistencies} incohérence(s) restante(s)")
    
    conn.close()
    
    return inconsistencies_found, inconsistencies_fixed

def create_stock_monitoring_query():
    """
    Crée une requête SQL pour surveiller les incohérences futures
    """
    
    query = """
    -- Requête pour détecter les incohérences de stock FIFO
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
    ORDER BY difference DESC;
    """
    
    return query

def main():
    """
    Fonction principale
    """
    inconsistencies_found, inconsistencies_fixed = diagnose_and_fix_all_stock()
    
    print(f"\n" + "=" * 80)
    print("📄 REQUÊTE DE SURVEILLANCE POUR L'AVENIR")
    print("=" * 80)
    print("Pour surveiller les incohérences futures, utilisez cette requête SQL:")
    print()
    print(create_stock_monitoring_query())
    print()
    print("💡 CONSEILS POUR ÉVITER LES INCOHÉRENCES FUTURES:")
    print("1. Validez toujours les réceptions avant de livrer des dotations")
    print("2. Utilisez le système FIFO pour toutes les sorties de stock")
    print("3. Vérifiez régulièrement le stock avec ce script")
    print("4. Évitez les modifications manuelles directes de la table items")
    print("5. Assurez-vous que les dotations livrées sont correctement comptabilisées")
    
    return inconsistencies_found, inconsistencies_fixed

if __name__ == "__main__":
    from datetime import datetime
    main()
