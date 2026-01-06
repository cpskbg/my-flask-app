#!/usr/bin/env python3
"""
Script de migration pour ajouter les champs de décharge signée à la table dotations
Auteur: Système de Gestion Stock
Date: 2025-11-23
"""

import sqlite3
import os
import sys
from datetime import datetime

def add_decharge_signee_fields():
    """
    Ajoute les champs pour stocker les fichiers PDF de décharges signées
    à la table dotations existante
    """
    
    # Chemin vers la base de données
    db_path = 'stock_management.db'
    
    if not os.path.exists(db_path):
        print(f"❌ Erreur: Base de données introuvable à {db_path}")
        print("Veuillez exécuter ce script depuis le répertoire principal de l'application.")
        return False
    
    try:
        # Connexion à la base de données
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        print("🔄 Connexion à la base de données établie")
        
        # Vérifier si la table dotations existe
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='dotations'
        """)
        
        if not cursor.fetchone():
            print("❌ Erreur: La table 'dotations' n'existe pas")
            return False
        
        print("✅ Table 'dotations' trouvée")
        
        # Vérifier les colonnes existantes
        cursor.execute("PRAGMA table_info(dotations)")
        existing_columns = [row[1] for row in cursor.fetchall()]
        
        print(f"📋 Colonnes existantes: {', '.join(existing_columns)}")
        
        # Nouvelles colonnes à ajouter
        new_columns = [
            ('decharge_signee_path', 'VARCHAR(500)', 'Chemin vers le fichier PDF signé'),
            ('decharge_signee_filename', 'VARCHAR(200)', 'Nom original du fichier PDF'),
            ('decharge_signee_date', 'DATETIME', 'Date d\'import du PDF signé')
        ]
        
        columns_added = []
        
        for column_name, column_type, description in new_columns:
            if column_name not in existing_columns:
                try:
                    # Ajouter la colonne
                    alter_query = f"ALTER TABLE dotations ADD COLUMN {column_name} {column_type}"
                    cursor.execute(alter_query)
                    columns_added.append(column_name)
                    print(f"✅ Colonne '{column_name}' ajoutée ({description})")
                except sqlite3.Error as e:
                    print(f"❌ Erreur lors de l'ajout de la colonne '{column_name}': {e}")
                    return False
            else:
                print(f"ℹ️  La colonne '{column_name}' existe déjà")
        
        if columns_added:
            # Valider les changements
            conn.commit()
            print(f"\n🎉 Migration réussie! {len(columns_added)} colonne(s) ajoutée(s):")
            for col in columns_added:
                print(f"   • {col}")
        else:
            print("\nℹ️  Aucune modification nécessaire - toutes les colonnes existent déjà")
        
        # Afficher la structure mise à jour
        print("\n📊 Structure mise à jour de la table 'dotations':")
        cursor.execute("PRAGMA table_info(dotations)")
        columns_info = cursor.fetchall()
        
        for col in columns_info:
            col_name = col[1]
            col_type = col[2]
            nullable = "NULL" if col[3] == 0 else "NOT NULL"
            default = f"DEFAULT {col[4]}" if col[4] is not None else ""
            pk = "PRIMARY KEY" if col[5] == 1 else ""
            
            print(f"   • {col_name:<25} {col_type:<15} {nullable:<8} {default:<10} {pk}")
        
        # Créer le répertoire uploads si nécessaire
        upload_dir = os.path.join('uploads', 'decharges_signees')
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir, exist_ok=True)
            print(f"\n📁 Répertoire créé: {upload_dir}")
        else:
            print(f"\n📁 Répertoire déjà existant: {upload_dir}")
        
        return True
        
    except sqlite3.Error as e:
        print(f"❌ Erreur SQLite: {e}")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False
    finally:
        if conn:
            conn.close()
            print("🔒 Connexion à la base de données fermée")

def verify_migration():
    """
    Vérifie que la migration a été effectuée correctement
    """
    db_path = 'stock_management.db'
    
    if not os.path.exists(db_path):
        print("❌ Base de données introuvable")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Vérifier les nouvelles colonnes
        cursor.execute("PRAGMA table_info(dotations)")
        columns = [row[1] for row in cursor.fetchall()]
        
        required_columns = [
            'decharge_signee_path',
            'decharge_signee_filename', 
            'decharge_signee_date'
        ]
        
        missing_columns = [col for col in required_columns if col not in columns]
        
        if missing_columns:
            print(f"❌ Colonnes manquantes: {', '.join(missing_columns)}")
            return False
        else:
            print("✅ Tous les champs de décharge signée sont présents")
            return True
            
    except Exception as e:
        print(f"❌ Erreur lors de la vérification: {e}")
        return False
    finally:
        if conn:
            conn.close()

def main():
    """
    Fonction principale du script de migration
    """
    print("=" * 60)
    print("🚧 MIGRATION: Ajout des champs de décharge signée")
    print("=" * 60)
    print(f"📅 Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📁 Répertoire: {os.getcwd()}")
    print("-" * 60)
    
    # Exécuter la migration
    if add_decharge_signee_fields():
        print("\n" + "=" * 60)
        print("🔍 VÉRIFICATION DE LA MIGRATION")
        print("=" * 60)
        
        if verify_migration():
            print("\n🎉 Migration terminée avec succès!")
            print("\n📋 Prochaines étapes:")
            print("   1. Redémarrez l'application Flask")
            print("   2. Testez l'import de PDF signé pour une dotation livrée")
            print("   3. Vérifiez que les fichiers sont correctement sauvegardés")
        else:
            print("\n❌ La vérification a échoué")
            sys.exit(1)
    else:
        print("\n❌ La migration a échoué")
        sys.exit(1)

if __name__ == "__main__":
    main()
