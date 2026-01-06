#!/usr/bin/env python3
"""
Script pour migrer le stock de l'ancien système (Item.quantity) vers le nouveau système FIFO (ReceptionStock)
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app import app, db, Item, Reception, ReceptionStock, User
from datetime import datetime

def migrate_stock_to_fifo():
    """Migre le stock des articles vers le système FIFO"""
    with app.app_context():
        print("🔄 Migration du stock vers le système FIFO...")
        
        # Récupérer l'utilisateur admin pour la réception
        admin_user = User.query.filter_by(role='admin').first()
        if not admin_user:
            print("❌ Aucun utilisateur admin trouvé")
            return
        
        # Récupérer les articles qui ont du stock mais pas de ReceptionStock
        items_to_migrate = []
        for item in Item.query.filter(Item.quantity > 0).all():
            stock_entries = ReceptionStock.query.filter_by(item_id=item.id).all()
            if not stock_entries:
                items_to_migrate.append(item)
        
        print(f"📦 Articles à migrer : {len(items_to_migrate)}")
        
        for item in items_to_migrate:
            print(f"\n🔄 Migration de : {item.name} (Stock: {item.quantity})")
            
            # Créer une réception fictive pour migrer le stock
            reception = Reception(
                numero=f"MIG-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                fournisseur_id=1,  # Utiliser le premier fournisseur
                user_id=admin_user.id,
                date_reception=datetime.now().date(),
                notes=f"Réception de migration pour {item.name}"
            )
            db.session.add(reception)
            db.session.flush()  # Pour obtenir l'ID
            
            # Créer l'entrée de stock FIFO
            stock_entry = ReceptionStock(
                reception_id=reception.id,
                item_id=item.id,
                quantity=item.quantity,
                quantite_restante=item.quantity,
                prix_unitaire_ht=0.0,  # Prix par défaut
                prix_unitaire_ttc=0.0,
                taux_tva=20.0,
                date_reception=datetime.now()
            )
            db.session.add(stock_entry)
            
            # Mettre à jour la quantité de l'article à 0 (ancien système)
            item.quantity = 0
            
            print(f"✅ Créé : Réception #{reception.numero} avec {item.quantity} unités")
        
        try:
            db.session.commit()
            print(f"\n✅ Migration terminée avec succès !")
            print(f"✅ {len(items_to_migrate)} articles migrés vers le système FIFO")
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ Erreur lors de la migration : {e}")

if __name__ == "__main__":
    migrate_stock_to_fifo()
