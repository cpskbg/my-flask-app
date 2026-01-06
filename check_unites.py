from app import app, db, Unite, Service, BudgetNature
import os

with app.app_context():
    print("--- UNITES AND SERVICES ---")
    unites = Unite.query.all()
    for u in unites:
        print(f"UNITE: {u.nom} (ID: {u.id})")
        services = Service.query.filter_by(unite_id=u.id).all()
        for s in services:
            print(f"  - SERVICE: {s.nom} (ID: {s.id})")
    
    print("\n--- BUDGET SAMPLES ---")
    budgets = BudgetNature.query.limit(5).all()
    for b in budgets:
        print(f"BUDGET: {b.nature} ({b.annee})")
        print(f"  Centre TTC: {b.budget_centre_ttc}")
        print(f"  Unite TTC: {b.budget_unite_ttc}")
        print(f"  Calculated Total: {b.budget_total}")
