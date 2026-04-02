from db_manager import DatabaseManager

def setup():
    db = DatabaseManager()
    print("Initializing Database...")
    db.init_db()
    
    # 1. Definir Datos Maestros (Copia de los defaults en app.py)
    aspect_ranges = {
        "sales": (0, 20),
        "inventory_reduction": (0, 15),
        "margin": (0, 10),
        "labor": (0, 12),
        "logistics": (0, 8),
    }

    module_options = [
        "Inventory Optimization", "Pricing", "Merchandising", "Customer Experience",
        "Supply Chain", "Retail Insights", "Store Operations", "Loyalty", "Omnichannel", "Data Science"
    ]

    module_profiles = {
        "Inventory Optimization": ["inventory_reduction"],
        "Pricing": ["margin"],
        "Merchandising": ["sales"],
        "Customer Experience": ["sales", "margin"],
        "Supply Chain": ["logistics"],
        "Retail Insights": ["sales", "inventory_reduction"],
        "Store Operations": ["labor"],
        "Loyalty": ["sales"],
        "Omnichannel": ["sales", "logistics"],
        "Data Science": ["margin", "inventory_reduction"],
    }

    benefit_params = {}
    for module in module_options:
        benefit_params[module] = {}
        for aspect in module_profiles.get(module, []):
            min_val, max_val = aspect_ranges[aspect]
            benefit_params[module][aspect] = {"min": min_val + 1, "max": max_val - 1}

    annual_investments = {i: {'software': 0, 'impl': 0, 'extra': 0} for i in range(1, 11)}

    # 2. Sincronizar
    print("Syncing Default Configuration to Database...")
    db.sync_all(
        module_options,
        module_profiles,
        benefit_params,
        annual_investments,
        aspect_ranges
    )
    print("✅ Database setup complete!")

if __name__ == "__main__":
    setup()
