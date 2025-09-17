import pandas as pd
import numpy as np

# Generate sample data
def create_sample_data():
    np.random.seed(42)
    
    # Base data
    articles = [f"{i:012d}" for i in range(1001, 1011)]
    oms = [1001, 1002, 1003, 1004]
    locations = ['Warehouse A', 'Warehouse B', 'Warehouse C', 'Warehouse D', 'Warehouse E']
    
    data = []
    for article in articles:
        for om in oms:
            # Generate reasonable inventory and sales data
            sales = np.random.randint(20, 100)
            stock = np.random.randint(50, 200)
            location = np.random.choice(locations)
            
            data.append({
                'Article': article,
                'OM': om,
                'Inventory': stock,
                'Sales': sales,
                'Location': location,
                'Safety Stock': sales * 1.2,
                'Pending Received': np.random.randint(0, 50)
            })
    
    df = pd.DataFrame(data)
    return df

# Save sample file
if __name__ == "__main__":
    df = create_sample_data()
    df.to_excel("sample_inventory_data.xlsx", index=False)
    print("Sample file generated: sample_inventory_data.xlsx")