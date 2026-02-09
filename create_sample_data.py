import pandas as pd

data = {
    'SSA': ['SSA-001', 'SSA-002'],
    'Billing Account': ['ACC-12345', 'ACC-67890'],
    'CUSTOMER NAME': ['John Smith', 'Jane Doe'],
    'Accot Subtype': ['Premium', 'Standard'],
    'Department': ['Sales', 'Support'],
    'Address': ['123 Main St, City', '456 Oak Ave, Town'],
    'Status(Active/Inactive)': ['Active', 'Active'],
    'Outstanding amount in Rs': [5000.50, 2500.75],
    'CLOSURE DATE': ['2026-03-01', '2026-04-15']
}

df = pd.DataFrame(data)
df.to_excel('sample_data.xlsx', index=False)
print("✅ Created sample_data.xlsx")
