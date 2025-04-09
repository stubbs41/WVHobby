#!/usr/bin/env python3
import pandas as pd
import os

print("Creating sample PO data file for testing...")

# Sample data for a purchase order
data = {
    'Item': ['1001', '1002', '1003', '1004', '1005'],
    'Description': [
        'Widget A - Small',
        'Widget B - Medium',
        'Widget C - Large',
        'Special Tool Kit',
        'Maintenance Supplies'
    ],
    'Quantity': [10, 5, 8, 2, 3],
    'Price': [12.99, 24.50, 34.75, 149.99, 89.50],
    'UOM': ['EA', 'EA', 'EA', 'KIT', 'BOX'],
    'Category': ['Hardware', 'Hardware', 'Hardware', 'Tools', 'Maintenance'],
    'Vendor_SKU': ['V-1001', 'V-1002', 'V-1003', 'V-1004', 'V-1005']
}

# Create DataFrame
df = pd.DataFrame(data)

# Create a sample directory if it doesn't exist
sample_dir = 'sample'
if not os.path.exists(sample_dir):
    os.makedirs(sample_dir)

# Save as Excel file
output_path = os.path.join(sample_dir, 'PO12345.xlsx')
df.to_excel(output_path, index=False)

print(f"Sample PO file created: {output_path}")
print("Use this file to test the PO Formatter application.") 