#!/usr/bin/env python3
"""
Simple script to format a CSV file for AMAIN with no headers
This addresses the 'Invalid quantity found: Sku,Qty' error
"""

import pandas as pd
import os
import sys

def format_amain_csv(input_file, output_file=None):
    """Process an input file into AMAIN format with no headers"""
    print(f"Reading {input_file}...")
    
    # Try to read the input file
    try:
        if input_file.lower().endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading file: {str(e)}")
        return
    
    print(f"Found {len(df)} rows in file")
    
    # Try to find SKU and QTY columns
    sku_columns = [col for col in df.columns if 'sku' in col.lower() or 'item' in col.lower() or 'part' in col.lower()]
    qty_columns = [col for col in df.columns if 'qty' in col.lower() or 'quantity' in col.lower()]
    
    if not sku_columns or not qty_columns:
        print("Error: Could not find SKU and QTY columns. File must have columns for item SKU and quantity.")
        return
    
    # Select the best columns to use
    sku_col = 'Sku' if 'Sku' in df.columns else sku_columns[0]
    qty_col = 'Qty' if 'Qty' in df.columns else qty_columns[0]
    
    print(f"Using columns: {sku_col} and {qty_col}")
    
    # Create output file name if not provided
    if not output_file:
        base = os.path.splitext(input_file)[0]
        output_file = f"{base}_AMAIN_fixed.csv"
    
    # Write the CSV file without headers
    with open(output_file, 'w') as f:
        # Write each row without headers
        for _, row in df.iterrows():
            f.write(f"{row[sku_col]},{row[qty_col]}\n")
    
    print(f"Successfully created AMAIN-compatible file: {output_file}")
    print("This file format will work with AMAIN's import system without 'Invalid quantity found' errors")

if __name__ == "__main__":
    # Check if a file was provided
    if len(sys.argv) < 2:
        print("Usage: python amain_fix.py <input_file.csv>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    format_amain_csv(input_file, output_file) 