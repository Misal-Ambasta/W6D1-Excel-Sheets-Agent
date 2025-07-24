#!/usr/bin/env python3
"""
Create sample Excel data for testing the Excel Sheets Agent
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def create_sample_data():
    """Create sample Excel file with multiple sheets and various data types"""
    
    # Set random seed for reproducibility
    np.random.seed(42)
    random.seed(42)
    
    # Create sample data for different sheets
    
    # Sheet 1: Sales Data
    print("Creating Sales Data sheet...")
    sales_data = []
    regions = ['North', 'South', 'East', 'West', 'Central']
    products = ['Product A', 'Product B', 'Product C', 'Product D', 'Product E']
    
    for i in range(5000):  # 5000 rows for testing large file handling
        sales_data.append({
            'Order_ID': f'ORD-{i+1:05d}',
            'Date': datetime(2024, 1, 1) + timedelta(days=random.randint(0, 365)),
            'Customer_Name': f'Customer {i+1}',
            'Region': random.choice(regions),
            'Product': random.choice(products),
            'Quantity': random.randint(1, 100),
            'Unit_Price': round(random.uniform(10, 1000), 2),
            'Total_Amount': 0,  # Will calculate below
            'Sales_Rep': f'Rep {random.randint(1, 20)}'
        })
    
    sales_df = pd.DataFrame(sales_data)
    sales_df['Total_Amount'] = sales_df['Quantity'] * sales_df['Unit_Price']
    
    # Sheet 2: Customer Data
    print("Creating Customer Data sheet...")
    customer_data = []
    cities = ['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix', 'Philadelphia']
    
    for i in range(1000):
        customer_data.append({
            'Customer_ID': f'CUST-{i+1:04d}',
            'Customer_Name': f'Customer {i+1}',
            'Email': f'customer{i+1}@email.com',
            'Phone': f'+1-{random.randint(100,999)}-{random.randint(100,999)}-{random.randint(1000,9999)}',
            'City': random.choice(cities),
            'Registration_Date': datetime(2023, 1, 1) + timedelta(days=random.randint(0, 730)),
            'Status': random.choice(['Active', 'Inactive', 'Pending']),
            'Credit_Limit': random.randint(1000, 50000)
        })
    
    customer_df = pd.DataFrame(customer_data)
    
    # Sheet 3: Product Inventory
    print("Creating Product Inventory sheet...")
    inventory_data = []
    categories = ['Electronics', 'Clothing', 'Home & Garden', 'Sports', 'Books']
    
    for i, product in enumerate(products):
        inventory_data.append({
            'Product_Code': f'PROD-{i+1:03d}',
            'Product_Name': product,
            'Category': random.choice(categories),
            'Stock_Quantity': random.randint(0, 1000),
            'Reorder_Level': random.randint(10, 100),
            'Unit_Cost': round(random.uniform(5, 500), 2),
            'Selling_Price': round(random.uniform(10, 1000), 2),
            'Supplier': f'Supplier {random.randint(1, 10)}',
            'Last_Updated': datetime.now() - timedelta(days=random.randint(0, 30))
        })
    
    inventory_df = pd.DataFrame(inventory_data)
    
    # Sheet 4: Financial Summary (with some edge cases)
    print("Creating Financial Summary sheet...")
    financial_data = []
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for month in months:
        financial_data.append({
            'Month': month,
            'Revenue': random.randint(50000, 200000),
            'Expenses': random.randint(30000, 150000),
            'Profit': 0,  # Will calculate
            'Growth_Rate': round(random.uniform(-10, 25), 2),
            'Notes': f'Performance for {month} 2024' if random.random() > 0.3 else None  # Some null values
        })
    
    financial_df = pd.DataFrame(financial_data)
    financial_df['Profit'] = financial_df['Revenue'] - financial_df['Expenses']
    
    # Create Excel file with multiple sheets
    print("Writing to Excel file...")
    with pd.ExcelWriter('sample_data.xlsx', engine='openpyxl') as writer:
        sales_df.to_excel(writer, sheet_name='Sales_Data', index=False)
        customer_df.to_excel(writer, sheet_name='Customer_Data', index=False)
        inventory_df.to_excel(writer, sheet_name='Product_Inventory', index=False)
        financial_df.to_excel(writer, sheet_name='Financial_Summary', index=False)
    
    print("✅ Sample Excel file 'sample_data.xlsx' created successfully!")
    print(f"📊 File contains:")
    print(f"   - Sales_Data: {len(sales_df)} rows")
    print(f"   - Customer_Data: {len(customer_df)} rows")
    print(f"   - Product_Inventory: {len(inventory_df)} rows")
    print(f"   - Financial_Summary: {len(financial_df)} rows")

if __name__ == "__main__":
    create_sample_data()
