import pandas as pd
import numpy as np
import os
from datetime import datetime
import re

def load_data(file_path='Sales Data - Superstore.xls'):
    """
    Load data from Excel file
    """
    print(f"Loading data from {file_path}...")
    try:
        # Try with default engine
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error with default engine: {e}")
        # Try with xlrd engine for older .xls files
        try:
            df = pd.read_excel(file_path, engine='xlrd')
        except Exception as e:
            print(f"Error with xlrd engine: {e}")
            return None
    
    print(f"Data loaded successfully. Shape: {df.shape}")
    return df

def preprocess_for_powerbi(df):
    """
    Preprocess the data specifically for Power BI
    """
    if df is None:
        print("No data to preprocess.")
        return None
    
    print("\n===== PREPROCESSING FOR POWER BI =====")
    processed_df = df.copy()
    
    # 1. Handle missing values
    missing_before = processed_df.isnull().sum().sum()
    print(f"Missing values before preprocessing: {missing_before}")
    
    # Fill numeric columns with mean
    numeric_cols = processed_df.select_dtypes(include=['number']).columns
    for col in numeric_cols:
        if processed_df[col].isnull().sum() > 0:
            processed_df[col].fillna(processed_df[col].mean(), inplace=True)
    
    # Fill categorical columns with mode
    cat_cols = processed_df.select_dtypes(include=['object']).columns
    for col in cat_cols:
        if processed_df[col].isnull().sum() > 0:
            processed_df[col].fillna(processed_df[col].mode()[0], inplace=True)
    
    # 2. Optimize date columns for Power BI
    date_cols = []
    for col in processed_df.columns:
        if 'date' in col.lower() or 'time' in col.lower():
            try:
                processed_df[col] = pd.to_datetime(processed_df[col])
                date_cols.append(col)
                print(f"Converted {col} to datetime")
            except:
                print(f"Could not convert {col} to datetime")
    
    # 3. Create date dimension table for Power BI
    if date_cols:
        date_dim = create_date_dimension(processed_df, date_cols[0])
        print(f"Created date dimension table with {len(date_dim)} rows")
    else:
        date_dim = None
        print("No date columns found to create date dimension")
    
    # 4. Clean and standardize text fields
    for col in cat_cols:
        # Convert to string if not already
        processed_df[col] = processed_df[col].astype(str)
        # Remove extra spaces
        processed_df[col] = processed_df[col].str.strip()
        # Standardize case for better grouping in Power BI
        processed_df[col] = processed_df[col].str.title()
    
    # 5. Create geography dimension for Power BI if geographic data exists
    geo_cols = [col for col in processed_df.columns if col.lower() in 
                ['country', 'region', 'state', 'city', 'postal code', 'zip', 'zipcode']]
    
    if geo_cols:
        geo_dim = create_geography_dimension(processed_df, geo_cols)
        print(f"Created geography dimension table with {len(geo_dim)} rows")
    else:
        geo_dim = None
        print("No geography columns found to create geography dimension")
    
    # 6. Create product dimension for Power BI if product data exists
    product_cols = [col for col in processed_df.columns if col.lower() in 
                   ['product id', 'product name', 'category', 'sub-category', 'subcategory']]
    
    if product_cols:
        product_dim = create_product_dimension(processed_df, product_cols)
        print(f"Created product dimension table with {len(product_dim)} rows")
    else:
        product_dim = None
        print("No product columns found to create product dimension")
    
    # 7. Create customer dimension for Power BI if customer data exists
    customer_cols = [col for col in processed_df.columns if col.lower() in 
                    ['customer id', 'customer name', 'segment', 'customer segment']]
    
    if customer_cols:
        customer_dim = create_customer_dimension(processed_df, customer_cols)
        print(f"Created customer dimension table with {len(customer_dim)} rows")
    else:
        customer_dim = None
        print("No customer columns found to create customer dimension")
    
    # 8. Create fact table (sales/orders)
    fact_table = create_fact_table(processed_df)
    print(f"Created fact table with {len(fact_table)} rows")
    
    # Check for missing values after preprocessing
    missing_after = processed_df.isnull().sum().sum()
    print(f"Missing values after preprocessing: {missing_after}")
    
    # Return all tables as a dictionary
    tables = {
        'processed_data': processed_df,
        'fact_table': fact_table
    }
    
    if date_dim is not None:
        tables['date_dimension'] = date_dim
    
    if geo_dim is not None:
        tables['geography_dimension'] = geo_dim
    
    if product_dim is not None:
        tables['product_dimension'] = product_dim
    
    if customer_dim is not None:
        tables['customer_dimension'] = customer_dim
    
    return tables

def create_date_dimension(df, date_column):
    """
    Create a date dimension table for Power BI
    """
    # Get min and max dates from the data
    min_date = df[date_column].min()
    max_date = df[date_column].max()
    
    # Create a range of dates
    date_range = pd.date_range(start=min_date, end=max_date, freq='D')
    
    # Create the date dimension table
    date_dim = pd.DataFrame({
        'Date': date_range,
        'Day': date_range.day,
        'Month': date_range.month,
        'MonthName': date_range.strftime('%B'),
        'Quarter': date_range.quarter,
        'Year': date_range.year,
        'DayOfWeek': date_range.dayofweek,
        'DayName': date_range.strftime('%A'),
        'WeekOfYear': date_range.isocalendar().week,
        'IsWeekend': date_range.dayofweek.isin([5, 6]),
        'IsMonthEnd': date_range.is_month_end,
        'IsMonthStart': date_range.is_month_start,
        'IsQuarterEnd': date_range.is_quarter_end,
        'IsQuarterStart': date_range.is_quarter_start,
        'IsYearEnd': date_range.is_year_end,
        'IsYearStart': date_range.is_year_start
    })
    
    # Create DateKey for relationships (YYYYMMDD format)
    date_dim['DateKey'] = date_dim['Date'].dt.strftime('%Y%m%d').astype(int)
    
    return date_dim

def create_geography_dimension(df, geo_cols):
    """
    Create a geography dimension table for Power BI
    """
    # Select only geographic columns
    geo_data = df[geo_cols].copy()
    
    # Drop duplicates to get unique combinations
    geo_data = geo_data.drop_duplicates().reset_index(drop=True)
    
    # Add a geography key
    geo_data['GeoKey'] = geo_data.index + 1
    
    return geo_data

def create_product_dimension(df, product_cols):
    """
    Create a product dimension table for Power BI
    """
    # Select only product columns
    product_data = df[product_cols].copy()
    
    # Drop duplicates to get unique products
    product_data = product_data.drop_duplicates().reset_index(drop=True)
    
    # If 'Product ID' is not in the columns, create a product key
    if 'Product ID' not in product_data.columns:
        product_data['ProductKey'] = product_data.index + 1
    else:
        # Use Product ID as the key but ensure it's clean and unique
        product_data['ProductKey'] = product_data['Product ID'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True)
    
    return product_data

def create_customer_dimension(df, customer_cols):
    """
    Create a customer dimension table for Power BI
    """
    # Select only customer columns
    customer_data = df[customer_cols].copy()
    
    # Drop duplicates to get unique customers
    customer_data = customer_data.drop_duplicates().reset_index(drop=True)
    
    # If 'Customer ID' is not in the columns, create a customer key
    if 'Customer ID' not in customer_data.columns:
        customer_data['CustomerKey'] = customer_data.index + 1
    else:
        # Use Customer ID as the key but ensure it's clean and unique
        customer_data['CustomerKey'] = customer_data['Customer ID'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True)
    
    return customer_data

def create_fact_table(df):
    """
    Create a fact table for Power BI
    """
    # Identify measure columns (numeric columns that are not IDs or keys)
    measure_cols = df.select_dtypes(include=['number']).columns.tolist()
    
    # Identify dimension key columns
    key_cols = [col for col in df.columns if 'id' in col.lower() or 'key' in col.lower() or 'order' in col.lower()]
    
    # Identify date columns
    date_cols = [col for col in df.columns if 'date' in col.lower()]
    
    # Create a list of columns to include in the fact table
    fact_cols = key_cols + date_cols + measure_cols
    
    # Create the fact table with available columns
    available_cols = [col for col in fact_cols if col in df.columns]
    fact_table = df[available_cols].copy()
    
    # Create a fact key
    fact_table['FactKey'] = fact_table.index + 1
    
    # For date columns, create DateKey in YYYYMMDD format for relationships
    for date_col in [col for col in date_cols if col in fact_table.columns]:
        new_col_name = date_col.replace(' ', '') + 'Key'
        # Fix: Access the datetime attribute of each value in the column instead of using .dt
        fact_table[new_col_name] = fact_table[date_col].apply(lambda x: int(x.strftime('%Y%m%d')) if pd.notnull(x) else 0)
    
    return fact_table

def save_tables_for_powerbi(tables, output_dir='powerbi_data'):
    """
    Save all tables to separate Excel files for Power BI
    """
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Generate timestamp for unique filenames
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create a single Excel file with multiple sheets
    excel_path = os.path.join(output_dir, f'powerbi_data_{timestamp}.xlsx')
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for table_name, table_data in tables.items():
            if table_data is not None:
                # Clean sheet name (Excel has 31 char limit and no special chars)
                sheet_name = re.sub(r'[^\w]', '_', table_name)[:31]
                table_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Saved {table_name} to Excel sheet")
    
    print(f"All tables saved to: {excel_path}")
    
    # Also save as CSV files for flexibility
    for table_name, table_data in tables.items():
        if table_data is not None:
            csv_path = os.path.join(output_dir, f'{table_name}_{timestamp}.csv')
            table_data.to_csv(csv_path, index=False)
            print(f"Saved {table_name} to CSV: {csv_path}")
    
    return excel_path

def main():
    """
    Main function to run the preprocessing pipeline for Power BI
    """
    print("Starting sales data preprocessing for Power BI...")
    
    # Load the data
    df = load_data()
    
    # Preprocess the data for Power BI
    tables = preprocess_for_powerbi(df)
    
    # Save the processed tables
    if tables:
        save_tables_for_powerbi(tables)
    
    print("Preprocessing for Power BI complete!")

if __name__ == "__main__":
    main() 