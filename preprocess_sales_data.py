import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime

def load_data(file_path="Sales Data - Superstore.xls"):
    """
    Load the sales data from Excel file with fallback to different engines
    """
    try:
        # Try loading with openpyxl first
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"Successfully loaded data with openpyxl: {file_path}")
    except Exception as e:
        print(f"Error loading with openpyxl: {e}")
        try:
            # Fallback to xlrd
            df = pd.read_excel(file_path, engine='xlrd')
            print(f"Successfully loaded data with xlrd: {file_path}")
        except Exception as e:
            print(f"Error loading with xlrd: {e}")
            raise Exception(f"Failed to load data from {file_path}")
    
    return df

def explore_data(df):
    """
    Explore the dataset and provide basic statistics
    """
    print("\n--- Data Exploration ---")
    print(f"Shape: {df.shape}")
    print("\nData Types:")
    print(df.dtypes)
    print("\nMissing Values:")
    print(df.isnull().sum())
    print("\nSample Data:")
    print(df.head())
    print("\nBasic Statistics:")
    print(df.describe())
    
    return df

def preprocess_data(df):
    """
    Preprocess the data for analysis
    """
    print("\n--- Data Preprocessing ---")
    
    # Create a copy to avoid modifying original data
    processed_df = df.copy()
    
    # Handle missing values
    print("Handling missing values...")
    for col in processed_df.columns:
        if processed_df[col].isnull().sum() > 0:
            if processed_df[col].dtype == 'object':
                # Fill missing categorical values with 'Unknown'
                processed_df[col].fillna('Unknown', inplace=True)
            else:
                # Fill missing numerical values with median
                processed_df[col].fillna(processed_df[col].median(), inplace=True)
    
    # Convert date columns and extract features
    print("Processing date columns...")
    date_columns = processed_df.select_dtypes(include=['datetime64']).columns.tolist()
    
    # If no datetime columns detected, try to convert potential date columns
    if not date_columns:
        potential_date_cols = ['Order Date', 'Ship Date']
        for col in potential_date_cols:
            if col in processed_df.columns:
                try:
                    processed_df[col] = pd.to_datetime(processed_df[col])
                    date_columns.append(col)
                except:
                    print(f"Could not convert {col} to datetime")
    
    # Extract features from date columns
    for col in date_columns:
        processed_df[f'{col}_Year'] = processed_df[col].dt.year
        processed_df[f'{col}_Month'] = processed_df[col].dt.month
        processed_df[f'{col}_Quarter'] = processed_df[col].dt.quarter
        processed_df[f'{col}_MonthName'] = processed_df[col].dt.month_name()
    
    # Handle outliers in numerical columns
    print("Detecting and handling outliers...")
    numeric_cols = processed_df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    
    for col in numeric_cols:
        # Skip ID columns and date features
        if 'ID' in col or 'Year' in col or 'Month' in col or 'Quarter' in col:
            continue
            
        # Calculate IQR
        Q1 = processed_df[col].quantile(0.25)
        Q3 = processed_df[col].quantile(0.75)
        IQR = Q3 - Q1
        
        # Define outlier bounds
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        # Count outliers
        outliers = ((processed_df[col] < lower_bound) | (processed_df[col] > upper_bound)).sum()
        if outliers > 0:
            print(f"Found {outliers} outliers in {col}")
            
            # Cap outliers instead of removing them
            processed_df[col] = np.where(
                processed_df[col] < lower_bound,
                lower_bound,
                np.where(
                    processed_df[col] > upper_bound,
                    upper_bound,
                    processed_df[col]
                )
            )
    
    # Add additional useful features
    print("Adding additional features...")
    
    # Check if relevant columns exist
    if 'Sales' in processed_df.columns and 'Profit' in processed_df.columns:
        # Calculate profit margin
        processed_df['Profit Margin'] = (processed_df['Profit'] / processed_df['Sales']) * 100
        
        # Handle infinite values from division by zero
        processed_df['Profit Margin'].replace([np.inf, -np.inf], np.nan, inplace=True)
        processed_df['Profit Margin'].fillna(0, inplace=True)
        
        # Profit category
        processed_df['Profit Category'] = pd.cut(
            processed_df['Profit'],
            bins=[-float('inf'), 0, processed_df['Profit'].median(), float('inf')],
            labels=['Loss', 'Low Profit', 'High Profit']
        )
    
    # Check if discount column exists
    if 'Discount' in processed_df.columns:
        # Discount category
        processed_df['Discount Category'] = pd.cut(
            processed_df['Discount'],
            bins=[-0.01, 0.01, 0.2, 0.5, 1.0],
            labels=['No Discount', 'Small Discount', 'Medium Discount', 'Large Discount']
        )
    
    return processed_df

def generate_basic_visualizations(df, output_dir='visualizations'):
    """
    Generate basic visualizations for the dataset
    """
    print("\n--- Generating Visualizations ---")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Set the style
    sns.set(style='whitegrid')
    plt.rcParams.update({'figure.figsize': (12, 8)})
    
    # Visualize numeric columns
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    for col in numeric_cols[:5]:  # Limit to first 5 to avoid too many plots
        if df[col].nunique() > 1:  # Only plot if there's variation
            plt.figure()
            sns.histplot(df[col], kde=True)
            plt.title(f'Distribution of {col}')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/{col}_distribution.png')
            plt.close()
    
    # Visualize categorical columns
    cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
    for col in cat_cols[:5]:  # Limit to first 5
        if df[col].nunique() < 15:  # Only plot if not too many categories
            plt.figure()
            sns.countplot(y=df[col], order=df[col].value_counts().index)
            plt.title(f'Count of {col}')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/{col}_counts.png')
            plt.close()
    
    # Correlation matrix for numeric columns
    plt.figure(figsize=(14, 10))
    numeric_df = df.select_dtypes(include=['float64', 'int64'])
    # Only include columns with reasonable number of values to avoid too large correlation matrix
    if numeric_df.shape[1] > 15:
        numeric_df = numeric_df.iloc[:, :15]
    
    corr = numeric_df.corr()
    mask = np.triu(np.ones_like(corr, dtype=bool))
    sns.heatmap(corr, mask=mask, annot=True, fmt='.2f', cmap='coolwarm', square=True)
    plt.title('Correlation Matrix')
    plt.tight_layout()
    plt.savefig(f'{output_dir}/correlation_matrix.png')
    plt.close()
    
    # Sales and Profit Analysis if columns exist
    if 'Sales' in df.columns and 'Profit' in df.columns:
        # Sales vs Profit Scatter plot
        plt.figure()
        sns.scatterplot(x='Sales', y='Profit', data=df, alpha=0.6)
        plt.title('Sales vs Profit')
        plt.tight_layout()
        plt.savefig(f'{output_dir}/sales_vs_profit.png')
        plt.close()
        
        # Check for Category and Region columns
        if 'Category' in df.columns:
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Category', y='Sales', data=df)
            plt.title('Sales by Category')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/sales_by_category.png')
            plt.close()
            
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Category', y='Profit', data=df)
            plt.title('Profit by Category')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/profit_by_category.png')
            plt.close()
        
        if 'Region' in df.columns:
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Region', y='Sales', data=df)
            plt.title('Sales by Region')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/sales_by_region.png')
            plt.close()
            
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Region', y='Profit', data=df)
            plt.title('Profit by Region')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/profit_by_region.png')
            plt.close()
    
    print(f"Visualizations saved to {output_dir}/")

def save_processed_data(df, output_dir='processed_data'):
    """
    Save the processed data to CSV and Excel formats
    """
    print("\n--- Saving Processed Data ---")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Generate timestamp for filenames
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Save to CSV
    csv_path = f"{output_dir}/processed_sales_data_{timestamp}.csv"
    df.to_csv(csv_path, index=False)
    print(f"Saved CSV: {csv_path}")
    
    # Save to Excel
    excel_path = f"{output_dir}/processed_sales_data_{timestamp}.xlsx"
    df.to_excel(excel_path, index=False)
    print(f"Saved Excel: {excel_path}")
    
    return csv_path, excel_path

def main():
    """
    Main function to run the preprocessing pipeline
    """
    print("=== Sales Data Preprocessing Tool ===")
    
    # Load data
    df = load_data()
    
    # Explore data
    df = explore_data(df)
    
    # Preprocess data
    processed_df = preprocess_data(df)
    
    # Generate visualizations
    generate_basic_visualizations(processed_df)
    
    # Save processed data
    csv_path, excel_path = save_processed_data(processed_df)
    
    print("\n=== Processing Complete ===")
    print(f"Original data shape: {df.shape}")
    print(f"Processed data shape: {processed_df.shape}")
    print(f"CSV saved to: {csv_path}")
    print(f"Excel saved to: {excel_path}")

if __name__ == "__main__":
    main() 