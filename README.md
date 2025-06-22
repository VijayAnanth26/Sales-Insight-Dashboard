# Sales Data Preprocessing Tool

This tool preprocesses the Superstore Sales data for analysis and visualization.

## Features

- **Data Loading**: Loads the Excel file with fallback to different engines
- **Data Exploration**: Provides basic statistics and information about the dataset
- **Data Preprocessing**:
  - Handles missing values
  - Converts date columns and extracts features
  - Detects and handles outliers
- **Data Visualization**:
  - Creates distribution plots for numeric columns
  - Creates bar plots for categorical columns
  - Generates correlation matrix for numeric columns
- **Data Export**: Saves processed data to CSV and Excel formats

## Requirements

```
pandas
numpy
matplotlib
seaborn
openpyxl
xlrd
```

## Installation

Install the required packages:

```bash
pip install pandas numpy matplotlib seaborn openpyxl xlrd
```

## Usage

1. Place the script in the same directory as your "Sales Data - Superstore.xls" file
2. Run the script:

```bash
python preprocess_sales_data.py
```

3. The script will:
   - Create a `processed_data` directory with the cleaned dataset
   - Create a `visualizations` directory with generated plots

## Output

- **Processed Data**: CSV and Excel files in the `processed_data` directory
- **Visualizations**: PNG files in the `visualizations` directory

## Customization

You can modify the script to:
- Change input file path by editing the `file_path` parameter in `load_data()`
- Change output directories by editing parameters in `save_processed_data()` and `generate_basic_visualizations()`
- Add more preprocessing steps in the `preprocess_data()` function
- Add more visualizations in the `generate_basic_visualizations()` function 