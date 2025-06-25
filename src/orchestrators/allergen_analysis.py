from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side

import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))
from src.modules.IO import IOHandler

import pandas as pd

def clean_data(df1, df2):
    def format_allergens(df: pd.DataFrame) -> pd.DataFrame:
        # Identify columns that contain allergen data
        allergy_columns = [col for col in df.columns if 'Allerg' in col]
        
        # Dictionary of known allergen name corrections
        allergy_dict = {
            "SOYBEA": "SOYBEAN",
            "PINNUT": "PINE NUT",
            "TREE": "TREE NUT",
            "FILBER": "FILBERT"
        }

        # Apply corrections to each allergen column
        for col in allergy_columns:
            df[col] = df[col].apply(
                lambda x: (
                    next((x.replace(k, v) for k, v in allergy_dict.items() if isinstance(x, str) and k in x), x)
                )
            )
            df[col] = df[col].astype("category")

        return df.dropna(subset=allergy_columns, how='all')

    # Clean 'Prod Numb' column to remove quotes and '='
    df2 = df2.copy()
    df2['Prod Numb'] = df2['Prod Numb'].str.replace('=', '', regex=False).str.replace('"', '', regex=False).str.strip()
    
    # Merge on the respective columns
    merged = df1.merge(df2, left_on='Product Number', right_on='Prod Numb', how='inner')
    
    # Convert datatypes
    merged['Date'] = pd.to_datetime(merged['Date'], errors='coerce')
    merged['Ship Quantity'] = pd.to_numeric(merged['Ship Quantity'], errors='coerce')
    merged['Net Sales'] = pd.to_numeric(merged['Net Sales'], errors='coerce')
    merged['Product Number'] = merged['Product Number'].astype("string")
    merged['Product'] = merged['Product'].astype("string")
    
    # Rename columns for clarity
    merged.rename(columns={
        'Product Number': 'Product_ID',
        'Product': 'Product_Description',
        'Ship Quantity': 'QTY_Shipped',
        'Net Sales': 'Sales'
    }, inplace=True)
    
    # Drop irrelevant columns
    merged.drop(columns=["Product Code", "Prod Numb", "Description", "Actv", "Allerg 8", "Allerg 9", "Warehouse Code"], inplace=True)

    # Handle allergen cols
    df = format_allergens(merged).reset_index(drop=True)
    
    return df

def calculate_allergen_metrics(df: pd.DataFrame):
    allergen_columns = [col for col in df.columns if col.startswith('Allerg')]
    all_allergens = set()
    for col in allergen_columns:
        all_allergens.update(df[col].dropna().unique())

    total_shipped_total = df['QTY_Shipped'].sum()
    total_sales = df['Sales'].sum()
    total_unique = df['Product_ID'].nunique()

    metrics_list = []
    for allergen in all_allergens:
        mask = df[allergen_columns].apply(lambda row: allergen in row.values, axis=1)
        total_units = df[mask]['QTY_Shipped'].sum()
        sales = df[mask]['Sales'].sum()
        unique_products = df[mask]['Product_ID'].nunique()

        metrics_list.append({
            'Allergen': allergen,
            'Total_Shipped': total_units.round(4),
            'Ratio_Shipped': round(total_units / total_shipped_total, 4) if total_shipped_total else 0,
            'Ratio_Unique_Products': round(unique_products / total_unique, 4) if total_unique else 0,
            'Ratio_Sales': round(sales / total_sales, 4) if total_shipped_total else 0
        })

    return pd.DataFrame(metrics_list)

def export_to_excel_with_charts(df, output_filename, y_axis_max=None):
    """
    Export DataFrame to Excel with a single chart showing all metrics per allergen,
    including Total_Shipped and ratio metrics. The chart is enlarged for better visibility.
    """
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Allergens', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Allergens']

        # Format headers
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Metrics to include in the chart
        metrics = ['Ratio_Shipped', 'Ratio_Unique_Products', 'Ratio_Sales']
        chart = workbook.add_chart({'type': 'column'})

        for metric in metrics:
            col_index = df.columns.get_loc(metric)
            chart.add_series({
                'name':       ['Allergens', 0, col_index],
                'categories': ['Allergens', 1, 0, len(df), 0],
                'values':     ['Allergens', 1, col_index, len(df), col_index],
            })

        chart.set_title({'name': '2024 Metrics by Allergen (SP)'})
        chart.set_x_axis({'name': 'Allergen'})

        y_axis_params = {'name': 'Value'}
        if y_axis_max is not None:
            y_axis_params['max'] = y_axis_max
        chart.set_y_axis(y_axis_params)

        chart.set_size({'width': 900, 'height': 625})  # Enlarged chart

        worksheet.insert_chart(1, len(df.columns) + 2, chart)

        # Add a second sheet for bullet point notes
        notes_sheet = workbook.add_worksheet('Notes')
        notes = [
            "Setup:",
            "• Product_Allergens.csv: Called Ryan Jahnke to pull allergen data for all active and inactive products",
            "• Product_Shipments.csv: Cognos -> Team Content -> Jake Sommerfeldt -> YEARLY -> Allergen_Analysis -> Ellipses -> Edit: Adjust date filter and run as CSV",
            "• Pull code from Github at: ''",
            "• Rename and place CSV files in project directory: assets\\ip"
        ]
        for row, note in enumerate(notes):
            notes_sheet.write(row, 0, note)

if __name__ == "__main__":
    # Import data
    products_shipped = IOHandler.import_csv("assets\\ip\\Product_Shipments.csv")
    product_allergens = IOHandler.import_csv("assets\\ip\\Product_Allergens.csv")

    # Merge, rename columns, adj data types, drop irrelevant data
    data = clean_data(products_shipped, product_allergens)

    # Add ftrs for each allergen
    data = calculate_allergen_metrics(data)

    # Export
    #print(data.info())
    #IOHandler.export_csv(data, "assets\\op\\2024_SP_Allergen_Analysis.csv")
    export_to_excel_with_charts(data, "assets\\op\\2024_SP_Allergen_Analysis.xlsx", 1)
