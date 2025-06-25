"""
=== IO ===
Encapsulates import and export functions as well as common DB interactions

"""
import os
import pandas as pd

class IOHandler:
    def __init__(self):
        """
        Initialize

        Args:
        """

    def import_csv(file_path: str):
        """
        Loads a CSV file from a file path

        Args:
            file_path (str): Name of CSV file to import
        """

        try:
            # Read the CSV file into a DataFrame
            df = pd.read_csv(file_path, dtype=str)
            #print(f"Loaded data from {file_path} successfully")
            return df
        except Exception as e:
            #print(f"Error reading {file_path}.csv: {e}")
            return None

    def export_csv(df: pd.DataFrame, file_path: str):
        """
        Saves a CSV file to a file path

        Args:
            df (pd.DataFrame): Data to be saved
            table_name (str): Name of the table to export
        """
        try:
            df.to_csv(file_path)
            print(f"Saved data to {file_path} successfully")
        except Exception as e:
            print(f"Error saving data to {file_path}: {e}")
            raise

    def import_excel_files(directory_path):
        dataframes = []
        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(directory_path, filename)
                xl = pd.ExcelFile(file_path)
                df = xl.parse(xl.sheet_names[0])
                if 'Customer' in df.columns:
                    df = df.drop(columns=['Customer'])
                if 'Customer Code' not in df.columns:
                    raise ValueError(f"'Customer Code' column not found in the first sheet of {filename}")
                df = df[['Customer Code']]
                dataframes.append(df)
        return dataframes
