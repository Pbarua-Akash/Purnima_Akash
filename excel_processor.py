import pandas as pd
import os
from datetime import datetime  # Import datetime for timestamp

def process_excel(input_file, output_dir):
    """
    Processes the Excel file and exports each column (except the first) as a CSV file.
    
    Args:
        input_file (str): Path to the input Excel file.
        output_dir (str): Path to the output directory.
    """
    # Generate a unique folder name using a timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Format: YYYYMMDD_HHMMSS
    exported_files_dir = os.path.join(output_dir, f"exported_files_{timestamp}")
    
    # Create the folder
    os.makedirs(exported_files_dir)
    print(f"Created folder: {exported_files_dir}")

    # Load the large Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Get the first column (this will always be included)
    first_column = df.columns[0]

    # Loop through each column (excluding the first column) and export them with the first column
    for column in df.columns[1:]:
        # Clean the column name by removing trailing commas and replacing spaces with underscores
        cleaned_column = column.rstrip(',').replace(' ', '_')

        # Select the first column and the current column
        df_selected = df[[first_column, column]]
        
        # Construct the output file name using the cleaned column name
        output_file = os.path.join(exported_files_dir, f"First_Column_and_{cleaned_column}.csv")
        
        # Save the selected columns as a CSV file
        df_selected.to_csv(output_file, index=False)
        
        print(f"Exported: {output_file}")