import pandas as pd
import argparse
import sys

def redistribute_missing_values(data):
    data = data.copy()  # Work on a copy of the data
    i = 0
    while i < len(data):
        if pd.isna(data[i]):  # If current value is NaN
            # Find the next non-NaN value and count consecutive NaNs
            start_idx = i
            while i < len(data) and pd.isna(data[i]):
                i += 1
            end_idx = i  # Index of the next non-NaN value
            
            if i < len(data):  # Only proceed if a next non-NaN value exists
                next_value = data[i]
                count_nans = end_idx - start_idx
                
                # Calculate the equal share, round to nearest integer, and fill the NaNs
                if count_nans > 0:
                    equal_share = round(next_value / (count_nans + 1))
                    data[start_idx:end_idx] = [equal_share] * count_nans
                    data[i] = equal_share  # Update the next non-NaN value as well

        i += 1  # Move to the next value

    return data

def main(file_path, sheet_name, column_index, full_output_file, column_output_file):
    # Load the specified sheet
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error loading file or sheet: {e}")
        sys.exit(1)
    
    # Check if the specified column index is within range
    if column_index < 0 or column_index >= len(df.columns):
        print(f"Column index '{column_index}' is out of range for sheet '{sheet_name}'.")
        sys.exit(1)

    # Extract the specified column by index
    column_data = df.iloc[:, column_index]
    
    # Apply the redistribution function
    redistributed_column_data = redistribute_missing_values(column_data)

    # Update the original DataFrame with the rounded redistributed column data
    df.iloc[:, column_index] = redistributed_column_data

    # Save the entire modified DataFrame to a new Excel file
    df.to_excel(full_output_file, index=False)
    print(f"Redistributed data with updates in column index '{column_index}' has been saved to {full_output_file}")

    # Save only the redistributed column data to a CSV file, rounded to integers
    redistributed_column_data = redistributed_column_data.astype(int)  # Ensure no decimal places
    redistributed_column_data.to_csv(column_output_file, index=False)
    print(f"Redistributed column data from index '{column_index}' has been saved to {column_output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Redistribute missing values in a specified column of an Excel sheet, save the entire modified sheet, and output the column data.")
    parser.add_argument("file_path", help="Path to the Excel file.")
    parser.add_argument("sheet_name", help="Name of the sheet to read.")
    parser.add_argument("column_index", type=int, help="Index of the column to process (0-based).")
    parser.add_argument("-f", "--full_output", default="redistributed_full_output.xlsx", help="Output file name for the entire modified sheet.")
    parser.add_argument("-c", "--column_output", default="redistributed_column.csv", help="Output file name for the modified column data.")

    args = parser.parse_args()
    main(args.file_path, args.sheet_name, args.column_index, args.full_output, args.column_output)
