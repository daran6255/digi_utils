import pandas as pd
from datetime import datetime

def _convert_to_dd_mm_yyyy(date_string):
    formats = ['%d-%m-%Y', '%d/%m/%Y', '%d %b %Y', '%d %B %Y',
               '%d-%b-%y', '%Y-%m-%d', '%A, %d %B %Y', '%d/%m/%y', '%d-%b-%Y', '%d.%m.%y', '%dth %B %Y', '%d.%m.%Y']
    for fmt in formats:
        try:
            parsed_date = datetime.strptime(date_string, fmt)
            return parsed_date.date().strftime('%d/%m/%Y')  # Return date string without time
        except ValueError:
            pass
    return date_string  # Return the original string if no format matched

def clean(input_file,cleaned_file,header):
    # Get the Excel file path from the user
    file_path = input_file

    # Load the Excel file
    data = pd.read_excel(file_path)

    # Ask the user which row to use as the header
    header_row = header

    # Set the header
    header = data.iloc[header_row - 1]
    data = data.iloc[header_row:]

    # Set the header
    data.columns = header

    # Remove empty rows (rows with all NaN values)
    data_cleaned = data.dropna(how='all')


    # Iterate over each column and apply the date conversion function if the column type is object (string)
    for column in data_cleaned.columns:
        if data_cleaned[column].dtype == object:
            data_cleaned[column] = data_cleaned[column].apply(lambda x: _convert_to_dd_mm_yyyy(x) if isinstance(x, str) else x)

    # Save the modified DataFrame back to a new Excel file without time in dates
    data_cleaned.to_excel(cleaned_file, index=False)

    print("Data cleaned and processed. Saved to .xlsx")
    
    return cleaned_file
