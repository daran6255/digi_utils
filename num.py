import pandas as pd
import re

# Define your custom mapping of alphanumeric characters to numbers
custom_mapping = {
    '0': '0', '1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6', '7': '7', '8': '8',
    '9': '9', 'Q': '0', 'W': '2', 'R': '0', 'T': '7',
    'ZD': '20', '£': '3', '$': '5', '€': '3', 'O': '0', 'I': '1', 'l': '1', 'Z': '2', 'S': '5', 'B': '8',
    '/': '1', '|': '1','(': '1',')': '1','':' ',

    # Add more mappings as needed
}

def excel_to_numeric_index(excel_index):
    """
    Convert Excel-style column index (e.g., 'A', 'B', 'AA') to numeric index (0-based).
    Supports columns up to 'ZZZ'.
    """
    index = 0
    for char in excel_index:
        if 'A' <= char <= 'Z':
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError("Invalid Excel column index")
    return index - 1  # Convert to 0-based index

def correct_numbers_in_excel(input_file, output_file, columns_to_correct):
    def apply_custom_mapping(text):
        for key, value in custom_mapping.items():
            if len(key) > 1:  # Check if the key has more than one character
                if key in text:  # Check if the key is present in the text
                    text = text.replace(key, value)
            else:
                text = re.sub(re.escape(key), value, text, flags=re.IGNORECASE)
        return text

    def extract_and_correct_numbers(text):
        # Try to convert the text to an integer
        try:
            value = float(text)  # Try to convert to float
            if value.is_integer():  # Check if the value is an integer
                return str(int(value))  # Return as an integer string
            else:
                return str(value)  # Return as a float string
        except ValueError:
            # If conversion to float fails, apply custom mapping
            return apply_custom_mapping(text)

    try:
        # Load the Excel file
        df = pd.read_excel(input_file)

        # Apply correction to specified columns
        for excel_index in columns_to_correct:
            col_index = excel_to_numeric_index(excel_index)
            col_name = df.columns[col_index]
            df[col_name] = df[col_name].apply(lambda x: extract_and_correct_numbers(str(x)))

        
        
        df.to_excel(output_file, index=False)
        print(f"Corrected numbers saved to {output_file}")
        return output_file
    except Exception as e:
        print(f"Error correcting numbers: {e}")
        return None