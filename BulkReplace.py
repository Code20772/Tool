import pandas as pd

def bulk_replace_names(input_file, mapping_file, output_file):
    # Read input Excel file
    df = pd.read_excel(input_file, engine='openpyxl')  # Specify the engine
    
    # Read mapping file containing old and new names
    mapping_df = pd.read_excel(mapping_file, engine='openpyxl')  # Specify the engine
    
    # Create a dictionary mapping old names to new names
    mapping_dict = dict(zip(mapping_df['Old Name'], mapping_df['New Name']))
    
    # Replace old names with new names
    for index, row in df.iterrows():
        old_name = row['Name']
        if old_name in mapping_dict:
            df.at[index, 'Name'] = mapping_dict[old_name]
    
    # Write the updated data to the output Excel file
    df.to_excel(output_file, index=False)

# Example usage
input_file = "C:/Users/dhana/OneDrive/Desktop/input.xlsx"
mapping_file = "C:/Users/dhana/OneDrive/Desktop/mapping.xlsx"
output_file = "C:/Users/dhana/OneDrive/Desktop/output.xlsx"
bulk_replace_names(input_file, mapping_file, output_file)
