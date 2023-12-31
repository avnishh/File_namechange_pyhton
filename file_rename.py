import os
import pandas as pd

# Define the folder where your files are located and the Excel file path.
#pip3 install pandas openpyxl
#folder_raw_path= r"c:\Users\admin\Desktop\iai\data"
current_directory = os.getcwd()
folder_path_raw = f"{current_directory}\data"
#print(folder_path_raw)
excel_file_path_raw = f"{current_directory}\data\data.xlsx"
file_extension=r".docx"

# Read the Excel file into a DataFrame.
df = pd.read_excel(excel_file_path_raw)

# Iterate through the DataFrame and rename files.
for index, row in df.iterrows():
    current_nam = row['Current Name']
    current_name=f"{current_nam}{file_extension}"
    new_nam = row['New Name']
    new_name=f"{new_nam}{file_extension}"
    

    current_file_path = os.path.join(folder_path_raw, current_name)
    new_file_path = os.path.join(folder_path_raw, new_name)

    try:
        os.rename(current_file_path, new_file_path)
        print(f"Renamed '{current_name}' to '{new_name}'.")
    except FileNotFoundError:
        print(f"File '{current_name}' not found in the folder.")
    except FileExistsError:
        print(f"File '{new_name}' already exists in the folder.")

print("File renaming complete.")
