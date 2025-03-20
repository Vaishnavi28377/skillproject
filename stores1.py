import pandas as pd

# Replace 'your_excel_file.xlsx' with the actual path to your Excel file
file_path = "D:\storesdataset.xlsx"

# Read the data from the Excel file into a Pandas DataFrame
try:
    df = pd.read_excel("D:\storesdataset.xlsx")  # Assuming data is in the first sheet
    print("Data loaded successfully!")
except FileNotFoundError:
    print(f"Error: File not found at {"D:\storesdataset.xlsx"}")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    exit()

# Display the first few rows of the DataFrame to verify the data
print(df.head())