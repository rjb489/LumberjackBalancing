import pandas as pd

# Request file to be analyzed (temporary method for demo)
  # function: input
file_path = input("Please enter the path to the Excel file: ")

# Use pandas to read through excel file, store data in variable
  # function: pd.read_excel
data = pd.read_excel(file_path)

# Arbitrary filtration for names, used to remove edge cases in final version
  # function: str.match
filtered_names = data[~data['Name'].str.match(r'^(J.*|.*n)$')]

# Select specific values based on arbitrary values for now
filtered_workload = filtered_names[
    (filtered_names['Workload Percentage'] > 40) & 
    (filtered_names['Workload Percentage'] < 90)
].copy()

# Create new row for values that have been selected, filtered, and used for operatoins
filtered_workload.loc[:, 'New Workload'] = (filtered_workload['Workload Percentage'] ** 2) / 100

# Create new excel file with new category, output and display to user
output_file = "filtered_modified_workload.xlsx"
filtered_workload.to_excel(output_file, index=False)

print(f"The modified data has been saved to {output_file}")