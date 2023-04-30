import os
import pandas as pd
from unidecode import unidecode

# set the input and output folder paths
dir_path = "c:/temp/data/SAP/"

def run():
    print('Exporting all xlsx files to csv format.')
    # Loop through all files in the directory
    for file_name in os.listdir(dir_path):
        # Check if the file is an xlsx file
        if file_name.endswith('.xlsx'):
            # Read the xlsx file into a pandas dataframe
            df = pd.read_excel(os.path.join(dir_path, file_name))
            # apply unidecode to all columns
            df = df.applymap(lambda x: unidecode(x) if isinstance(x, str) else x)
            # converting the text to lowercase
            # Define a function to convert strings to lowercase
            def lower_case(x):
                if isinstance(x, str):
                    return x.lower()
                else:
                    return x
            # Apply the function to all elements of the DataFrame
            df = df.applymap(lower_case)
            # Set the path and file name for the csv file
            csv_file_name = os.path.splitext(file_name)[0] + '.csv'
            csv_file_path = os.path.join(dir_path + 'csv/', csv_file_name)
            # Write the dataframe to a csv file
            df.to_csv(csv_file_path, sep=';', index=False)
    print('Exporting to csv complete!')