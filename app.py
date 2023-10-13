from flask import Flask, render_template, request
import os
import pandas as pd

app = Flask(__name__)

def combine_and_export_sheets(input_folder, output_folder='output_folder'):
    try:
        # Convert backslashes to forward slashes
        input_folder = input_folder.replace("\\", "/")
        output_folder = output_folder.replace("\\", "/")

        # Check if the input directory exists
        if not os.path.exists(input_folder) or not os.path.isdir(input_folder):
            raise ValueError(f'Input directory "{input_folder}" does not exist or is not a directory.')

        # Create the output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)

        combined_data = {}  # Move outside the loop to combine data from all subfolders

        # Loop through each subfolder in the main folder
        for subfolder in os.listdir(input_folder):
            subfolder_path = os.path.join(input_folder, subfolder)

            # Check if it's a directory
            if os.path.isdir(subfolder_path):

                # Loop through each file in the subfolder
                for file in os.listdir(subfolder_path):
                    file_path = os.path.join(subfolder_path, file)

                    # Check if it's an Excel file
                    if os.path.isfile(file_path) and file.endswith('.xls'):
                        # Read all sheets from the Excel file
                        sheets = pd.read_excel(file_path, sheet_name=None)

                        # Combine sheets with the corresponding sheets from other files
                        for sheet_name, sheet_data in sheets.items():
                            # Add columns for file and subfolder names
                            sheet_data['Occupations'] = os.path.splitext(file)[0]
                            sheet_data['Year'] = subfolder

                            # Concatenate the data
                            if sheet_name in combined_data:
                                combined_data[sheet_name] = pd.concat([combined_data[sheet_name], sheet_data], ignore_index=True)
                            else:
                                combined_data[sheet_name] = sheet_data

        # Export combined data to separate Excel files for each sheet
        for sheet_name, sheet_data in combined_data.items():
            sheet_output_path = os.path.join(output_folder, f'{sheet_name}.xlsx')
            sheet_data.to_excel(sheet_output_path, index=False)

        result_message = f'Processed and exported sheets from {len(os.listdir(input_folder))} subfolders.'
        return result_message
    except Exception as e:
        error_message = f'An error occurred: {e}'
        # You might want to log the exception for further investigation
        return error_message

@app.route('/', methods=['GET', 'POST'])
def home():
    result_message = None

    if request.method == 'POST':
        input_directory = request.form['input_directory']
        output_directory = request.form.get('output_directory', 'output_folder')

        # Call the function with the provided directories
        result_message = combine_and_export_sheets(input_directory, output_folder=output_directory)

    return render_template('home.html', result_message=result_message)

if __name__ == '__main__':
    app.run(debug=False)
