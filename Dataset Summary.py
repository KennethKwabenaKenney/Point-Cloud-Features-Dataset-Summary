import os
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog

# Initialize the QApplication
app = QApplication(sys.argv)

# Open a file dialog to select the folder containing CSV files
options = QFileDialog.Options()
options |= QFileDialog.ReadOnly
folder_path = QFileDialog.getExistingDirectory(None, "Select Folder Containing CSV Files", options=options)

if folder_path:
    # Initialize an empty list to hold data for all files
    data = []

    # Iterate through all files in the selected folder
    for file_name in sorted(os.listdir(folder_path)):
        if file_name.endswith('.csv'):
            file_path = os.path.join(folder_path, file_name)
            
            # Extract Scene ID from the file name (number before the first "_")
            scene_id = file_name.split('_')[0]

            # Read the CSV file into a DataFrame
            df = pd.read_csv(file_path)

            # Count the occurrences of each unique object name in the "Ext_Class_Label" column
            object_counts = df['Ext_Class_Label'].value_counts()

            # Sum counts for "Unclassified" and "Unknown"
            no_of_unclassified = object_counts.get('Unclassified', 0) + object_counts.get('Unknown', 0)

            # Sum counts for specific object names
            pole_like_objects = ['Pole-like object', 'Street lamp', 'Traffic sign', 'Traffic light']
            no_of_pole_like_objects = sum(object_counts.get(obj, 0) for obj in pole_like_objects)

            # Sum counts for all classified objects excluding "Unclassified" and "Unknown"
            no_of_classified = object_counts.sum() - no_of_unclassified

            # Prepare data for this file
            file_data = {
                'Scene ID': scene_id,
                'Point Density': df['Total Points'].sum(),
                'No. of Pole-like Objects': no_of_pole_like_objects,
                'No. of Classified': no_of_classified,
                'No. of Unclassified': no_of_unclassified
            }
            data.append(file_data)

    # Create a DataFrame from the collected data
    result_df = pd.DataFrame(data)

    # Save the results to an Excel file
    excel_path = os.path.join(folder_path, 'Dataset_Summary.xlsx')
    result_df.to_excel(excel_path, index=False, sheet_name='Summary')

    print(f"Data has been saved to {excel_path}")
else:
    print("No folder selected.")

# Exit the application
app.quit()
sys.exit()
