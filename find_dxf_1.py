import os
import shutil
import pandas as pd
from tkinter import filedialog
from tkinter import Tk, Button, Label, Entry, StringVar



def browse_directory():
    """Opens a file dialog and updates directory_entry with the chosen path."""
    filename = filedialog.askdirectory()
    directory_var.set(filename)

def browse_new_directory():
    """Opens a file dialog and updates new_dir_entry with the chosen path."""
    filename = filedialog.askdirectory()
    new_dir_var.set(filename)

def browse_file():
    """Opens a file dialog and updates file_entry with the chosen path."""
    filename = filedialog.askopenfilename(filetypes=(("Excel files", "*.xls"), ("all files", "*.*")))
    file_var.set(filename)

def process_files():
    """Performs the .dxf file copying based on the chosen directories and Excel file."""
    directory_to_search = directory_var.get()
    excel_file_path = file_var.get()
    new_folder_path = new_dir_var.get()

    # Create the new directory if it doesn't exist
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)

    # Read the Excel file
    df = pd.read_excel(excel_file_path, engine='xlrd')

    # Check if "Reference_Number" or "Reference number" column exists in the DataFrame
    if "Reference_Number" in df.columns:
        column_name = "Reference_Number"
    elif "Reference number" in df.columns:
        column_name = "Reference number"
    elif "Reference #" in df.columns: # added the Reference # column name in case it is needed
        column_name = "Reference #"
    else:
        print("Neither 'Reference #' nor 'Reference number' columns exist in the DataFrame.")
        exit(1)

    # Prepare a list to store names of missing .dxf files
    missing_dxf_files = []

    for file in os.listdir(directory_var.get()):
            try:
                loadFile = open('%s\%s' % (directory_var.get(), file), 'rb')
                print('.', end='')
                loadFile.read(1)
                loadFile.close()
            except:
                print('Error opening file %s' % file)



    # Loop through each row in the specified column
    for ref_number in df[column_name]:
        # Construct the .dxf file name
        dxf_file_name = str(ref_number) + '.dxf'

        # Construct the full path to the .dxf file
        dxf_file_path = os.path.join(directory_to_search, dxf_file_name)

        # Check if the .dxf file exists in the directory
        if os.path.isfile(dxf_file_path):
            # Construct the path to the new .dxf file location
            new_dxf_file_path = os.path.join(new_folder_path, dxf_file_name)

            # Copy the .dxf file to the new location
            shutil.copy(dxf_file_path, new_dxf_file_path)
        else:
            print(f"{dxf_file_name} is missing.")
            missing_dxf_files.append(dxf_file_name)

    # If there are missing files, write them into a text file in the same location as copied .dxf files
    if missing_dxf_files:
        with open(os.path.join(new_folder_path, "missing_dxf_files.txt"), "w") as f:
            for filename in missing_dxf_files:
                f.write(f"{filename}\n")
        print("A list of missing .dxf files has been saved to 'missing_dxf_files.txt' in the copied .dxf files directory.")
    else:
        print("All .dxf files were found and copied successfully.")

    print("Process Completed!")

# Create main Tkinter window
root = Tk()

# Create StringVars to hold directory and file paths
directory_var = StringVar()
file_var = StringVar()
new_dir_var = StringVar()

# Create directory Entry and Browse button
directory_entry = Entry(root, textvariable=directory_var, width=50)
directory_entry.pack()
directory_button = Button(root, text="Browse for .dxf Directory", command=browse_directory)
directory_button.pack()

# Create file Entry and Browse button
file_entry = Entry(root, textvariable=file_var, width=50)
file_entry.pack()
file_button = Button(root, text="Browse for Excel File", command=browse_file)
file_button.pack()

# Create new directory Entry and Browse button
new_dir_entry = Entry(root, textvariable=new_dir_var, width=50)
new_dir_entry.pack()
new_dir_button = Button(root, text="Browse for New Directory", command=browse_new_directory)
new_dir_button.pack()

# Create a button to trigger the file processing
process_button = Button(root, text="Process Files", command=process_files)
process_button.pack()

# Run the Tkinter event loop
root.mainloop()
