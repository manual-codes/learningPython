import os
import shutil

# Step 1: Navigate to the path
source_path = r'C:\F12345_My_Project'

# Step 2: Make a copy of the parent folder and save it to the desktop
desktop_path = r'D:\Users\Manuel\Desktop'
parent_folder = os.path.basename(source_path)
destination_path = os.path.join(desktop_path, parent_folder)

try:
    shutil.copytree(source_path, destination_path)
    print(f'Copied to: {destination_path}')
except Exception as e:
    print(f'Error copying: {str(e)}')

# Step 3: Rename all files in the copied folder
for root, dirs, files in os.walk(destination_path):
    for filename in files:
        original_path = os.path.join(root, filename)
        new_filename = f"{parent_folder}_{filename}"
        new_path = os.path.join(root, new_filename)
        try:
            os.rename(original_path, new_path)
            print(f'Renamed: {original_path} -> {new_path}')
        except Exception as e:
            print(f'Error renaming: {str(e)}')
