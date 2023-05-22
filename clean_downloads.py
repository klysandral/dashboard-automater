import os

#Clean up Downloads Folder
folder = "C:\\Users\\Warehouse-MGR\\Downloads"
extension = ".csv"

# traverse the folder
for root, dirs, files in os.walk(folder):
    # loop through the files
    for file in files:
        # check if the file extension matches
        if file.endswith(extension) and root == folder:
            # delete the file
            os.remove(os.path.join(root, file))
