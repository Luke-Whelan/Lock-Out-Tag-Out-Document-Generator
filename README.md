# Lock-Out-Tag-Out-Document-Generator
VBA code to create lock-out-tag-out documents



This set of VBA macros takes information from the associated spreadsheet (Equipment ID, Description of Equipment, and Location) and uses it to generate  a custom Lock-Out-Tag-Out document for each selected equipment ID. Sample images of each isolator for each energy source have been provided and saved in each respective folder. These are copied into their respective files and formatted to the correct size and position within the document. 



Before running the scripts, check that the appropriate Microsoft Object Library is enabled (this enables Excel to interact with Word):
 - Open LOTO Tracker.xlsm
 - Open the VBA editor (press Alt + F11)
 - Tools > References > Microsoft Word 16.0 Object Library (set the box to checked)
 - Click OK

In LockOutTagOutDocGenerator() and templateFinder(), modify the file path to work with where you have saved this folder:
- Search for "enter your path here" to find the relevant line

Run the script to generate LOTO documents:
- Close the VBA editor
- In the Excel file, select cells A2:A4 (or any of the equipment IDs, such as A3:A4 etc.)
- View the macros in the worksheet (Alt + F8)
- Double-click LockOutTagOutDocGenerator
- The script will generate LOTO documents for the selected documents
- Run the script again for Groups 2 and 3 as desired (see other tabs)

To see which template version each piece of equipment has used, run templateFinder():
- In the Excel file, select cells A2:A4
- View the macros in the worksheet (Alt + F8)
- Double-click templateFinder

To reset the test:
- You need to replace the folders for Group 1, 2 and 3
- A fresh set of these folders can be found in "Backup of test folders"
- If you do not replace them, the script will not run as the files it is trying to create already exist

To view the scripts in the Word files:
- Open "Templates" folder
- Open any of the Word files in here
- View > Macros > View Macros > Edit (bordersAndResize should be selected)






:) 
