# Import Packages
import pandas as pd
import numpy as np
import os
import openpyxl
import re

# Method that checks if a data point is null
# Parameter: string
# Returns: yes if string is null, no if not null
def isNaN(string):
    return string != string

# Method that returns the index of the last occurence of a certain value in a string
# Parameters: string, value to look for
# Returns: the index and the count of that value
def index(string, val):
    count = 0
    dash_index = 0
    for i, n in enumerate(string):
        if n == val:
            count += 1
            dash_index = i
    return dash_index, count

# Takes in 2 characters
# Returns if they are both capital letters (A - Z) or digits (1 - 10)
def are_same_type(char1, char2):
    return ((58 > ord(char1) > 47) and (58 > ord(char2) > 47)) or ((91 > ord(char1) > 64) and (91 > ord(char2) > 64)) 


# Used to expand a Cell Marque ID from abbreviated form. 
# Uses string manipulation
# Inputs one ID as a string and returns a list of the new IDs
# AA-1A-G -> AA-1A, AA-1B, AA-1C...AA-1G
def expand(ID):
    newIDS = np.array([])
    second_dash_index, dash_count = index(ID, "-")  
    if len(ID) > 1 and dash_count > 0:
        first_letter_index = second_dash_index - 1   
        first_letter = ID[first_letter_index]
        first_letter_ASCII = ord(first_letter)  
        if (not isNaN(ID[second_dash_index + 1])):
            final_letter = ID[second_dash_index + 1]
            final_letter_ASCII = ord(final_letter)
            
        if dash_count > 1 and (are_same_type(first_letter, final_letter)):        
            newIDS = np.append(newIDS, ID[0 : second_dash_index])
            for i in range (final_letter_ASCII - first_letter_ASCII):
                newID = ID[0 : first_letter_index] + chr(first_letter_ASCII + 1 + i)
                newIDS = np.append(newIDS, newID)
            dash_count = 0
        else: 
            newIDS = np.append(newIDS, ID)
    #print(newIDS)
    return newIDS
    
# Splits up lists of Block ID on spaces, commas, amprisands, or semicolons
# Catches a few cases that are formatted differently (with spaces in their abbreviated form)
# Removes extra spaces 
# Then expands each block ID set by calling the expand method
# Returns the new, expanded IDs for an entire row
def generate_new_IDS(data, old_CM_IDS):
    new_CM_IDS = np.array([])
    if not isNaN(old_CM_IDS):
        if not (old_CM_IDS.startswith("BRN CRB") or old_CM_IDS.startswith("BRN CEBL") or old_CM_IDS.startswith("CIN I") or old_CM_IDS.startswith("CIN II") or old_CM_IDS.startswith("CIN III") or old_CM_IDS.startswith("Umbilical Cord") or old_CM_IDS.startswith("Spermatic Cord")):
            old_CM_IDS = re.split(', | | & |; ', old_CM_IDS)  
        else: 
            old_CM_IDS = re.split(', | & |; ', old_CM_IDS)          
        for ID in old_CM_IDS:
            ID = ID.strip()
            if (ID == ' ') or (ID == '&') or (ID == '') or len(ID) < 2:
                old_CM_IDS.remove(ID)        
        for ID in old_CM_IDS: 
            expanded_IDS = expand(ID)
            new_CM_IDS = np.append(new_CM_IDS, expanded_IDS)
    return new_CM_IDS

# Adds the new rows for expanded block IDs
def add_new_rows_for_IDS(data, i):
    old_CM_IDS = data.loc[i, "Keepers: CM ID#"]
    new_CM_IDS = generate_new_IDS(data, old_CM_IDS)
    new_rows = pd.DataFrame(columns = data.columns)
    if new_CM_IDS.size > 1:
        data.at[i, "Keepers: CM ID#"] = new_CM_IDS[0]
        for j in range(new_CM_IDS.size - 1):
            newID_row = pd.DataFrame(data = data.iloc[[i]], columns = data.columns)
            newID_row.at[i, "Keepers: CM ID#"] = new_CM_IDS[j + 1]
            new_rows = pd.concat([new_rows, newID_row], ignore_index = True)
        data = pd.concat([data, new_rows], ignore_index = True)
    data.reset_index(drop = True, inplace = True)
    return data

# Organizes one Customer Blocks vendor sheet
# Takes in dataframe
# Returns cleaned dataframe and unqualified sheet
def organize_data(data):
    # grabs the source data
    source = data.iloc[0, 0]
    data = data.drop([0])
    # drops NA columns
    data = data.dropna(how='all')
    columns = list(data.iloc[0])
    data.columns = columns
    # adds source as a column
    columns.append("Source")
    data = data.iloc[1:]
    num_rows = data.shape[0]
    # adds source data
    data = data.assign(Source = np.full(num_rows, source))
    data.drop(data.iloc[:, 12:], axis = 1, inplace = True)
    # drops NA column headers
    columns = [column for column in columns if not isNaN(column)]
    for j in columns: 
        columns = [j.strip() for j in columns]
    data.columns = columns
    # resets indexing
    data.reset_index(inplace = True, drop = True)
    i = 0
    # creates unqualified sheet of customer blocks
    unqualified_sheet = pd.DataFrame(columns = columns)
    # Expanding the IDS from current format for each row
    # Also fills in missing data that is meant to be assumed
    # by using previous row information
    while i < num_rows - 1:
        if data.shape[0] > 1:
            if data.iloc[i + 1, 0] == data.iloc[i, 0]:
                if isNaN(data.iloc[i + 1, 1]):
                    data.iloc[i + 1, 1] = data.iloc[i, 1]
                if isNaN(data.iloc[i + 1, 2]):
                    data.iloc[i + 1, 2] = data.iloc[i, 2]
                if isNaN(data.iloc[i + 1, 4]):
                    data.iloc[i + 1, 4] = data.iloc[i, 4]  
        # adds blank Block ID rows to the unqualified sheet
        if isNaN(data.loc[i, "Keepers: CM ID#"]):
            unqualified_sheet = pd.concat([unqualified_sheet, data.iloc[[i]]], ignore_index = True, sort = False)
        # adds new rows
        data = add_new_rows_for_IDS(data, i) 
        i += 1
    # edge case for the last row
    if num_rows > 0:
        data = add_new_rows_for_IDS(data, i)
    # removes the unqualified ones since they are in their own sheet
    data.dropna(subset=['Keepers: CM ID#'], inplace = True)
    data.reset_index(inplace = True, drop = True)
    return data, unqualified_sheet



# Import Data using current working directory
pwd = "C:/Users/x248189/OneDrive - MerckGroup/Documents"
path = "//Frozen Customer Blocks Highlighted Edit.xlsx"

# Load as Openpyxl Workbook
original_data = openpyxl.load_workbook(pwd + path)

# Load first sheet as Pandas DataFrame
customer_blocks = pd.read_excel(pwd + path, sheet_name = "Affiliated Dermatology", header = None)

# Initialize a dictionary of dataframes for each vendor
vendor_dict = {}

# Organize the first sheet of data
customer_blocks, unqualified = organize_data(customer_blocks)
# Drop the data so the customer blocks is initialized only as columns
customer_blocks = customer_blocks.drop([0])
customer_columns = customer_blocks.columns

# Remove the unnecessary sheets
sheets_to_remove = ["Sheet1", "Sheet2", "Sheet3", "Unknown", "Sheet4", "1", "2", "3", "Sheet5", "Tissue Solutions"]
modified_list = original_data.sheetnames
for sheet in sheets_to_remove:
   modified_list.remove(sheet)

unqualified_customer_blocks = pd.DataFrame(columns = customer_columns)
j = 0
# Looping through every sheet
for sheet in modified_list:  
    j+=1
    print(j)
    print(sheet)
    new_data = pd.read_excel(pwd + path, sheet_name = sheet, header = None)
    new_data, unqualified_sheet = organize_data(new_data)
    unqualified_sheet.columns = customer_columns
    unqualified_customer_blocks = pd.concat([unqualified_customer_blocks, unqualified_sheet], ignore_index = True, sort = False)
    new_data.columns = customer_columns
    vendor_name = sheet
    vendor_dict[vendor_name] = new_data
    frame_to_add = vendor_dict[vendor_name]
    customer_blocks = pd.concat([customer_blocks, frame_to_add], ignore_index = True, sort = False)

# Reset indexing after concatenating data
customer_blocks.reset_index(drop = True, inplace = True)
customer_blocks.rename(columns = {'Keepers: CM ID#':'Block ID #'}, inplace = True)

# Export customer blocks
customer_blocks.to_excel(pwd + '//Output Files/Cleaned Customer Blocks.xlsx', index = False)
unqualified_customer_blocks.to_excel(pwd + '//Output Files/Unqualified Customer Blocks.xlsx', index = False)

print("Export Successful.")