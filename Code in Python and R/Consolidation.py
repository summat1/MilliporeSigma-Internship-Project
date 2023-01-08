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

# Import Data using current working directory
pwd = "C:/Users/x248189/OneDrive - MerckGroup/Documents"

# Working with QTI, reading in data
path = "//Qualified Tissue Inventory_FROZEN 6.27.22.xlsx"
active_inventory = pd.read_excel(pwd + path, sheet_name = "Active Inventory", header = None)

# Cleaning QTI
columns = list(active_inventory.iloc[0, :])
columns = [column for column in columns if not isNaN(column)]
active_inventory.drop(active_inventory.iloc[:, 12:], axis = 1, inplace = True)
active_inventory = active_inventory.drop([0])
active_inventory.columns = columns
active_inventory = active_inventory[active_inventory["Block ID #"].notna()]
active_inventory.rename(columns = {"Lymph node":"Sample Type"}, inplace = True)
active_inventory["Sample Type"] = "FFPE Block"
active_inventory.reset_index(drop = True, inplace = True)

i = 0

# Creating a sourcing data spreadsheet
# By joining together cleaned closet and customer blocks
closet_and_customer = pd.DataFrame()
path = "//Output Files/Cleaned Closet Blocks.xlsx"
closet_blocks = pd.read_excel(pwd + path)
path = "//Output Files/Cleaned Customer Blocks.xlsx"
customer_blocks = pd.read_excel(pwd + path)
all_sourcing_data = pd.concat([customer_blocks, closet_blocks], ignore_index = True)
all_sourcing_data["Block ID #"] = all_sourcing_data["Block ID #"].str.strip()
active_inventory["Block ID #"] = active_inventory["Block ID #"].str.strip()

# Exporting source data and active inventory
all_sourcing_data.to_excel(pwd + "//Output Files/All Sourcing Data - Qualified Only.xlsx", index = False)
active_inventory.to_excel(pwd + "//Output Files/Cleaned Qualified Tissue Inventory.xlsx", index = False)

print("Export Successful.")