# Import Packages
import pandas as pd
import numpy as np
import os
import openpyxl
import re

# Reading in Sourcing Data and Active Inventory
pwd = "C:/Users/x248189/OneDrive - MerckGroup/Documents"

source_data = pd.read_excel(pwd + "//Output Files/All Sourcing Data - Qualified Only.xlsx")
active_inventory = pd.read_excel(pwd + "//Output Files/Cleaned Qualified Tissue Inventory.xlsx")

# Merging Sourcing Data with Active Inventory
source_data["Block ID #"] = source_data["Block ID #"].fillna("NA")

merged_QTI_output_1 = pd.merge(active_inventory, source_data, how = "inner", on = 'Block ID #')
merged_QTI_output_1.to_excel(pwd + "//Output Files/First Merge.xlsx", index = False)

# Renaming to merge with "Previous Block ID" column as well
source_data.rename(columns = {'Block ID #': 'Previous Block ID'}, inplace = True)

merged_QTI_output_2 = pd.merge(active_inventory, source_data, how = "inner", on = "Previous Block ID")
merged_QTI_output_2.to_excel(pwd + "//Output Files/Second Merge.xlsx", index = False)

# Concatentating the 2 merges into the final output
final_output = pd.concat([merged_QTI_output_1, merged_QTI_output_2], ignore_index = True)
final_output.sort_values(by=['Block ID #'], inplace=True)
final_output.to_excel(pwd + "//Output Files/QTI + Source Data Merged.xlsx", index = False)

print("Export Successful.")

# Exporting files
#unknown_source = source_data[source_data["Source"] == "UNKNOWN"]
#unknown_sources = pd.concat([unknown_source, ml_source], ignore_index = True)
#unqualified_unknown_sources = unqualified_closet_blocks[unqualified_closet_blocks["Surgical Number"].notna()]
#unknown_sources = pd.concat([unknown_sources, unqualified_unknown_sources], ignore_index = True)
#unknown_sources.rename(columns = {'Previous Block ID' : 'Block ID #'}, inplace = True)
#unknown_sources = unknown_sources.dropna(axis='columns', how = 'all')
#unknown_sources.to_excel(pwd + "//Output Files/Unknown_Sources.xlsx", index = False)