# Import file
# Import modules and set parameters
import pandas as pd
import numpy as np
import os


# Set parameters
pd.set_option('display.max_rows', None)

# Use predefined path and file for now but will need to allow for user to select/input this info
folder_path = '/Users/dougroehm/Code/repos/mtdor-extract-cuv/TY23-cap-final-appraisals/'
file_name = 'U202_TY23_CAP_Final_Appraisal.xlsx'
file_path = folder_path + file_name

# Create a Pandas DataFrame from file
# df = pd.ExcelFile(file_path)

# Create a Pandas DataFrame of the "CUV" worksheet
df = pd.read_excel(file_path, sheet_name='CUV', header=None, usecols="A:D", names=["Col1", "Col2", "Col3", "Col4"])

# Print current progress
print(df.info())

