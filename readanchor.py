# %%
import pandas as pd
import numpy as np
from pathlib import Path

# %%
class ReadAnchorData:
    @staticmethod
    def get_set_forecast(filepaths : Path):
        for i in filepaths:
            sheetnames = pd.ExcelFile(i).sheet_names
            dfs = ReadAnchorData.readfcst(i, sheetnames)
        return dfs

    def get_type(filepath : Path[0]):
        # Each file name will specify what the data type is, this is to standardise the naming convention as given in the types dict.
        # key = filename, value = standardised name
        types = {
            'Bom_Report' : 'BOM',
            'SetsPerMonthcombined': 'Combined',
            'SetsPerMonthConst': 'Constrained',
            'SetsPerMonthUnConst': 'Unconstrained'
        }
        filepath_str = str(filepath)

        try:
            for key, val in types:
                if key in filepath_str:
                    return val
        except:
            return 'unknown'

    def get_release(filepath : Path):
        # The release is given by the name of the parent folder
        parentfol = str(filepath.parent.absolute())
        return parentfol[len(parentfol)-7:]


    def get_xlsx_files(base_dir):
        # Returns all xlsx files (incl subdirs)
        return Path(base_dir).glob("*/*.xlsx")

    def readfcst(filepath, sheetname):
        # The xlsx files can contain multiple sheets and no std naming convention. This will open each sheet until it finds the matching indicator columns
        # Each forecast will have these indicator columns, but the forecast column headers (202101, 202102 etc) can be different. These are transposed to the YYYYMM column
        for sheet in sheetname:
            df : pd.DataFrame = pd.read_excel(filepath, sheet)
            indicator_cols = ['Total Forecast']
            if indicator_cols[0] in df.columns:
                df = df.drop('Total Forecast', axis=1)
                dates = [i for i in df.columns if '20' in i] # This will get all forecast column names
                dims = [i for i in df.columns if '20' not in i] # Inverse of the above
                df = df.melt(id_vars=dims, value_vars=dates, var_name='YYYYMM', value_name='SetForecast')
                df['release'] = ReadAnchorData.get_release(filepath)
                df['KeyFigure'] = ReadAnchorData.get_type(filepath)
                return df
    
    def read_bom(filepath, sheetname):
        for sheet in sheetname:
            df : pd.DataFrame = pd.read_excel(filepath, sheet)
            indicator_cols = ['BOM']
            if indicator_cols[0] in df.columns:
                df = df.drop('Total Forecast', axis=1)
                df['release'] = ReadAnchorData.get_release(filepath)
                return df

# %%
base_dir = r'\\na.jnj.com\dpyusdfsroot\RY_Company\Supply Chain Mgmt\Spine Plan-NPI\Conduit\Jonny\Powerbi\Anchor'
all_files = ReadAnchorData.get_xlsx_files(base_dir)

# %%
ReadAnchorData.get_release(list(all_files)[0])


