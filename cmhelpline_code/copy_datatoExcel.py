import pandas as pd
import os
import openpyxl


def cmhelpline_copy_data_L1():
    """Convert Level 1 (L1) CM Helpline data from .xls to .xlsx and copy into the L-1 workbook.

    Reads 'cmhelpline_1.xls', converts it to 'converted_cmhelpline_1.xlsx' using pandas,
    clears any existing data in the 'data_cmhl' worksheet of 'L-1.xlsx' (preserving references),
    and copies all records starting at cell A1.

    Returns:
        None
    """
    # Base directory relative to this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    cmhelpline_dir = os.path.join(project_dir, 'cmhelpline')

    old_file_path = os.path.join(cmhelpline_dir, 'cmhelpline_1.xls')
    new_file_path = os.path.join(cmhelpline_dir, 'converted_cmhelpline_1.xlsx')

    if not os.path.exists(old_file_path):
        print(f"Error: Could not find the file at {old_file_path}")
        return

    # 1. Read the old .xls file and convert to .xlsx
    df = pd.read_excel(old_file_path)
    df.to_excel(new_file_path, index=False)
    # os.remove(old_file_path)  # Remove the old .xls file after conversion

    # 2. Load the workbooks
    source_wb = openpyxl.load_workbook(new_file_path)
    dest_wb = openpyxl.load_workbook(os.path.join(cmhelpline_dir, 'L-1.xlsx'))

    # 3. Select the source worksheet
    source_sheet = source_wb.active

    # 4. CLEAR THE DATA ONLY (Keep the sheet intact to preserve references)
    if 'data_cmhl' in dest_wb.sheetnames:
        dest_sheet = dest_wb['data_cmhl']
        # This deletes all data from row 1 down to the very last row used
        if dest_sheet.max_row > 0:
            dest_sheet.delete_rows(1, dest_sheet.max_row)
    else:
        dest_sheet = dest_wb.create_sheet('data_cmhl')

    # 5. Iterate through the source data and copy it starting at A1
    for row_idx, row_data in enumerate(source_sheet.iter_rows(values_only=True), start=1):
        for col_idx, cell_value in enumerate(row_data, start=1):
            dest_sheet.cell(row=row_idx, column=col_idx, value=cell_value)

    # 6. Save the changes
    dest_wb.save(os.path.join(cmhelpline_dir, 'L-1.xlsx'))

    # 7. Close workbooks to free memory
    source_wb.close()
    dest_wb.close()
    # os.remove(new_file_path)  # Remove the temporary .xlsx file after copying

    print("Old data cleared safely, and new data successfully copied!")



def download_Final_cmhelpline():
    """Concatenate L2, L3, and L4 CM Helpline Excel files and copy into the high-level workbook.

    Reads the 'Data' sheet from 'cmhelpline_L2.xls', 'cmhelpline_L3.xls', and
    'cmhelpline_L4.xls', concatenates them into 'concatenated_high_level.xlsx',
    clears any existing data in the 'data_cmhl' worksheet of 'high_level.xlsx'
    (preserving references), and copies the records starting at row 3, column 1.

    Returns:
        None
    """
    # Base directory relative to this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    cmhelpline_dir = os.path.join(project_dir, 'cmhelpline')

    concatenconcatenated_file = os.path.join(cmhelpline_dir, 'concatenated_high_level.xlsx')
    high_level_excel = os.path.join(cmhelpline_dir, 'high_level.xlsx')

    # Check if the file exists before trying to remove it
    if os.path.exists(concatenconcatenated_file):
        os.remove(concatenconcatenated_file)
   
    dfs = []
    for f in ['cmhelpline_L2.xls', 'cmhelpline_L3.xls', 'cmhelpline_L4.xls']:
        fp = os.path.join(cmhelpline_dir, f)
        if os.path.exists(fp):
            try:
                dfs.append(pd.read_excel(fp, sheet_name='Data'))
            except Exception:
                try:
                    dfs.append(pd.read_excel(fp))
                except Exception:
                    pass
    if dfs:
        df = pd.concat(dfs, ignore_index=True)
    else:
        df = pd.DataFrame()
    df.to_excel(concatenconcatenated_file, index=False)
    # os.remove(os.path.join(cmhelpline_dir, 'cmhelpline_L2.xls'))
    # os.remove(os.path.join(cmhelpline_dir, 'cmhelpline_L3.xls'))
    # os.remove(os.path.join(cmhelpline_dir, 'cmhelpline_L4.xls'))

    source_wb = openpyxl.load_workbook(concatenconcatenated_file)
    dest_wb = openpyxl.load_workbook(high_level_excel)
    source_sheet = source_wb.active 
    
    # 4. CLEAR THE DATA ONLY (Keep the sheet intact to preserve references)
    if 'data_cmhl' in dest_wb.sheetnames:
        dest_sheet = dest_wb['data_cmhl']
        # This deletes all data from row 1 down to the very last row used
        if dest_sheet.max_row > 0:
            dest_sheet.delete_rows(1, dest_sheet.max_row)
    else:
        dest_sheet = dest_wb.create_sheet('data_cmhl')
    
    # 5. Iterate through the source data and copy it starting at A1
    for row_idx, row_data in enumerate(source_sheet.iter_rows(values_only=True), start=3):
        for col_idx, cell_value in enumerate(row_data, start=1):
            dest_sheet.cell(row=row_idx, column=col_idx, value=cell_value)
    
    # 6. Close source workbook before overwriting
    source_wb.close()

    # 7. Save the changes and close
    dest_wb.save(high_level_excel)
    dest_wb.close()
    # os.remove(concatenconcatenated_file)  # Remove the temporary concatenated file after copying
        
    print("Old data cleared safely, and new data successfully copied!")

if __name__ == '__main__':
    download_Final_cmhelpline()
    cmhelpline_copy_data_L1()
