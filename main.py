from dataclasses import asdict
from tqdm import tqdm
import xlwings as xw
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import colorama
from colorama import Fore, Back, Style
from weeklykpi import *

colorama.init(autoreset=True)

print(Back.BLUE + "Welcome to Performance Data Crafter Automation v2!")
print(Back.BLUE + "***************by RWEPM-Mindanao***************\n")

# ----- Get the raw files ------------
folder_path = input("Enter the folder path of your Raw files: ") # ----- Ask to input folder path
week = input("Enter the Week Number: ") # ----- Ask to input folder path
# folder_path = r"D:\1Performance\per WK util\Python Testing"
abs_folder_path = os.path.abspath(folder_path)

# --------- Run weekly KPI ---------------
weekly_kpi(abs_folder_path, week)
# ------------------------------------------------------------------

print("\nInitializing PCA........")
folder_path = os.path.join(folder_path, 'Consolidated')
abs_folder_path = os.path.abspath(folder_path)
files = os.listdir(abs_folder_path) # ----- list of file names


# ----- Set Destination of crafted file ------------
crafted_folder_path = os.path.join(folder_path,'crafted')
if not os.path.exists(crafted_folder_path):
    os.makedirs(crafted_folder_path)
#  ------------------------------------------------------


# ------ Open Reference Excel file -----------
reference_path = "Reference.xlsx" # ------ Reference file should be located at the script directory
wb_peak_bh_period = xw.Book(reference_path)

ws_peak_average = wb_peak_bh_period.sheets["peak BH Period"]
ws_peak_bh_period = wb_peak_bh_period.sheets["RawData"]

app = xw.apps.active
# --------------------------------------------

def get_average_util_per_trunk(pd_data):
    # ------- Get NE Name
    pd_data['new'] = pd_data.iloc[:, 0].str.extract('^(.*?),.*$')

    # ------- rearange Header position
    ne_name = pd_data.pop('new')
    pd_data.insert(1, 'ne_name',ne_name)

    # ------- Get average per Trunk/Node
    pd_data.iloc[:,:] = pd_data.iloc[:,:].replace(0, np.NaN) # replace all 0 to Nan to exclude it from average
    df_average_per_trunk = pd_data.groupby('ne_name').mean().reset_index()

    return(df_average_per_trunk)


def set_crafted_filename(filename):
    crafted_file_name_path = os.path.join(crafted_folder_path,filename)
    return crafted_file_name_path

def get_raw_data(file_path):
    pd_data = pd.read_excel(file_path)
    return pd_data

# ----- Clear Contents from Reference File (RawData & peak BH Period Sheet)
def clear_reference_file():
    ws_peak_bh_period.range('A1').expand().clear()
    ws_peak_average.range('A2').expand().clear()    
    
# ----- Get max row from Raw data file
def get_max_row(df_raw_data):
    max_row = df_raw_data.shape[0] + 1
    return max_row


# ----- Identifier-----
#  0 ---- Volume
#  1 ---- KPI1 %util
#  2 ---- KPI2
def get_kpi(df_raw_data,crafted_file_name_path,identifier):
    # ----- set variable to store the daily average util to dataframe
    ave_array = {}
    bh_period_array = {}
    df_daily_ave_util = pd.DataFrame(ave_array)
    df_bh_period = pd.DataFrame(bh_period_array)


    # ---- Get the Average per Trunk if KPI1 ----
    pbar.update(5)
    if identifier == 0 or identifier == 1:
      # ---- check if raw file if it is per port
      if (",H" in df_raw_data.iloc[2, 0]) or (",Frame:" in df_raw_data.iloc[2, 0]):
        pbar.set_description('Getting the Average Util per Trunk')
        df_raw_data = get_average_util_per_trunk(df_raw_data)

    # -------- get Num of rows from Raw data (ETH_VOL) -----
    max_row = get_max_row(df_raw_data)
    
    for x in range(25, 170, 24):
      pbar.update(1)
      # ----- get Date
      date = df_raw_data.iloc[0:0,x-1].name
      date = datetime.fromisoformat(date).strftime('%Y-%m-%d')

      pbar.set_description('Computing Hourly Util')
      pbar.update(5)
      if x == 25:
        # ----- Get hourly util
        # ----- iloc ((startRows, endRows), (startCol, endCol))
        hourly_util = df_raw_data.iloc[:, 0:25]
        ws_peak_bh_period.range(1,1).options(index=False).value = hourly_util
        # ws_peak_bh_period.range(1,1).value = hourly_util.set_index('ne_name')

        # ----- Copy Node names to peak BH period Computation sheet (Reference excel)
        ws_peak_average.range((2,1),(max_row,1)).value = '=RawData!A2'

        # ----- set Average Util per BH period
        ws_peak_average.range((2,2),(max_row,22)).value = '=IFERROR(AVERAGE(RawData!B2:E2),"NaN")'
        # ----- Set Peak BH
        ws_peak_average.range((2,23),(max_row,23)).value = "=INDEX($B$1:$V$1,MATCH(X2,B2:V2,0))"
        # ----- Set Peak Average
        ws_peak_average.range((2,24),(max_row,24)).value = '=MAX(B2:V2)'

      else:
        # ----- Get hourly util
        # ----- iloc ((startRows, endRows), (startCol, endCol))
        hourly_util = df_raw_data.iloc[:, x-24:x]
        ws_peak_bh_period.range(1,2).options(index=False).value = hourly_util

      # ----- Get the Average util based from peak BH period
      pbar.set_description('Getting the Average util')
      pbar.update(2)
      daily_ave_util = ws_peak_average.range('A1').expand().options(pd.DataFrame).value['Peak Average']
      df_daily_ave_util[date] = daily_ave_util

      # --- Get BH period
      df_bh_period[date] = ws_peak_average.range('A1').expand().options(pd.DataFrame).value['Peak BH']

    # ---- Get the most frequent BH period per node
    df_bh_period["Average BH Period"] = df_bh_period.mode(axis=1)[0]
    # ---- insert BH Period header
    df_bh_period.insert(0,'Peak BH Period','')

    # ----- Convert to Percentage if % Uil ---- 
    pbar.set_description('Converting Data to Percent')
    pbar.update(5)
    if identifier == 1 or identifier == 2:
      # df_daily_ave_util.iloc[:, 0:7] = df_daily_ave_util.iloc[:, 0:7].astype(str) + '%'
      df_daily_ave_util.iloc[:, 0:7] = (df_daily_ave_util.iloc[:, 0:7] / 100).round(4)

    # ----- Get average weekly util
    pbar.set_description('Getting the Average Weekly util')
    pbar.update(5)
    df_daily_ave_util.iloc[:,0:7] = df_daily_ave_util.iloc[:,0:7].replace(0, np.NaN) # replace all 0 to Nan to exclude it from average
    df_daily_ave_util['average'] = df_daily_ave_util.mean(axis='columns')
    

    # ----- Assign KPI
    pbar.set_description('Assigning KPI')
    pbar.update(10)
    df_daily_ave_util.iloc[:,0:8] = df_daily_ave_util.iloc[:,0:8].replace(np.NaN, 0) # Replace first the  blanks/Nan to 0 to form a float value
    
    conditions = [
      (df_daily_ave_util['average'] == ""),
      (df_daily_ave_util['average'] > 0.9),
      (df_daily_ave_util['average'] > 0.7) & (df_daily_ave_util['average'] <= 0.9),
      (df_daily_ave_util['average'] <= 0.7)
    ]

    kpi_results = ['Passed', 'Failed', 'Fair', 'Passed']
    df_daily_ave_util['kpi'] = np.select(conditions, kpi_results)

    # ----- Get GDS DB Format (Transpose) ----------
    df_daily_ave_util.iloc[:,0:8] = df_daily_ave_util.iloc[:,0:8].replace(np.NaN, 0)
    pbar.set_description('Converting to GDS Database Format')
    pbar.update(10)
    gds_db_format = pd.melt(df_daily_ave_util.reset_index(),
                      id_vars=['ne_name','average','kpi'],
                      value_vars= df_daily_ave_util.iloc[:,0:7],
                      var_name= 'date',
                      value_name='util')

    gds_db_format = gds_db_format[['ne_name', 'util', 'date', 'average', 'kpi']]

    # --------- average Daily final result ---------------
    # print(df_daily_ave_util) 
    # --------- GDS DB Format -----------
    # print(gds_db_format)  

    # print(gds_db_format.info())
    # print(df_daily_ave_util.info())

    # --- combine Util and Peak BH period
    final = pd.concat([df_daily_ave_util, df_bh_period], axis=1)
    
    # ----- Save to Excel
    writer = pd.ExcelWriter(crafted_file_name_path, engine='xlsxwriter')
    final.to_excel(writer, sheet_name='Average Daily Util')

    # ----- Exclude KPI2 from making GDS format (need to separate for uploading) --------
    if identifier != 2:
      gds_db_format.to_excel(writer, sheet_name='GDS DB Format', index=False)

    writer.save()

    # ----- Make separate CSV file for GDS Format for KPI2 --------
    if identifier == 2:
      gds_db_format.to_csv(crafted_file_name_path.replace("xlsx", "csv"), encoding='utf-8', index=False)

    pbar.update(8)
    pbar.set_description('Saving Excel File')

    return

# ---- Identifier
# 1 - SNR
# 2 - MaxRate
def get_snr_maxrate(df_raw_data,crafted_file_name_path,identifier):
    # ------- Get NE Name
    df_raw_data['new'] = df_raw_data.iloc[:, 0].str.extract('^(.*?),.*$')

    # ------- re-arange Header position
    ne_name = df_raw_data.pop('new')
    df_raw_data.insert(0, 'ne_name',ne_name)

    # ----- Get average weekly util (round up)
    pbar.set_description('Getting the Average')
    pbar.update(20)
    df_raw_data['average'] = df_raw_data.iloc[:,2:9].mean(axis='columns').apply(np.ceil)
    

    # ----- Assign KPI
    pbar.set_description('Asigning KPI')
    pbar.update(20)
    # df_raw_data.iloc[:,0:10] = df_raw_data.iloc[:,0:10].replace(np.NaN, 0) # Replace first the  blanks/Nan to 0 to form a float value

    if identifier == 1:
      conditions = [
        (df_raw_data['average'] == ""),
        (df_raw_data['average'] < 7),
        (df_raw_data['average'] >= 7) & (df_raw_data['average'] <= 9),
        (df_raw_data['average'] >= 10)
      ]
    elif identifier == 2:
      conditions = [
        (df_raw_data['average'] == ""),
        (df_raw_data['average'] < 20),
        (df_raw_data['average'] >= 20) & (df_raw_data['average'] <= 49),
        (df_raw_data['average'] >= 50)
      ]

    kpi_results = ['No Reading', 'Failed', 'Fair', 'Passed']
    df_raw_data['kpi'] = np.select(conditions, kpi_results)

    # ------ replace '0' to No Reading
    df_raw_data['kpi'] = df_raw_data['kpi'].replace("0", 'No Reading')

    # rename Port name
    df_raw_data = df_raw_data.rename(columns={df_raw_data.columns[1]: 'port_name'})

    # ----- Get GDS DB Format (Transpose) ----------
    pbar.set_description('Converting to GDS Database Format')
    pbar.update(20)
    gds_db_format = pd.melt(df_raw_data.reset_index(),
                      id_vars=['ne_name','port_name','average','kpi'],
                      value_vars= df_raw_data.iloc[:,2:9],
                      var_name= 'date',
                      value_name='util')

    gds_db_format = gds_db_format[['ne_name','port_name', 'util', 'date', 'average', 'kpi']]
    gds_db_format.iloc[:,2:5] = gds_db_format.iloc[:,2:5].replace(np.NaN, 0)

    # ----- Save to Excel
    writer = pd.ExcelWriter(crafted_file_name_path, engine='xlsxwriter')
    df_raw_data.to_excel(writer, sheet_name='Average', index=False)
    writer.save()

    # ----- Save to CSV for GDS database --------
    gds_db_format.to_csv(crafted_file_name_path.replace("xlsx", "csv"), encoding='utf-8', index=False)

    pbar.update(20)
    pbar.set_description('Saving Excel File')

    return

def get_rx(df_raw_data,crafted_file_name_path):
    pbar.update(60)
    pbar.set_description('Getting the Average')
    # ----- Get average weekly util (round up)
    # df_raw_data.iloc[:,0:7] = df_raw_data.iloc[:,0:7].replace(0, np.NaN)
    df_raw_data['average'] = df_raw_data.iloc[:,1:8].mean(axis='columns')
    df_raw_data['roundup'] = df_raw_data['average'].round(0)

    # ----- Assign KPI
    conditions = [
      (df_raw_data['average'] == ""),
      (df_raw_data['average'] < -29),
      (df_raw_data['average'] < -28 ) & (df_raw_data['average'] >= -29),
      (df_raw_data['average'] >= -28)
    ]

    kpi_results = ['No Reading', 'Failed', 'Fair', 'Passed']
    df_raw_data['kpi'] = np.select(conditions, kpi_results)
    
    # ------ replace '0' to No Reading
    df_raw_data['kpi'] = df_raw_data['kpi'].replace("0", 'NaN')
    df_raw_data['roundup'] = df_raw_data['roundup'].replace(np.NaN, '')

    # ----- rename Port name
    df_raw_data = df_raw_data.rename(columns={df_raw_data.columns[0]: 'onu_name'})

    # ----- add new column for Physical address
    df_raw_data.insert(0,'physical_address','')

    # ----- Save to CSV for GDS database --------
    df_raw_data.to_csv(crafted_file_name_path, encoding='utf-8', index=False, na_rep='NaN')

    pbar.update(20)
    pbar.set_description('Saving Excel File')


def get_throughput(df_raw_data,crafted_file_name_path):
    # ------- Get NE Name
    df_raw_data['new'] = df_raw_data.iloc[:, 0].str.extract('^(.*?),.*$')

    # ------- re-arange Header position
    ne_name = df_raw_data.pop('new')
    df_raw_data.insert(0, 'ne_name',ne_name)

    # ----- Get average weekly util (round up)
    pbar.set_description('Getting the Average')
    pbar.update(20)

    # ---- Replace (-) values to NaN
    df_raw_data.iloc[:,2:9] = df_raw_data.iloc[:,2:9].mask(df_raw_data.iloc[:,2:9]< 0)


    df_raw_data['average'] = df_raw_data.iloc[:,2:9].mean(axis='columns').round(4)

    # rename Port name
    df_raw_data = df_raw_data.rename(columns={df_raw_data.columns[1]: 'onu_name'})

    # ----- Get GDS DB Format (Transpose) ----------
    pbar.set_description('Converting to GDS Database Format')
    pbar.update(20)
    gds_db_format = pd.melt(df_raw_data.reset_index(),
                      id_vars=['ne_name','onu_name','average'],
                      value_vars= df_raw_data.iloc[:,2:9],
                      var_name= 'date',
                      value_name='util')

    gds_db_format = gds_db_format[['ne_name','onu_name', 'util', 'date', 'average']]
    gds_db_format.iloc[:,2:5] = gds_db_format.iloc[:,2:5].replace(np.NaN, 0)

    # ----- Save to Excel
    writer = pd.ExcelWriter(crafted_file_name_path, engine='xlsxwriter')
    df_raw_data.to_excel(writer, sheet_name='Average', index=False)
    writer.save()

    # ----- Save to CSV for GDS database --------
    gds_db_format.to_csv(crafted_file_name_path.replace("xlsx", "csv"), encoding='utf-8', index=False)

    pbar.update(20)
    pbar.set_description('Saving Excel File')

    return



# ------- Main ----------
print(Fore.BLUE + "\nGetting Raw Files........")
for file_name in files:
    if "ETH_VOL" in file_name:
        clear_reference_file()

        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting Volume Util (DS)........")
          f_name = "KPI1_DS_VOL_"
        else:
          print(Fore.GREEN + "\nStarted Crafting Volume Util (US)........")
          f_name = "KPI1_US_VOL_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        
        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(1)

        # -------- Get Raw Volume file store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 0
        get_kpi(df_raw,crafted_file_name_path,identifier)
        pbar.close()
    
    elif "ETH_UTIL_HR" in file_name:
        clear_reference_file()

        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting KPI1 (DS)........")
          f_name = "KPI1_DS_"
        else:
          print(Fore.GREEN + "\nStarted Crafting KPI1 (US)........")
          f_name = "KPI1_US_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(1)

        # -------- Get Raw file and store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 1
        get_kpi(df_raw,crafted_file_name_path,identifier)
        pbar.close()

    elif "PON_UTIL" in file_name:
        clear_reference_file()

        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting KPI2 (DS)........")
          f_name = "KPI2_DS_"
        else:
          print(Fore.GREEN + "\nStarted Crafting KPI2 (US)........")
          f_name = "KPI2_US_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(1)

        # -------- Get Raw file and store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 2
        get_kpi(df_raw,crafted_file_name_path,identifier)
        pbar.close()
    
    elif ("_ETH_UTIL_" in file_name) and  ("_VDSL_" in file_name):
        clear_reference_file()

        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting VDSL UPLINK UTIL (DS)........")
          f_name = "VDSL_KPI1_DS_"
        else:
          print(Fore.GREEN + "\nStarted Crafting VDSL UPLINK UTIL (US)........")
          f_name = "VDSL_KPI1_US_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(1)

        # -------- Get Raw file and store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 1
        get_kpi(df_raw,crafted_file_name_path,identifier)
        pbar.close()

    elif ("VDSL" in file_name) and ("SINR" in file_name):
        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting VDSL SNR (DS)........")
          f_name = "VDSL_DS_SNR_"
        else:
          print(Fore.GREEN + "\nStarted Crafting VDSL SNR (US)........")
          f_name = "VDSL_US_SNR_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        
        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(20)

        # -------- Get Raw Volume file store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 1
        get_snr_maxrate(df_raw,crafted_file_name_path,identifier)
        pbar.close()

    elif ("VDSL" in file_name) and ("_MAX_RATE_" in file_name):
        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DS_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting VDSL Max Rate (DS)........")
          f_name = "VDSL_DS_MAX_RATE_"
        else:
          print(Fore.GREEN + "\nStarted Crafting VDSL Max Rate (US)........")
          f_name = "VDSL_US_MAX_RATE_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        
        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(20)

        # -------- Get Raw Volume file store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        identifier = 2
        get_snr_maxrate(df_raw,crafted_file_name_path,identifier)
        pbar.close()

    elif ("_RX_POWER_" in file_name):

        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_RX_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting RX Power........")
          f_name = "KPI3_RX_"
        else:
          print(Fore.GREEN + "\nStarted Crafting TX Power (DS)........")
          f_name = "KPI3_TX_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".csv")

        
        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(20)

        # -------- Get Raw Volume file store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        get_rx(df_raw,crafted_file_name_path)
        pbar.close()

    
    elif ("SPEED" in file_name):
        abs_file_path = os.path.join(abs_folder_path, file_name) # ---- Get the full file path
        wk_number = file_name[file_name.find("_W")+len("_W")-1:file_name.rfind(".")] # -------- Get Week Number

        # ------- Set crafted file name
        if "_DOWN_" in file_name:
          print(Fore.GREEN + "\nStarted Crafting DOWN Speed........")
          f_name = "DOWN_SPEED_"
        else:
          print(Fore.GREEN + "\nStarted Crafting UP Speed........")
          f_name = "UP_SPEED_"
        crafted_file_name_path = set_crafted_filename( f_name + wk_number + ".xlsx")

        pbar = tqdm(total = 100)
        pbar.set_description("Capturing Data from Raw file")
        pbar.update(20)

        # -------- Get Raw Volume file store as dataframe ------
        df_raw = get_raw_data(abs_file_path)

        # -------- Compute for KPI
        get_throughput(df_raw,crafted_file_name_path)
        pbar.close()

        
clear_reference_file()

# # ----- close xlwings
wb_peak_bh_period.save()
app.quit()
# wb_peak_bh_period.close()

input(Fore.BLUE + "\nCompleted!!. Press Enter to exit\n")
