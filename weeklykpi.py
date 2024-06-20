import pandas as pd
from pathlib import Path
import os

# folder_path ='D:/1Performance/per WK util/2023/WK16/New Folder/'
# week = '16'



def weekly_kpi(folder_path, week):

    print("\nInitializing........")

    # ----------------FILTER ALL MYCOM FILES (HOURLY)-------------------------

    # Read the Excel file
    df = pd.read_excel('KPIs.xlsx', sheet_name='Hourly')

    # Get unique values from the 'main_file_name' and 'command' column
    report_file = df['main_file_name'].unique().tolist()
    commands = df['command'].unique().tolist()

    # Create an empty dictionary
    my_commands = {}
    my_filenames = {}

    # Iterate over each item in the report_file list
    for item in report_file:
        # Get the commands associated with the current item
        commands_per_item = df.loc[df['main_file_name'] == item, 'command'].tolist()
        # Assign the commands to the dictionary with the current item as the key
        my_commands[item] = commands_per_item

    # Iterate over each command
    for command in commands:
        # Get the corresponding final file name
        final_file_name = df.loc[df['command'] == command, 'final_file_name'].iloc[0]
        # Assign the final file name to the dictionary with the command as the key
        my_filenames[command] = final_file_name

    # ---------- Output of my_commands and my_filenames -----------
    # report_file = ['KPI1_DS_US_RX_UPLINK_HT', 'KPI1_DS_US_RX_UPLINK_FH', 'MIN_DS_ETH_VOL_HT_WK', 'MIN_DS_ETH_VOL_FH_WK', 'MIN_DS_US_PON_UTIL_HT', 'MIN_DS_US_PON_UTIL_FH', 'MIN_VDSL_DS_US_UTIL_HT', 'MIN_CPU_MEM_HR_FH_WK']
    # my_commands = {'KPI1_DS_US_RX_UPLINK_HT':['huaw_gpon_RX_bw_occupancy_EthPort', 'huaw_gpon_TX_bw_occupancy_EthPort','Rx power'], 'KPI1_DS_US_RX_UPLINK_FH':['gpon_Down_Prop_Traff_Util_ULPort_FH', 'gpon_Up_Prop_Traff_Util_ULPort_FH', 'Rx_power.Optical_Power'], 'MIN_DS_ETH_VOL_HT_WK':['huaw_gpon_eport_rx_rate_ne_sum_Gbps'], 'MIN_DS_ETH_VOL_FH_WK':['gpon_DwnStream_Vol_ULPort_FH_Gbps_Wireline_1'], 'MIN_DS_US_PON_UTIL_HT':['huaw_gpon_Dwnstrm_bw_occupancy_UNI Port_NEMaxTimemax', 'huaw_gpon_UNI Upstream Bandwidth Occupancy_UNIPort'],'MIN_DS_US_PON_UTIL_FH':['gpon_Down_Prop_Traff_Util_PONPort_FH','gpon_Up_Prop_Traff_Util_PONPort_FH','Tx_power.Optical_Power'],'MIN_VDSL_DS_US_UTIL_HT':['huaw_gpon_RX_bw_occupancy_EthPort_VDSL','huaw_gpon_TX_bw_occupancy_EthPort_VDSL'],'MIN_CPU_MEM_HR_FH_WK':['MEMORY_USAGE_RATE.CPU_Memory_Usage','CPU_USAGE_RATE.CPU_Memory_Usage']}
    # my_filenames = {'huaw_gpon_RX_bw_occupancy_EthPort':'MIN_DS_ETH_UTIL_HR_HT_WK', 'huaw_gpon_TX_bw_occupancy_EthPort':'MIN_US_ETH_UTIL_HR_HT_WK', 'Rx power':'MIN_RX_POWER_UPLINK_HT_WK','gpon_Down_Prop_Traff_Util_ULPort_FH':'MIN_DS_ETH_UTIL_HR_FH_WK','gpon_Up_Prop_Traff_Util_ULPort_FH':'MIN_US_ETH_UTIL_HR_FH_WK','Rx_power.Optical_Power':'MIN_RX_POWER_UPLINK_FH_WK','huaw_gpon_eport_rx_rate_ne_sum_Gbps':'MIN_DS_ETH_VOL_HT_WK','gpon_DwnStream_Vol_ULPort_FH_Gbps_Wireline_1':'MIN_DS_ETH_VOL_FH_WK','huaw_gpon_Dwnstrm_bw_occupancy_UNI Port_NEMaxTimemax':'MIN_DS_PON_UTIL_HR_HT_WK','huaw_gpon_UNI Upstream Bandwidth Occupancy_UNIPort':'MIN_US_PON_UTIL_HR_HT_WK','gpon_Down_Prop_Traff_Util_PONPort_FH':'MIN_DS_PON_UTIL_HR_FH_WK','gpon_Up_Prop_Traff_Util_PONPort_FH':'MIN_US_PON_UTIL_HR_FH_WK','Tx_power.Optical_Power':'MIN_TX_POWER_PON_FH_WK','huaw_gpon_RX_bw_occupancy_EthPort_VDSL':'MIN_VDSL_DS_ETH_UTIL_NE_HR_HT_WK','huaw_gpon_TX_bw_occupancy_EthPort_VDSL':'MIN_VDSL_US_ETH_UTIL_NE_HR_HT_WK','MEMORY_USAGE_RATE.CPU_Memory_Usage':'MIN_CPU_FH_HR_WK','CPU_USAGE_RATE.CPU_Memory_Usage':'MIN_MEM_FH_HR_WK'}
    # -------------------------------------------------------------

    # ----- Set Destination of crafted file ------------
    raw_mycom_files_folder = os.path.join(folder_path,'Raw Mycom Files')
    if not os.path.exists(raw_mycom_files_folder):
        os.makedirs(raw_mycom_files_folder)

                                          
    for i in list(report_file):
        raw_file_name = str(i)+'.csv'
        # file_path = folder_path + raw_file_name
        file_path = os.path.join(folder_path ,raw_file_name)

        path = Path(file_path)    
        
        if path.is_file():
            data = pd.read_csv(path)

            data.columns.values[0:2] = ["Short name", "command"] # rename first two headers
            
            cols = ['Short name'] # fill in blank elements
            data.loc[:,cols] = data.loc[:,cols].ffill()        

            sub_str = 'MINDANAO,' # find and replace MINDANAO
            data[data['Short name'].str.contains(sub_str)]
            data['Short name'] = data['Short name'].replace(to_replace= sub_str, value = '', regex=True)     

            data = data.astype(str) # convert dataframe to str then replace ' %' and ' Gbps' to ''
            data = data.apply(lambda x: x.str.replace('%',''))
            data = data.apply(lambda x: x.str.replace(' Gbps',''))

            for j in range(len(my_commands[i])): # filter command
                
                command = my_commands[i][j] # which command to filter
                file_tosave = data[(data['command']==command)]
                file_tosave = file_tosave.drop(['command'], axis=1)
                
                file_name = my_filenames[[command][0]] + week
                out_path = os.path.join(raw_mycom_files_folder, str(file_name)+'.csv')
                # out_path = raw_mycom_files_folder + '/' + str(file_name)+'.csv'

                #file_tosave.to_csv(out_path, engine='xlsxwriter', index=False)
                file_tosave.to_csv(out_path, index=False)
                print(file_name+' has been successfully generated!')
                
        else:
            print(f'The file {file_path} does not exist')
            continue
            
    print('\nDone filtering files (Hourly)!\n\n')




    # ----------------FILTER ALL MYCOM FILES (DAILY)-------------------------

    # Read the Excel file
    df = pd.read_excel('KPIs.xlsx', sheet_name='Daily')

    # Get unique values from the 'main_file_name' and 'command' column
    report_file = df['main_file_name'].unique().tolist()
    commands = df['command'].unique().tolist()

    # Create an empty dictionary
    my_commands = {}
    my_filenames = {}

    # Iterate over each item in the report_file list
    for item in report_file:
        # Get the commands associated with the current item
        commands_per_item = df.loc[df['main_file_name'] == item, 'command'].tolist()
        # Assign the commands to the dictionary with the current item as the key
        my_commands[item] = commands_per_item

    # Iterate over each command
    for command in commands:
        # Get the corresponding final file name
        final_file_name = df.loc[df['command'] == command, 'final_file_name'].iloc[0]
        # Assign the final file name to the dictionary with the command as the key
        my_filenames[command] = final_file_name 

    # ---------- Output of my_commands and my_filenames -----------
    # report_file = ['MIN_RX_POWER_FH_WK']
    # my_commands = {'MIN_RX_POWER_FH_WK':['RxPwr_ONUPort_FH','UP_Speed.Traffic_Analysis','Down_Speed.Traffic_Analysis','Tx_power.Optical_Power']}
    # my_filenames = {'RxPwr_ONUPort_FH':'MIN_RX_POWER_FH_WK','UP_Speed.Traffic_Analysis':'MIN_UP_SPEED_FH_WK','Down_Speed.Traffic_Analysis':'MIN_DOWN_SPEED_FH_WK','Tx_power.Optical_Power':'MIN_TX_POWER_FH_WK'}
    # -------------------------------------------------------------

    for i in list(report_file):
        raw_file_name = str(i)+'.csv'
        # file_path = folder_path + raw_file_name
        file_path = os.path.join(folder_path ,raw_file_name)
        path = Path(file_path)    
        
        if path.is_file():
            data = pd.read_csv(path)

            
            data.columns.values[0:2] = ["Short name", "command"] # rename first two headers
            cols = ['Short name'] # fill in blank elements
            data.loc[:,cols] = data.loc[:,cols].ffill()

            #-- convert col to date
            new_col = pd.to_datetime(data.columns[2:9]).strftime("%Y-%m-%d").tolist()
            new_col.insert(0,"")


            sub_str = 'MINDANAO,' # find and replace MINDANAO
            data[data['Short name'].str.contains(sub_str)]
            data['Short name'] = data['Short name'].replace(to_replace= sub_str, value = '', regex=True)
            
            data = data.astype(str) # convert dataframe to str
            data = data.apply(lambda x: x.str.replace(',PON','')) #remove ,PON
            
            for j in range(len(my_commands[i])): # filter command
                
                command = my_commands[i][j] # which command to filter
                file_tosave = data[(data['command']==command)]
                file_tosave = file_tosave.drop(['command'], axis=1)
                
                file_tosave.columns = new_col
                # print(file_tosave.columns)

                if i == "MIN_UP_DOWN_HT_WK":
                    file_tosave = get_ht_nodes_per_region(file_tosave)
                
                file_name = my_filenames[[command][0]] + week
                out_path = os.path.join(raw_mycom_files_folder, str(file_name)+'.csv')
                # out_path = raw_mycom_files_folder + '/' + str(file_name)+'.csv'

                file_tosave.to_csv(out_path, index=False)
                print(file_name+' has been successfully generated!')
                
        else:
            print(f'The file {file_path} does not exist')
            continue


    # for MIN_RX_POWER_HT_WK
    file_path = os.path.join(folder_path, 'MIN_RX_POWER_HT_WK.xlsx')
    # file_path = folder_path + 'MIN_RX_POWER_HT_WK.xlsx'
    path = Path(file_path)

    if path.is_file():
        df = pd.read_excel(file_path)
        
        df.columns = df.iloc[5]
        df = df[6:]
        # df.columns = pd.to_datetime(df.columns).strftime("%Y-%m-%d")

        # col_list = df.columns.values.tolist()
        
        out_path = os.path.join(raw_mycom_files_folder, 'MIN_RX_POWER_HT_WK'+week+'.csv')
        # out_path = raw_mycom_files_folder + '/MIN_RX_POWER_HT_WK'+week+'.csv'

        df.to_csv(out_path, index=False)
        print('MIN_RX_POWER_HT_WK'+week+' has been successfully generated!')
    else:
        print(f'The file {file_path} does not exist')
            
    print('\nDone filtering files (Daily)!\n\n')


    # ---- for MIN_DOWN_SPEED_HT_WK
    file_path = os.path.join(folder_path, 'MIN_DOWN_SPEED_HT_WK.csv')
    path = Path(file_path)

    if path.is_file():
        df = pd.read_csv(file_path)
        df.iloc[:,1:] = df.iloc[:,1:].div(1024,axis=0)
        df.columns.values[0:1] = [""]
        out_path = os.path.join(raw_mycom_files_folder, 'MIN_DOWN_SPEED_HT_WK'+week+'.csv')

        df.to_csv(out_path, index=False)
        
        print('MIN_DOWN_SPEED_HT_WK'+week+' has been successfully generated!')
    else:
        print(f'The file {file_path} does not exist')
            


    # ---- for MIN_UP_SPEED_HT_WK
    file_path = os.path.join(folder_path, 'MIN_UP_SPEED_HT_WK.csv')
    path = Path(file_path)

    if path.is_file():
        df = pd.read_csv(file_path)
        df.iloc[:,1:] = df.iloc[:,1:].div(1024,axis=0)
        df.columns.values[0:1] = [""]
        out_path = os.path.join(raw_mycom_files_folder, 'MIN_UP_SPEED_HT_WK'+week+'.csv')

        df.to_csv(out_path, index=False)

        print('MIN_UP_SPEED_HT_WK'+week+' has been successfully generated!')
    else:
        print(f'The file {file_path} does not exist')
            
    print('\nDone filtering files (Daily)!\n\n')



    #-------------------------- CONSOLIDATE ALL RAW FILES --------------------------

    pair_ref = []
    pair_files = {}

    # Read the Excel file
    df = pd.read_excel('KPIs.xlsx', sheet_name='Consolidate')

    # Get unique values
    pair_ref = df['Fiberhome'].unique().tolist()

    # Add variable string to each item in the list
    pair_ref = [item + week for item in pair_ref]

    # Create a dictionary and pair it with the values from the 'Huawei' column with variable 'week' added
    pair_files = {pair_ref[i]: df.loc[i, 'Huawei'] + week for i in range(len(pair_ref))}

    # Create a dictionary and pair it with the values under the 'Consolidated' column, with variable 'week' added to the keys and values
    my_filenames_final = {pair_ref[i]: df.loc[i, 'Consolidated'] + week for i in range(len(pair_ref))}

    # ---------- Output of pair_ref,  pair_files, and my_filenames_final ---------------------------------------------------------------
    # pair_ref = ['MIN_DS_ETH_UTIL_HR_FH_WK'+week, 'MIN_US_ETH_UTIL_HR_FH_WK'+week, 'MIN_RX_POWER_UPLINK_FH_WK'+week,
    #         'MIN_DS_ETH_VOL_FH_WK'+week, 'MIN_DS_PON_UTIL_HR_FH_WK'+week, 'MIN_US_PON_UTIL_HR_FH_WK'+week, 'MIN_RX_POWER_FH_WK'+week]
    # pair_files = {'MIN_DS_ETH_UTIL_HR_FH_WK'+week:'MIN_DS_ETH_UTIL_HR_HT_WK'+week,
    #             'MIN_US_ETH_UTIL_HR_FH_WK'+week:'MIN_US_ETH_UTIL_HR_HT_WK'+week,
    #             'MIN_RX_POWER_UPLINK_FH_WK'+week:'MIN_RX_POWER_UPLINK_HT_WK'+week,
    #             'MIN_DS_ETH_VOL_FH_WK'+week:'MIN_DS_ETH_VOL_HT_WK'+week,
    #             'MIN_DS_PON_UTIL_HR_FH_WK'+week:'MIN_DS_PON_UTIL_HR_HT_WK'+week,
    #             'MIN_US_PON_UTIL_HR_FH_WK'+week:'MIN_US_PON_UTIL_HR_HT_WK'+week,
    #             'MIN_RX_POWER_FH_WK'+week:'MIN_RX_POWER_HT_WK'+week}
    # my_filenames_final = {'MIN_DS_ETH_UTIL_HR_FH_WK'+week:'MIN_DS_ETH_UTIL_HR_WK'+week,
    #             'MIN_US_ETH_UTIL_HR_FH_WK'+week:'MIN_US_ETH_UTIL_HR_WK'+week,
    #             'MIN_RX_POWER_UPLINK_FH_WK'+week:'MIN_RX_POWER_UPLINK_WK'+week,
    #                     'MIN_DS_ETH_VOL_FH_WK'+week:'MIN_DS_ETH_VOL_WK'+week,
    #                     'MIN_DS_PON_UTIL_HR_FH_WK'+week:'MIN_DS_PON_UTIL_HR_WK'+week,
    #                     'MIN_US_PON_UTIL_HR_FH_WK'+week:'MIN_US_PON_UTIL_HR_WK'+week,
    #                     'MIN_RX_POWER_FH_WK'+week:'MIN_RX_POWER_WK'+week}
    # ------------------------------------------------------------------------------------------------------------------------------------

    #  ------------ for VDSL -----------------
    single_file = ['MIN_VDSL_DS_ETH_UTIL_NE_HR_HT_WK'+week, 'MIN_VDSL_US_ETH_UTIL_NE_HR_HT_WK'+week]
    my_single_filenames_final = {'MIN_VDSL_DS_ETH_UTIL_NE_HR_HT_WK'+week:'MIN_VDSL_DS_ETH_UTIL_NE_HR_HT_WK'+week,
                        'MIN_VDSL_US_ETH_UTIL_NE_HR_HT_WK'+week:'MIN_VDSL_US_ETH_UTIL_NE_HR_HT_WK'+week}
    #  ---------------------------------------


    # ----- Set Destination of crafted file ------------
    conso_raw_mycom_files_folder = os.path.join(folder_path,'Consolidated')
    if not os.path.exists(conso_raw_mycom_files_folder):
        os.makedirs(conso_raw_mycom_files_folder)

    for i in list(pair_ref):
        raw_file_name = str(i)+'.csv'
        file_path = os.path.join(raw_mycom_files_folder, raw_file_name)
        # file_path = raw_mycom_files_folder + '/' + raw_file_name
    
        path = Path(file_path)    
        
        if path.is_file():
            data1 = pd.read_csv(path)
            
            x = pair_files[[i][0]]
            raw_file_name2 = str(x)+'.csv'
            #-- path of HT raw file
            file_path2 = os.path.join(raw_mycom_files_folder, raw_file_name2)
            # file_path2 = raw_mycom_files_folder + '/' + raw_file_name2

            #-- check if HT raw file exist
            if Path(file_path2).is_file(): 
                data2 = pd.read_csv(file_path2)
                
                frames = [data1, data2]
                result = pd.concat(frames)

            else:
                result = data1
                
            file_name = my_filenames_final[[i][0]]

            out_path = os.path.join(conso_raw_mycom_files_folder, str(file_name)+'.xlsx')
            # out_path = conso_raw_mycom_files_folder + '/' + str(file_name)+'.xlsx'
            
            result.to_excel(out_path, index=False)
            
            print(file_name+' has been successfully generated!')
        
        else:
            print(f'The file {file_path} does not exist')
            continue

    for j in list(single_file):
        raw_file_name = str(j)+'.csv'
        src = os.path.join(raw_mycom_files_folder, raw_file_name)
        # src = raw_mycom_files_folder + '/' + raw_file_name
        # src ='C:/Users/lasalve/Documents/GT Files/Network Health Check/raw files/Wk'+week+'/raw/'+str(j)+'.csv'
        path = Path(src)    
        
        if path.is_file():
            data1 = pd.read_csv(path)
        
            file_name = my_single_filenames_final[[j][0]]

            out_path = os.path.join(conso_raw_mycom_files_folder, str(file_name)+'.xlsx')
            # out_path = conso_raw_mycom_files_folder + '/' + str(file_name)+'.xlsx'
            data1.to_excel(out_path, index=False)
            
    #         src = 'C:/Users/lasalve/Documents/GT Files/Network Health Check/raw files/Wk'+week+'/raw/'+str(i)+'.csv'
    #         dst = 'C:/Users/lasalve/Documents/GT Files/Network Health Check/raw files/Wk'+week+'/raw/consolidated/'+str(file_name)+'.xlsx'
    #         shutil.copyfile(src, dst)

            print(str(file_name)+' has been successfully generated!')
        
        else:
            print(f'The file {src} does not exist')
            continue

    print('\nDone consolidating all files!\n\n')
    return

def get_ht_nodes_per_region(df_ports):
    
    df_ports['new'] = df_ports.iloc[:, 0].str.extract('^(.*?),.*$')

    # ------- re-arange Header position
    ne_name = df_ports.pop('new')
    df_ports.insert(0, 'ne_name',ne_name)

    #Load Nodes to be filtered
    df_filter = pd.read_csv('HT Nodes.csv', index_col=False)
    # Convert panda to List
    df_filter = df_filter['Nodes'].values.tolist()

    #final output
    output = df_ports[df_ports.iloc[:,0].isin(df_filter)]  
    
    # --- delete ne_name
    del output["ne_name"]

    # print(output.iloc[1:2,1:].astype(float).div(1024,axis=0))

    output.iloc[:,1:] = output.iloc[:,1:].astype(float).div(1024,axis=0)
    # output.columns.values[0:1] = [""]

    return output
