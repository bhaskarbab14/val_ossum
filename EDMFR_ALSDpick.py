def edmfr_prog1(para1):

    if para1 == 'EDMFR':
        pass

    import os, webbrowser,datetime,time
    import pandas as pd
    import pickle, win32com
    from win32com import client
    import pythoncom
    import tkinter as tk
    from tkinter import messagebox, ttk


    NORM_FONT = ("Verdana", 10)

    def popupmsg(msg):
        popup = tk.Tk()
        popup.wm_title("!")
        label = ttk.Label(popup, text=msg, font=NORM_FONT)
        label.pack(side="bottom", fill="x", pady=10)
        B1 = ttk.Button(popup, text="Ok", command=popup.destroy)
        B1.pack()
        # popup.quit()
        popup.mainloop()

    def purgedir(parent):
        for root, dirs, files in os.walk(parent):
            for item in files:
                # Delete subordinate files
                filespec = os.path.join(root, item)
                # if filespec.endswith('.bak'):
                os.unlink(filespec)
            for item in dirs:
                # Recursively perform this operation for subordinate directories
                purgedir(os.path.join(root, item))

    dir = os.getcwd()
    SITE_ROOT = os.getcwd()
    username = os.environ["USERNAME"]
    ossum_path1 = 'C:\\Users\\' + username
    osssum_dir = 'OSSUM'
    ossum_path = ossum_path1 +'\\' + osssum_dir
    ALSdownloadpath = 'C:\\Users\\' + username + '\\ALSdownload'
    alspick_dir = 'ALSpickle'
    alspick_path = ossum_path  +'\\' + alspick_dir
    # alsformspickfile = alspick_path +'\\alsfilesdfforms.pickle'
    # EDMFR_dir = 'EDMFR'
    # EDMFR_path = ossum_path +'\\' + EDMFR_dir
    # datestring = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')
    # EDMFR_dt_path =EDMFR_path + '\\' + datestring
    # rada_dt_path = os.path.join(rada_path, datestring)
    # os.mkdir(datestring)
    # # os.chdir(ossum_path1)
    # print(EDMFR_dt_path)

    if os.path.exists(ossum_path1):
        if not os.path.exists(ALSdownloadpath):
            os.makedirs(ALSdownloadpath)

    if os.path.exists(ossum_path1):
        if not os.path.exists(alspick_path):
            os.makedirs(alspick_path)

    # if os.path.exists(ossum_path1):
    #     if not os.path.exists(EDMFR_dt_path):
    #         os.makedirs(EDMFR_dt_path)
    #



    ######

    path_sp = SITE_ROOT + '\Source Documents'
    # path_sp = PARENT_ROOT + '\Source Documents'
    os.chdir(path_sp)
    lists = os.listdir(path_sp)
    print(path_sp)
    # webbrowser.open(path_sp)

    #
    # pythoncom.CoInitialize()
    # xlapp = client.dynamic.Dispatch("Excel.Application")
    # xlapp.Application.Quit()
    # pythoncom.CoUninitialize()



    os.chdir(alspick_path)
    lists = os.listdir(alspick_path)
    # sub = '.pickle'
    sub = 'alsfilesdfforms.xlsx'
    # uat = 'uat'
    files = [mystr for mystr in lists if sub in mystr]
    lists=files
    print(alspick_path)


    def xcelread(filep, sheetname, usecolr):
        dfx7 = pd.DataFrame()
        dfx6 = pd.DataFrame()
        dfx5 = pd.DataFrame()

        for val in range(len(filep)):
            xls = pd.ExcelFile(filep[val])
            if sheetname in xls.sheet_names:
                dfx6 = xls.parse(sheetname, index_col=None, na_values=['NA'], parse_cols=usecolr)
                dfx5 = xls.parse('CRFDraft', index_col=None, na_values=['NA'], parse_cols='A:C')
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx7 = dfx7.append(dfx6, ignore_index=True, sort=True)
        return dfx7
    '''
    if os.path.exists(ossum_path1):
    
        filename = path_sp + '\\' + 'SharepointALSdwn.xlsm'
        xlapp = client.DispatchEx('Excel.Application')
        wbs = xlapp.Workbooks
        wb = xlapp.Workbooks.Open(Filename=filename,ReadOnly=True)
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        xlapp.DisplayAlerts = False
        wb.Save()
        xlapp.Quit()
        os.chdir(path_sp)
        df_als_sp= pd.DataFrame()
        df_als_count = pd.DataFrame({"ALS_COUNT": [0]})
        df_als_sp1 = pd.read_excel("SharepointALSdwn.xlsm")
        study_null = df_als_sp1.loc[((df_als_sp1['Study'].isna()) | (df_als_sp1['ALS Version'].isna()))]
        if ((len(df_als_sp1) > 0) & (len(study_null) == 0)):
                als_count = len(df_als_sp1)
                df_als_count = pd.DataFrame({"ALS_COUNT": [als_count]})
        else:
            als_count = 0
    
        print(len(df_als_sp))
        os.chdir(alspick_path)
        alspick_countfile = alspick_path + '\\ALS_COUNT.xlsx'
        print(alspick_countfile)
        print(os.path.exists(alspick_countfile))
        if os.path.exists(alspick_path):
            if os.path.exists(alspick_countfile):
                print(os.path.exists(alspick_countfile))
                print(df_als_sp)
                df_als_sp = pd.read_excel("ALS_COUNT.xlsx")
            else:
                print("test")
                df_als_count.to_excel("ALS_COUNT.xlsx")
        else:
            print("test")
            df_als_count.to_excel("ALS_COUNT.xlsx")
    
        xls = 'alsfilesdfforms.xlsx'
    
        alspickfilepath = alspick_path + '\\'+ xls
        xmod = datetime.datetime.today()
        # if os.path.exists(alspickfilepath):
        if os.path.exists(alspickfilepath):
            modx = os.path.getmtime(xls)
            xmod = datetime.datetime.fromtimestamp(modx)
        #
        print((~(df_als_count.iloc[0, 0] == als_count)))
        print((als_count))
        print(((df_als_count.iloc[0, 0])))
    
        # if (( not ((datetime.datetime.today().date() == xmod.date())==True)) | (~(df_als_count.iloc[0, 0] == als_count)) | ((len(lists)==0))):
        # if  (~(df_als_count.iloc[0, 0] == als_count)):
        print(( not ((datetime.datetime.today().date() == xmod.date())==True)))
        print(( (~(datetime.datetime.today() == xmod))))
        print(datetime.datetime.today())
        print(( not ((datetime.datetime.today().date() == xmod.date())==True)))
        print(datetime.datetime.today().date())
        print(xmod.date())
    '''
    os.chdir(ALSdownloadpath)
    lists = os.listdir(ALSdownloadpath)
    sub = '.xlsx'
    global_list = ['Cardio-Metabolic_1.0_Cardio-Metabolic_18AUG2020.xlsx', 'IHD_4.0_IHD_18AUG2020.xlsx',
                   'Global_11.0_GLOBAL_27AUG2020.xlsx', 'Neuroscience_2.0_NEUROSCIENCE_20AUG2020.xlsx',
                   'Oncology_6.0_ONCOLOGY_14AUG2020.xlsx', 'Ophthalmology_2.0_OPHTHALMOLOGY_15JUN2020.xlsx',
                   'Questionnaires_3.0_QUESTIONNAIRES_05AUG2020.xlsx', 'Respiratory_3.0_RESPIRATORY_10JUN2020.xlsx']

    files = [mystr for mystr in lists if sub in mystr and mystr not in global_list]

    print(files)

    print("reading alss")
    df_forms = xcelread(files, 'Forms', 'A:O')
    # C:\Users\pillibh2\OneDrive - Novartis Pharma AG\Desktop\Als_dwnld
    os.chdir(alspick_path)
    df_forms.to_excel("alsfilesdfforms.xlsx")
            # callfn()
    os.chdir(dir)
    popupmsg("Done with ALSs reading, please click on 'Run the generate report'")
    print("Done with ALSs reading, please click on 'Run the generate report'")
    '''
        if ((  (not (datetime.datetime.today().date()  == xmod))) | (~(df_als_count.iloc[0, 0] == als_count)) | ((len(lists)==0))):
            purgedir(ALSdownloadpath)
            print(len(study_null))
            if len(study_null) == 0:
                xl = client.Dispatch('Excel.Application')
                # wbs = xl.Workbooks
                wb = xl.Workbooks.Open(Filename=filename, ReadOnly=True)
                xl.Application.Run("SharepointALSdwn.xlsm!Module1.DownloadFile_From_URL")
                # wb.close()
                wb.Save()
                wb.Close()
                xl.Quit()
                # xl.Application.Quit()
                # purgedir(alspick_path)
                # if os.path.exists(alsformspickfile):
                #     os.remove(alsformspickfile)
            # if ((not ((datetime.datetime.today().date() == xmod.date()) == True)) | (not (os.path.exists(alspickfilepath)))):
                os.chdir(ALSdownloadpath)
                lists = os.listdir(ALSdownloadpath)
                sub = '.xlsx'
                files = [mystr for mystr in lists if sub in mystr]
                print("reading alss")
                df_forms = xcelread(files, 'Forms', 'A:O')
                
                os.chdir(alspick_path)
                df_forms.to_excel("alsfilesdfforms.xlsx")
                # pickle_out2 = open("alsfilesdfforms.pickle", "wb")
                # pickle.dump(df_forms, pickle_out2)
                # pickle_out2.close()
                if not (df_als_count.iloc[0, 0] == als_count):
                    df_als_count.to_excel("ALS_COUNT.xlsx")
    '''


    # pythoncom.CoInitialize()
    # xlapp = client.dynamic.Dispatch("Excel.Application")
    # xlapp.Application.Quit()
    # pythoncom.CoUninitialize()


    # elif len(lists) == 0:
    #     callfn()

    # SITE_ROOT = dir
    # if os.path.exists(alspick_path):
    # #     webbrowser.open(alspick_path)
    # #
    # #
    # # if os.path.exists(ALSdownloadpath):
    # #     webbrowser.open(ALSdownloadpath)


# edmfr_prog1('EDMFR')