def als_cons_prog_rada(para1, para2, para3, para4, para5, para6, para7, para8):

    import os
    # webbrowser,datetime,time, pyodbc
    import pandas as pd
    import pickle, win32com
    from win32com import client
    import pythoncom
    import numpy as np
    import tkinter as tk
    import pyexcelerate
    from joblib import Parallel, delayed
    import win32com.client as win32
    # from pyexcelerate import Workbook

    df_forms = pd.DataFrame()
    df_forms = pd.DataFrame()
    df_dicts = pd.DataFrame()
    df_dict_entries= pd.DataFrame()
    df_checks = pd.DataFrame()
    df_CheckSteps = pd.DataFrame()
    df_CheckActions = pd.DataFrame()
    df_DerivationSteps= pd.DataFrame()
    df_Derivations= pd.DataFrame()
    df_CheckSteps_500= pd.DataFrame()
    df_CheckSteps_501= pd.DataFrame()
    df_LabVariableMappings = pd.DataFrame()
    df_CustomFunctions = pd.DataFrame()
    #
    # sheetnames = ['CRFDraft','Forms', 'Fields', 'DataDictionaries', 'DataDictionaryEntries', 'Checks', 'CheckSteps',
    #               'CheckActions', 'Derivations', 'Matrices', 'DerivationSteps', 'LabVariableMappings', 'Folders',
    #               'CustomFunctions','CoderConfiguration','CoderSupplementalTerms', 'MASTERDASHBOARD', 'EXPAN']

    # sheetnames = ['Folders']
    if len(para5) > 0 :
        sheetnames = ['CRFDraft']
        print("para5", para5)
        sheetnames.extend(para5)
    else:
        # sheetnames = ['CRFDraft']
        # sheetnames.extend(para4)
        sheetnames = para4
        para5 = []
        print("para4", para4)
    # sheetnames = ['CRFDraft','Forms', 'Fields']
    # dir = os.getcwd()
    # SITE_ROOT = os.getcwd()
    # username = os.environ["USERNAME"]
    # ossum_path1 = 'C:\\Users\\' + username
    # osssum_dir = 'OSSUM'
    # ossum_path = ossum_path1 +'\\' + osssum_dir
    # cons_ALS_Global_path = 'Consolidated_ALS_Global'
    # ossum_cons_path = ossum_path + '\\' + cons_ALS_Global_path
    # if para1 == 'Cons_Study':
    #     ALSdownloadpath = 'C:\\Users\\' + username + '\\ALSdownload'
    # else:
    #     ALSdownloadpath = 'C:\\Users\\' + username + '\\ALSdownloadstd'
    #
    # if os.path.exists(ossum_path1):
    #     path_sp = SITE_ROOT + '\Source Documents'
    #     filename = path_sp + '\\' + 'SharepointALSdwn.xlsm'
    #     filename1 = path_sp + '\\' + 'StandardALSs.xlsm'
    #     filename2 = path_sp + '\\' + 'StandardALSdwn.xlsm'
    #     if not os.path.exists(ALSdownloadpath):
    #         os.makedirs(ALSdownloadpath)
    #
    #
    #
    # dir = os.getcwd()
    # path_sp = SITE_ROOT + '\Source Documents'
    #
    # os.chdir(path_sp)
    # lists = os.listdir(path_sp)
    df_ALS_SP = pd.DataFrame()
#################################

    # df_ALS_SP = pd.DataFrame()
    # if para1 == 'Cons_Study':

##############################################
    #     # df_ALS_SP = pd.read_sql(query, conn)
    #     # conn.close()
    # os.chdir(para6)
    # lists = os.listdir(para6)
    # 
    # df_ALS_SP = pd.read_excel("SharepointALSdwn.xlsm")
    # df_ALS_SP = df_ALS_SP.sort_values(by=['ALS Version'])
    #
        # df_ALS_SP = df_ALS_SP.sort_values(by=['ALS Version'])
        # df_ALS_SP.sort_values('ALS Version').groupby('Study').tail(1)
        # df_ALS_SP = df_ALS_SP.sort_values('ALS Version').groupby('Study').tail(1)
    #
        # df_ALS_SP = df_ALS_SP.filter(items=['Study', 'ALS Version', 'Name'])
    #
    #     # print(df_ALS_SP['Name'])


    # else:
    #
    #     ###########################
    # DATABASE1 = "Consolidated_ALS.xlsx"
    # DATABASE = para1 + "\\" + DATABASE1

    # para2 = 'GV_Study'
    # para2 = 'GV_Study'
    print("para4 len",len((para4)))
    print("para5 len",len((para5)))
    if (para2 == 'Cons_Study') & (len(para5) == 0) :
        DATABASE1 = "Consolidated_ALS.xlsx"
        DATABASE = para3 + "\\" + DATABASE1
        print("Consolidated_ALS is done DATABASE", DATABASE)
        print("Consolidated_ALS is done")

    elif (para2 == 'GV_Study') & (len(para5) == 0):
        DATABASE1 = "Consolidated_ALS_Global.xlsx"
        DATABASE = para3 + "\\" + DATABASE1
        print("Consolidated_ALS_Global is done")
        print("Consolidated_ALS is done DATABASE", DATABASE)

    elif (len(para5) > 0) & (('Study_Summary' not in para2) ):
        DATABASE1 = "Consolidated_ALS_Selected_Tabs(CAST).xlsx"
        DATABASE = para3 + "\\" + DATABASE1
        print("cast Consolidated_ALS_Global is done")
        print("Consolidated_ALS is done DATABASE", DATABASE)
    elif ((len(para5) == 0)) & (('Study_Summary' in para2) ):
        DATABASE1 = "Study Summary.xlsx"
        DATABASE = para3 + "\\" + DATABASE1
        print("Study Summary_is done")
        print("Study Summary_", DATABASE)
    #     DATABASE1 = "SG_Codelist.accdb"
    #     DATABASE = path_sp + "\\" + DATABASE1
    #     # conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % (DATABASE))
    #     # cursor = conn.cursor()
    #     # query = 'select * from [RaveX Standards]'
    #     # df_standard_als = pd.read_sql(query, conn)
    #     # df_standard_als = pd.read_excel("StandardALSs.xlsm")
    #     ###############################
    #
    #     xlapp = win32.gencache.EnsureDispatch('Excel.Application')
    #     wb = xlapp.Workbooks.Open(Filename=filename2, ReadOnly=0)
    #     wb.RefreshAll()
    #     xlapp.CalculateUntilAsyncQueriesDone()
    #     xlapp.DisplayAlerts = False
    #     wb.Save()
    #     xlapp.Quit()
    #     os.chdir(path_sp)
    #
    #     df_standard_als = pd.read_excel("StandardALSdwn.xlsm")
    #
    #     df_standard_als = df_standard_als.loc[(df_standard_als["Research Area Info (optional)"] == "ALS") &
    #                                           (df_standard_als["Status"] == "Final")]
    #     df_standard_als = df_standard_als.filter(items=['Name'])
    #
    #     # df_standard_als.loc[df_standard_als['Name'].astype(str).str.contains('#'), 'Name'] = \
    #     #     df_standard_als['Name'].astype(str).str.split('#', 1).str.get(0)
    #     df_ALS_SP = df_standard_als.loc[~df_standard_als['Name'].isna()]
    #     df_ALS_SP = df_ALS_SP.filter(items=["Name"])
    #     df_ALS_SP['Study'] =''
    #     df_ALS_SP['ALS Version'] =''
    #     df_ALS_SP.loc[df_ALS_SP['Name'].astype(str).str.contains('#'), 'Study'] = \
    #         df_ALS_SP['Name'].astype(str).str.split('_', 1).str.get(0)
    #     df_ALS_SP.loc[df_ALS_SP['Name'].astype(str).str.contains('#'), 'ALS Version'] = \
    #         df_ALS_SP['Name'].astype(str).str.split('_', 2).str.get(0)


############################################
    #
    # DATABASE1 = "SG_Codelist.accdb"
    # DATABASE = path_sp + "\\" + DATABASE1
    #
    # os.chdir(r"C:\Bhasp\NVTSonco-work\NVTSonco-work\RADA\Radaprogram\RADA\SPMDR")
    # os.chdir(dir)
    #
    # conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % ( DATABASE ))
    # cursor = conn.cursor()
    # query = 'select * from [ALS Repository]'
    #
    # df_ALS_SP = pd.read_sql(query, conn)
    # df_ALS_SP = pd.read_sql(query, conn)
    # conn.close()

    os.chdir(para7)
    print("para7 para",para7)
    lists = os.listdir(para7)
    if ((len(lists) > 0)) :
        print("para7 para", lists[0])
        df_ibow= pd.read_excel( lists[0])

        df_ibow = df_ibow.filter(items=['Trial', 'DM Ops Cluster','DB Phase','Actual DB Go-Live Date'])
        df_ibow = df_ibow.rename(columns={'Trial': "Study ID",'DM Ops Cluster':"Study Cluster",
                                          'DB Phase':"Study Status",'Actual DB Go-Live Date': "DB Go-Live Date"})

    os.chdir(para8)
    lists = os.listdir(para8)
    if ((len(lists) > 0)) :
        df_idts= pd.read_excel(lists[0])
        df_idts['# of DTS'] = 0
        f = {'# of DTS': 'count'}
        df_idts = df_idts.groupby(['Trial'], as_index=False).agg(f)
        df_idts = df_idts.rename(columns={'Trial': "Study ID"})


    os.chdir(para6)
    lists = os.listdir(para6)

    df_ALS_SP = pd.read_excel("SharepointALSdwn.xlsm")


    #
    df_ALS_SP = df_ALS_SP.sort_values(by=['ALS Version'])
    df_ALS_SP.sort_values('ALS Version').groupby('Study').tail(1)
    df_ALS_SP = df_ALS_SP.sort_values('ALS Version').groupby('Study').tail(1)

###############################################

    NORM_FONT = ("Verdana", 10)
    #
    # def popupmsg(msg):
    #     popup = tk.Tk()
    #     popup.wm_title("!")
    #     label = ttk.Label(popup, text=msg, font=NORM_FONT)
    #     label.pack(side="top", fill="x", pady=10)
    #     B1 = ttk.Button(popup, text="Ok", command=popup.destroy)
    #     B1.pack()
    #     popup.mainloop()

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
    #
    #
    # if os.path.exists(ossum_path1):
    #     if not os.path.exists(ALSdownloadpath):
    #         os.makedirs(ALSdownloadpath)

    # ,options={'strings_to_formulas': False}

    def pyexecelerate_to_excel(workbook_or_filename, df, sheet_name='Sheet1', origin=(1, 1), columns=True, index=False):
        """
        Write DataFrame to excel file using pyexelerate library
        """
        if not isinstance(workbook_or_filename, pyexcelerate.Workbook):
            location = workbook_or_filename
            workbook_or_filename = pyexcelerate.Workbook()
        else:
            location = None
        worksheet = workbook_or_filename.new_sheet(sheet_name)

        # Account for space needed for index and column headers
        column_offset = 0
        row_offset = 0

        if index:
            index = df.index.tolist()
            ro = origin[0] + row_offset
            co = origin[1] + column_offset
            worksheet.range((ro, co), (ro + 1, co)).value = [['Index']]
            worksheet.range((ro + 1, co), (ro + 1 + len(index), co)).value = list(map(lambda x: [x], index))
            column_offset += 1
        if columns:
            columns = df.columns.tolist()
            ro = origin[0] + row_offset
            co = origin[1] + column_offset
            worksheet.range((ro, co), (ro, co + len(columns))).value = [[*columns]]
            row_offset += 1

        # Write the data
        row_num = df.shape[0]
        col_num = df.shape[1]
        ro = origin[0] + row_offset
        co = origin[1] + column_offset
        worksheet.range((ro, co), (ro + row_num, co + col_num)).value = df.values.tolist()

        if location:
            workbook_or_filename.save(location)
    #

    def xcelread_studysum(filep):
        dfx7 = []
        # dfx7 = pd.DataFrame()
        dfx6 = pd.DataFrame()
        dfx5 = pd.DataFrame()
        list_df = []
        print("filep",filep)
        xls = pd.ExcelFile(filep,engine='openpyxl')

        dfx5 = xls.parse(sheet_name='CRFDraft', index_col=None, usecols='A:C', keep_default_na=False,
                         na_values=[''])
        print("dfx5",dfx5)
        dfx5.dropna(axis=0, how='all', inplace=True)
        # for sheetname in xls.sheet_names:
        sheet_names = xls.sheet_names
        sheet_names_sr = pd.Series(sheet_names)
        if 'Folders' in sheet_names:
            # if sheetname in sheetnames:
            dfx6 = xls.parse(sheet_name='Folders', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                # dfx6 = dfx6.loc[~dfx6['OID'].str.contains('0|1|2|3|4|5|6|7|8|9')]
                dfx6 = dfx6.loc[dfx6['OID'].astype(str).str.isdigit(), ['OID']]
                count_df = len(dfx6) - 1
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Visits'] = count_df
                dfx6 = dfx6.filter(items=['# of Visits'])
                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] = 'Folders'

                dfx7.append(dfx6)
        if 'Forms' in sheet_names:
            dfx6 = xls.parse(sheet_name='Forms', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                dfx6 = dfx6.loc[((dfx6['DraftFormActive'] == True) | (dfx6['DraftFormActive'] == 'TRUE')),
                                ['DraftFormName', 'DraftFormActive']]
                count_df = len(dfx6) -1
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Forms'] = count_df
                dfx6 = dfx6.filter(items=['# of Forms'])
        # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] =  'Forms'
                # dfx6['filename'] = filep
                dfx7.append(dfx6)
        if 'Checks' in sheet_names:
            dfx6 = xls.parse(sheet_name='Checks', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                dfx6 = dfx6.loc[((dfx6['CheckActive'] == True) | (dfx6['CheckActive'] == 'TRUE')), ['CheckActive']]
                count_df = len(dfx6)
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Checks'] = count_df
                dfx6 = dfx6.filter(items=['# of Checks'])

                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] =  'Checks'
                # dfx6['filename'] = filep
                dfx7.append(dfx6)
        if 'CustomFunctions' in sheet_names:
            dfx6 = xls.parse(sheet_name= 'CustomFunctions', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                count_df = len(dfx6)
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Custom Functions'] = count_df
                dfx6 = dfx6.filter(items=['# of Custom Functions'])
                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] = 'CustomFunctions'
                # dfx6['filename'] = filep
                dfx7.append(dfx6)
        if 'Derivations' in sheet_names:
            dfx6 = xls.parse(sheet_name='Derivations', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                dfx6 = dfx6.loc[((dfx6['Active'] == True) | (dfx6['Active'] == 'TRUE')), ['Active']]
                count_df = len(dfx6)
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Derivations'] = count_df
                dfx6 = dfx6.filter(items=['# of Derivations'])
                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] = 'Derivations'
                # dfx6['filename'] = filep
                dfx7.append(dfx6)

        if 'Fields' in sheet_names:
            dfx6= xls.parse(sheet_name= 'Fields', index_col=None, keep_default_na=False,
                                 na_values=[''])

            if len(dfx6) > 0:
                dfx6.dropna(axis=0, how='all', inplace=True)
                dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                colname = dfx6.columns[0]
                dfx6 = dfx6.loc[~dfx6[colname].isna()]

                dfx6 = dfx6.loc[((dfx6['DraftFieldActive'] == True) | (dfx6['DraftFieldActive'] == 'TRUE'))
                                # &
                                # (((dfx6['IsRequired'] == True) | (dfx6['IsRequired'] == 'TRUE')) &
                                #  ((dfx6['QueryNonConformance'] == True) | (dfx6['QueryNonConformance'] == 'TRUE')) &
                                #  ((dfx6['QueryFutureDate'] == True) | (dfx6['QueryFutureDate'] == 'TRUE')))
                ,
                                ['FormOID', 'FieldOID', 'DraftFieldActive', 'IsRequired', 'QueryNonConformance',
                                 'QueryFutureDate']]
                # ['FormOID','FieldOID','DraftFieldActive','IsRequired','QueryNonConformance','QueryFutureDate']]
                dfx6_1 = xls.parse(sheet_name='Forms', index_col=None, keep_default_na=False,
                                   na_values=[''])
                dfx6_1 = dfx6_1.loc[((dfx6_1['DraftFormActive'] == True) | (dfx6_1['DraftFormActive'] == 'TRUE')),
                                    ['OID', 'DraftFormActive']]

                dfx6 = dfx6.loc[dfx6['FormOID'].isin(dfx6_1['OID'])]
                dfx6_isreq = dfx6.loc[((dfx6['IsRequired'] == True) | (dfx6['IsRequired'] == 'TRUE'))]
                dfx6_isconf = dfx6.loc[
                    ((dfx6['QueryNonConformance'] == True) | (dfx6['QueryNonConformance'] == 'TRUE'))]
                dfx6_isfutr = dfx6.loc[((dfx6['QueryFutureDate'] == True) | (dfx6['QueryFutureDate'] == 'TRUE'))]
                count_isreq = len(dfx6_isreq)
                count_isconf = len(dfx6_isconf)
                count_isfutr = len(dfx6_isfutr)
                count_df = count_isreq + count_isconf + count_isfutr
                dfx6 = dfx6.iloc[:1]
                dfx6['# of Field Edit Checks'] = count_df
                dfx6 = dfx6.filter(items=['# of Field Edit Checks'])

                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] = 'Fields'
                # dfx6['filename'] = filep
                dfx7.append(dfx6)


        if (sheet_names_sr.str.contains('|'.join(['MASTERDASHBOARD','EXPAN'])) == True).any():
        # elif (sheet_names_sr.str.contains(sheetname) == True).any():
            sub1 = 'MASTERDASHBOARD'
            sub2 = 'EXPAN'
            # sheetnamesr = shtnames[shtnames.str.contains(sheetname)]
            sheetname_res = [i for i in sheet_names if ((sub1 in i) or (sub2 in i))]
            if len(sheetname_res) > 0:
                sheetname = sheetname_res[0]
                dfx6 = xls.parse(sheet_name=sheetname, index_col=None, keep_default_na=False,
                                 na_values=[''])
                dfx6 = dfx6.iloc[:, 1:]
                print("dfx6 columnss",dfx6)
                bool_dfvalues = ~dfx6.isna()
                count_na_columwise = bool_dfvalues.sum()
                count_na_dataframewise = count_na_columwise.sum()
                dfx6['# of Total CRFs'] = count_na_dataframewise
                # dfx6['Total CRFs'] = count_na_dataframewise
                dfx6 = dfx6.filter(items=['# of Total CRFs'])
                # dfx6 = dfx6.filter(items=['Total CRFs'])
                print("dfx6 columnss",dfx6.columns)
                dfx6 = dfx6.iloc[:1]
                # if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
                dfx6['sheetname'] = sheetname
                dfx7.append(dfx6)

        dfx7 = pd.concat(dfx7, ignore_index=True, sort=True)
        # else:
        #
        #     dfx7 = dfx6

        return dfx7


    #
    # def xcelread_studysumdel(filep,sheetname):
    #     dfx7 = []
    #     # dfx7 = pd.DataFrame()
    #     dfx6 = pd.DataFrame()
    #     dfx5 = pd.DataFrame()
    #     list_df = []
    #     print("filep",filep)
    #     xls = pd.ExcelFile(filep,engine='openpyxl')
    #
    #     dfx5 = xls.parse(sheet_name='CRFDraft', index_col=None, usecols='A:C', keep_default_na=False,
    #                      na_values=[''])
    #     print("dfx5",dfx5)
    #     dfx5.dropna(axis=0, how='all', inplace=True)
    #     # for sheetname in xls.sheet_names:
    #     sheet_names = xls.sheet_names
    #     sheet_names_sr = pd.Series(sheet_names)
    #     if sheetname in xls.sheet_names:
    #         # if sheetname in sheetnames:
    #         dfx6 = xls.parse(sheet_name=sheetname, index_col=None, keep_default_na=False,
    #                          na_values=[''])
    #
    #     elif (sheet_names_sr.str.contains(sheetname) == True).any():
    #         sub1 = 'MASTERDASHBOARD'
    #         sub2 = 'EXPAN'
    #         # sheetnamesr = shtnames[shtnames.str.contains(sheetname)]
    #         sheetname_res = [i for i in sheet_names if ((sub1 in i) or (sub2 in i))]
    #         if len(sheetname_res) > 0:
    #             sheetname = sheetname_res[0]
    #             dfx6 = xls.parse(sheet_name=sheetname, index_col=None, keep_default_na=False,
    #                              na_values=[''])
    #             dfx6 = dfx6.iloc[:, 1:]
    #             print("dfx6 columnss",dfx6)
    #             bool_dfvalues = ~dfx6.isna()
    #             count_na_columwise = bool_dfvalues.sum()
    #             count_na_dataframewise = count_na_columwise.sum()
    #             dfx6['Total CRFs'] = count_na_dataframewise
    #             dfx6 = dfx6.filter(items=['Total CRFs'])
    #             print("dfx6 columnss",dfx6.columns)
    #             dfx6 = dfx6.iloc[:1]
    #
    #
    #     if len(dfx6) > 0:
    #         dfx6.dropna(axis=0, how='all', inplace=True)
    #         dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
    #         colname = dfx6.columns[0]
    #         dfx6 = dfx6.loc[~dfx6[colname].isna()]
    #         # if not sheetname in ['CRFDraft']:
    #         dfx6['Study ID'] = dfx5.iloc[0, 2]
    #             # dfx6['DraftName'] = dfx5.iloc[0, 0]
    #         dfx6['sheetname'] = sheetname
    #         # dfx6['filename'] = filep
    #
    #         dfx7.append(dfx6)
    #
    #         dfx7 = pd.concat(dfx7, ignore_index=True, sort=True)
    #     else:
    #
    #         dfx7 = dfx6
    #
    #     return dfx7
    #
    #

    def xcelread_para(filep,sheetname):
        dfx7 = []
        # dfx7 = pd.DataFrame()
        dfx6 = pd.DataFrame()
        dfx5 = pd.DataFrame()
        list_df = []
        print("filep",filep)
        xls = pd.ExcelFile(filep,engine='openpyxl')
        dfx5 = xls.parse(sheet_name='CRFDraft', index_col=None, usecols='A:C', keep_default_na=False,
                         na_values=[''])
        print("dfx5",dfx5)
        dfx5.dropna(axis=0, how='all', inplace=True)
        # for sheetname in xls.sheet_names:
        if sheetname in xls.sheet_names:
            # if sheetname in sheetnames:
            dfx6 = xls.parse(sheet_name=sheetname, index_col=None, keep_default_na=False,
                             na_values=[''])
        if len(dfx6) > 0:
            dfx6.dropna(axis=0, how='all', inplace=True)
            dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
            colname = dfx6.columns[0]
            dfx6 = dfx6.loc[~dfx6[colname].isna()]
            if not sheetname in ['CRFDraft']:
                dfx6['Study ID'] = dfx5.iloc[0, 2]
                # dfx6['DraftName'] = dfx5.iloc[0, 0]
            dfx6['sheetname'] = sheetname
            # dfx6['filename'] = filep

            dfx7.append(dfx6)

            # dfx7 = dfx7.append(dfx6, ignore_index=True, sort=True)
            # if not sheetname in ['CRFDraft']:
            #     print(type(dfx7))
            #     print(len(dfx7))
            dfx7 = pd.concat(dfx7, ignore_index=True, sort=True)
        else:

            dfx7 = dfx6

        return dfx7



    def xcelread_all(filep, sheetnames):
    # def xcelread(filep, sheetname, usecolr, ** kwargs):
        dfx7 = pd.DataFrame()
        dfx6 = pd.DataFrame()
        dfx5 = pd.DataFrame()


        for val in range(len(filep)):
            xls = pd.ExcelFile(filep[val])
            dfx5 = xls.parse(sheet_name='CRFDraft', index_col=None, usecols='A:C', keep_default_na=False,
                             na_values=[''])
            dfx5.dropna(axis=0, how='all', inplace=True)
            for sheetname in xls.sheet_names:
                if sheetname in sheetnames:
                    dfx6 = xls.parse(sheet_name=sheetname, index_col=None, keep_default_na=False,
                             na_values=[''])
                    dfx6.dropna(axis=0, how='all', inplace=True)
                    dfx6 = dfx6.loc[:, ~dfx6.columns.str.contains('^Unnamed')]
                    if not sheetname in ['CRFDraft']:
                        dfx6['Study ID'] = dfx5.iloc[0, 2]
                        # dfx6['DraftName'] = dfx5.iloc[0, 0]
                        # dfx6['sheetname'] = sheetname
                    dfx7.append(dfx6)
        dfx7 = dfx7.append(dfx6, ignore_index=True, sort=True)
        return dfx7
    #
    # os.chdir(ALSdownloadpath)
    # lists = os.listdir(ALSdownloadpath)
    lists = os.listdir(para1)
    sub = '.xlsx'
    # global_list = ['Cardio-Metabolic_1.0_Cardio-Metabolic_18AUG2020.xlsx', 'IHD_4.0_IHD_18AUG2020.xlsx',
    #                'Global_11.0_GLOBAL_27AUG2020.xlsx', 'Neuroscience_2.0_NEUROSCIENCE_20AUG2020.xlsx',
    #                'Oncology_6.0_ONCOLOGY_14AUG2020.xlsx', 'Ophthalmology_2.0_OPHTHALMOLOGY_15JUN2020.xlsx',
    #                'Questionnaires_3.0_QUESTIONNAIRES_05AUG2020.xlsx', 'Respiratory_3.0_RESPIRATORY_10JUN2020.xlsx']
    # files = para1
    # lists = para1
    files = [mystr for mystr in lists if sub in mystr ]
    # files = [mystr for mystr in lists if sub in mystr and mystr not in global_list]
    # files =['CINC280D2201_Architect Loader Spreadsheet_Version 12.0.xlsx']
            # 'CLAG525B2101_Architect Loader Spreadsheet_Version12.xlsx']
    # print(files)

    print("reading alss")
    # print(ALSdownloadpath)
    # os.chdir(ALSdownloadpath)
    # os.chdir(r"C:\Users\PILLIBH2\ALSdownload")
    # print(os.curdir)

    sub = '.xlsx'

    # df_alltabs = xcelread_all(files, ['Forms','Fields','DataDictionaries','DataDictionaryEntries','Checks','CheckSteps',
    #                                   'CheckActions', 'Derivations', 'DerivationSteps','LabVariableMappings',
    #                                   'CustomFunctions' ])
    print("list para", files)
    # print("list type", type(files[0]))
    # print("list para lists", lists)
    df_alltabs = []
    # df_crfdraft = xcelread_para(files[0],'CRFDraft')
    alsfiles = []
    for file in files:
        alsfiles.append(para1 + '\\' + file)

    # print("alsfiles",alsfiles)
    # print("df_crfdraft",df_crfdraft)
    # print("file_number",[print(file_number) for file_number in alsfiles])
    print("sheetnames in py",sheetnames)
    print("sheetname",[print(sheetname)  for sheetname in sheetnames])
    if 'Study_Summary' in para2:
        df_alltabs = Parallel(n_jobs=-1, verbose=0, prefer="threads")(delayed(xcelread_studysum)(file_number)
                                                                      for file_number in alsfiles)

    else:
        df_alltabs = Parallel(n_jobs=-1, verbose=0, prefer="threads")(delayed(xcelread_para)(file_number,sheetname)
                                                                      for file_number in alsfiles for sheetname in sheetnames)

    print("type(df_alltabs)",type(df_alltabs))
    print(df_alltabs)
    # df = Parallel(n_jobs=-1, verbose=0, prefer="threads")(delayed(xcelread_para)(file_number) for file_number in files)
    if len(df_alltabs) > 0 :
        df_alltabs = pd.concat(df_alltabs, ignore_index=True)

    df_ALS_SP = df_ALS_SP.filter(items=['Study', 'ALS Version'])
    # df_ALS_SP = df_ALS_SP.filter(items=['Study', 'ALS Version','Name'],'Name'])
    df_ALS_SP = df_ALS_SP.rename(columns={'Study': "Study ID"})
    # df_ALS_SP.loc[df_ALS_SP['Name'].astype(str).str.contains('#'),'Name'] = df_ALS_SP['Name'].astype(str).str.split('#', 1).str.get(0)
    # print("df_ALS_SP.loc['Name']",df_ALS_SP['Name'])

    #######################################

    df_alltabs = df_alltabs.merge(df_ALS_SP, left_on=['Study ID'], right_on=['Study ID'], how='left',
                    suffixes=['', '_'], indicator=True)



    ######################################
    if len(df_alltabs) > 0:
        # df_alltabs.to_excel(r'C:\Bhasp\NVTSonco-work\NVTSonco-work\RADA\ALS\als_all.xlsx')
        print("reading alss")

        df_alltabs.replace(np.nan, '', inplace=True)
        df_forms = df_alltabs.loc[df_alltabs["sheetname"] == 'Forms']
        df_fields = df_alltabs.loc[df_alltabs["sheetname"] == 'Fields']
        df_folders = df_alltabs.loc[df_alltabs["sheetname"] == 'Folders']
        df_checks = df_alltabs.loc[df_alltabs["sheetname"] == 'Checks']
        df_Derivations = df_alltabs.loc[df_alltabs["sheetname"] == 'Derivations']
        df_CustomFunctions = df_alltabs.loc[df_alltabs["sheetname"] == 'CustomFunctions']
        if (('Study_Summary' in para2) ):
            df_masterdashboard = df_alltabs.loc[df_alltabs["sheetname"].str.contains('MASTERDASHBOARD|EXPAN', case=False)]
        if (('Study_Summary' not in para2) ):
            df_CRFDraft = df_alltabs.loc[df_alltabs["sheetname"] == 'CRFDraft']
            df_dicts = df_alltabs.loc[df_alltabs["sheetname"] == 'DataDictionaries']
            df_dict_entries = df_alltabs.loc[df_alltabs["sheetname"] == 'DataDictionaryEntries']
            # df_matx = df_alltabs.loc[df_alltabs["sheetname"] == 'Matrices']
            df_CheckSteps = df_alltabs.loc[df_alltabs["sheetname"] == 'CheckSteps']
            df_CheckActions = df_alltabs.loc[df_alltabs["sheetname"] == 'CheckActions']
            print("reading alss")
            df_DerivationSteps = df_alltabs.loc[df_alltabs["sheetname"] == 'DerivationSteps']
            df_LabVariableMappings = df_alltabs.loc[df_alltabs["sheetname"] == 'LabVariableMappings']
            df_Matrices = df_alltabs.loc[df_alltabs["sheetname"] == 'Matrices']
            df_CoderConfiguration = df_alltabs.loc[df_alltabs["sheetname"] == 'CoderConfiguration']
            df_CoderSupplementalTerms = df_alltabs.loc[df_alltabs["sheetname"] == 'CoderSupplementalTerms']

        if (('Study_Summary' not in para2) ):

            crfdraft_cols = ["DraftName", "DeleteExisting","ProjectName","ProjectType", "PrimaryFormOID","DefaultMatrixOID",
                             "ConfirmationMessage", "SignaturePrompt", "LabStandardGroup", "ReferenceLabs", "AlertLabs",
                             "SyncOIDProject", "SyncOIDDraft", "SyncOIDProjectType", "SyncOIDOriginIsVersion",'ALS Version']

            df_CRFDraft = df_CRFDraft.filter(items=crfdraft_cols)
            form_cols = ['Study ID',"OID", "Ordinal","DraftFormName", "DraftFormActive", "HelpText","IsTemplate", "IsSignatureRequired",
                         "IsEproForm", "ViewRestrictions","EntryRestrictions","LogDirection", "DDEOption", "ConfirmationStyle",
                         "LinkFolderOID", "LinkFormOID",'ALS Version']
            folders_cols = ['Study ID',"OID","Ordinal","FolderName","AccessDays","StartWinDays", "Targetdays",	"EndWinDays", "OverDueDays",
                            "CloseDays","ParentFolderOID","IsReusable",'ALS Version']
            df_forms = df_forms.filter(items=form_cols)
            fields_cols = ['Study ID',"FormOID", "FieldOID","Ordinal", "DraftFieldNumber","DraftFieldName","DraftFieldActive","VariableOID",
                           "DataFormat","DataDictionaryName","UnitDictionaryName","CodingDictionary","ControlType",
                           "AcceptableFileExtensions","IndentLevel","PreText","FixedUnit","HeaderText","HelpText",
                           "SourceDocument","IsLog","DefaultValue","SASLabel","SASFormat","EproFormat","IsRequired",
                           "QueryFutureDate","IsVisible","IsTranslationRequired","AnalyteName","IsClinicalSignificance",
                           "QueryNonConformance","OtherVisits","CanSetRecordDate","CanSetDataPageDate","CanSetInstanceDate",
                           "CanSetSubjectDate","DoesNotBreakSignature","LowerRange","UpperRange","NCLowerRange","NCUpperRange",
            "ViewRestrictions","EntryRestrictions","ReviewGroups","IsVisualVerify",'ALS Version']
            df_fields = df_fields.filter(items=fields_cols)
            df_folders = df_folders.filter(items=folders_cols)
            matx_cols = ['Study ID','MatrixName','OID','Addable','Maximum','ALS Version']
            df_Matrices= df_Matrices.filter(items=matx_cols)
            dict_cols = ['Study ID',"DataDictionaryName",'ALS Version']
            df_dicts = df_dicts.filter(items=dict_cols)
            dict_entries_cols = ['Study ID',"DataDictionaryName", "CodedData","Ordinal","UserDataString","Specify",'ALS Version']
            df_dict_entries = df_dict_entries.filter(items=dict_entries_cols)
            checks_cols = ['Study ID',"CheckName","CheckActive","BypassDuringMigration","Infix","CopySource",
                           "NeedsRetesting",
                           "RetestingReason",'ALS Version']
            df_checks = df_checks.filter(items=checks_cols)
            checks_steps_cols = ['Study ID',"CheckName", "StepOrdinal", "CheckFunction", "StaticValue", "DataFormat", "VariableOID",
                                 "FolderOID", "FormOID", "FieldOID", "RecordPosition", "CustomFunction", "LogicalRecordPosition",
                                 "Scope", "OrderBy","FormRepeatNumber", "FolderRepeatNumber",'ALS Version']
            df_CheckSteps = df_CheckSteps.filter(items=checks_steps_cols)
            check_actions_cols = ['Study ID',"CheckName","FolderOID","FormOID","FieldOID","VariableOID",
                                  "RecordPosition","PageRepeatNumber",
                                  "InstanceRepeatNumber","LogicalRecordPosition","Scope","OrderBy","ActionType","ActionString",
                                  "ActionOptions","ActionScript",'ALS Version']
            df_CheckActions = df_CheckActions.filter(items=check_actions_cols)
            derv_cols = ['Study ID',"DerivationName","Active","FolderOID","FormOID","FieldOID","VariableOID",
                         "RecordPosition",
                         "AllVariablesInFolders","AllVariablesInFields","FormRepeatNumber","FolderRepeatNumber",
                         "BypassDuringMigration","CopySource","NeedsRetesting","RetestingReason",'ALS Version']
            df_Derivations = df_Derivations.filter(items=derv_cols)
            derv_steps_cols = ['Study ID',"DerivationName","StepOrdinal","DataFormat","VariableOID","StepValue",
                               "StepFunction","FolderOID",
                               "FormOID","FieldOID","CustomFunction","RecordPosition","LogicalRecordPosition","Scope","OrderBy",
                               "FormRepeatNumber","FolderRepeatNumber",'ALS Version']
            df_DerivationSteps = df_DerivationSteps.filter(items=derv_steps_cols)
            labvarmap_cols =['Study ID',"GlobalVariableOID","FormOID","FieldOID","FolderOID","LocationMethod",'ALS Version']
            df_LabVariableMappings = df_LabVariableMappings.filter(items=labvarmap_cols)
            cf_cols =['Study ID',"FunctionName","SourceCode","Lang",'ALS Version']
            df_CustomFunctions = df_CustomFunctions.filter(items=cf_cols)
            CoderConfiguration_cols =['Study ID','FormOID','FieldOID','CodingLevel','Priority','Locale',
                                      'IsApprovalRequired','IsAutoApproval','ALS Version']
            df_CoderConfiguration = df_CoderConfiguration.filter(items=CoderConfiguration_cols)
            CoderSupplementalTerms_cols =['Study ID','FormOID','FieldOID','SupplementalTerm','ALS Version']
            df_CoderSupplementalTerms = df_CoderSupplementalTerms.filter(items=CoderSupplementalTerms_cols)
        if (('Study_Summary' in para2)):
            TotalCRFscols =['Study ID','# of Total CRFs','ALS Version']
            df_masterdashboard = df_masterdashboard.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Visits']
            df_folders = df_folders.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Forms']
            df_forms= df_forms.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Checks']
            df_checks= df_checks.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Custom Functions']
            df_CustomFunctions= df_CustomFunctions.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Derivations']
            df_Derivations= df_Derivations.filter(items=TotalCRFscols)
            TotalCRFscols =['Study ID','# of Field Edit Checks']
            df_fields= df_fields.filter(items=TotalCRFscols)

            studyid = 'Study ID'
            df_masterdashboard = df_masterdashboard.merge(df_folders, on=studyid).merge(df_forms, on=studyid).merge(
                df_checks, on=studyid).merge(df_CustomFunctions, on=studyid).merge(df_Derivations, on=studyid).merge(
                df_fields, on=studyid)
            df_masterdashboard['Total No of Checks (Checks + Derivations + Field Edit Checks)'] = \
                df_masterdashboard['# of Checks'] + df_masterdashboard['# of Derivations'] +\
                df_masterdashboard['# of Field Edit Checks']


            df_masterdashboard = df_masterdashboard.merge(df_ibow, left_on=['Study ID'], right_on=['Study ID'], how='left',
                                          suffixes=['', '_'], indicator=True)


            mycols = set(df_masterdashboard.columns)
            mycols.remove('_merge')
            df_masterdashboard = df_masterdashboard[mycols]

            df_masterdashboard = df_masterdashboard.merge(df_idts, left_on=['Study ID'], right_on=['Study ID'], how='left',
                                          suffixes=['', '_'], indicator=True)

            df_masterdashboard = df_masterdashboard.rename(columns={ "Study ID":'Study'})
            TotalCRFscols = ['Study','ALS Version','Study Cluster', 'Study Status','DB Go-Live Date','# of Visits',
                             '# of Forms','# of Total CRFs','# of Checks','# of Custom Functions',
                             '# of Derivations','# of Field Edit Checks',
                             'Total No of Checks (Checks + Derivations + Field Edit Checks)', '# of DTS']
            df_masterdashboard = df_masterdashboard.filter(items=TotalCRFscols)

        print("reading alss")

        # df_list = []
        # print("df_CRFDraft bef",df_CRFDraft)
        # df_list.append(df_CRFDraft)
        # df_list.append(df_forms)
        # df_list.append(df_fields)
        # df_list.append(df_folders)
        # df_list.append(df_dicts)
        # df_list.append(df_Matrices)
        # df_list.append(df_checks)
        # df_list.append(df_CheckSteps)
        # df_list.append(df_CheckActions)
        # df_list.append(df_Derivations)
        # df_list.append(df_DerivationSteps)
        # df_list.append(df_LabVariableMappings)
        # df_list.append(df_CustomFunctions)
        # df_list.append(df_CoderConfiguration)
        # df_list.append(df_CoderSupplementalTerms)
        # df_list.append(df_LabVariableMappings)
        # print("df_list bef",df_list)
        #
        # return df_alltabs
        #---------------------------------------
        # df_forms = xcelread(files, 'Forms','A:O')
        # df_forms = df_forms.loc[:, ~df_forms.columns.str.contains('^Unnamed')]
        # df_forms = df_forms.fillna('')
        # df_fields = xcelread(files, 'Fields', 'A:AS')
        # df_fields = df_fields.loc[:, ~df_fields.columns.str.contains('^Unnamed')]
        # df_fields = df_fields.fillna('')
        # df_dicts = xcelread(files, 'DataDictionaries', 'A:A')
        # df_dicts = df_dicts.loc[:, ~df_dicts.columns.str.contains('^Unnamed')]
        # df_dicts = df_dicts.fillna('')
        # df_dict_entries = xcelread(files, 'DataDictionaryEntries', 'A:E')
        # df_dict_entries = df_dict_entries.loc[:, ~df_dict_entries.columns.str.contains('^Unnamed')]
        # df_dict_entries = df_dict_entries.fillna('')
        # df_checks = xcelread(files, 'Checks', 'A:G')
        # df_checks = df_checks.loc[:, ~df_checks.columns.str.contains('^Unnamed')]
        # df_checks = df_checks.fillna('')
        # df_CheckSteps = xcelread(files, 'CheckSteps', 'A:P')
        # df_CheckSteps = df_CheckSteps.loc[:, ~df_CheckSteps.columns.str.contains('^Unnamed')]
        # df_CheckSteps = df_CheckSteps.fillna('')
        # df_CheckActions = xcelread(files, 'CheckActions', 'A:O')
        # df_CheckActions = df_CheckActions.loc[:, ~df_CheckActions.columns.str.contains('^Unnamed')]
        # df_CheckActions = df_CheckActions.fillna('')
        # df_Derivations = xcelread(files, 'Derivations', 'A:O')
        # df_Derivations = df_Derivations.loc[:, ~df_Derivations.columns.str.contains('^Unnamed')]
        # df_Derivations = df_Derivations.fillna('')
        # df_DerivationSteps = xcelread(files, 'DerivationSteps', 'A:P')
        # df_DerivationSteps = df_DerivationSteps.loc[:, ~df_DerivationSteps.columns.str.contains('^Unnamed')]
        # df_DerivationSteps = df_DerivationSteps.fillna('')
        # df_LabVariableMappings = xcelread(files, 'LabVariableMappings', 'A:O')
        # df_LabVariableMappings = df_LabVariableMappings.loc[:, ~df_LabVariableMappings.columns.str.contains('^Unnamed')]
        # df_LabVariableMappings = df_LabVariableMappings.fillna('')
        # df_CustomFunctions = xcelread(files, 'CustomFunctions', 'A:O')
        # df_CustomFunctions = df_CustomFunctions.loc[:, ~df_CustomFunctions.columns.str.contains('^Unnamed')]
        # df_CustomFunctions = df_CustomFunctions.fillna('')
    #
    # #--------------------------------
    # rows_num = 1000

    rows_num = 900000
    print(" os.path.isfile(DATABASE", DATABASE)
    if not os.path.isfile(DATABASE):
        df = pd.DataFrame()
        df.to_excel(DATABASE)
    if (('Study_Summary' not in para2)):
        if (len(df_CheckSteps) <= rows_num) & (len(df_CheckSteps) >=0):
            df_CheckSteps_500 = df_CheckSteps.iloc[:rows_num + 1]
        if len(df_CheckSteps) > rows_num:
            df_CheckSteps_500 = df_CheckSteps.iloc[:rows_num + 1]
            df_CheckSteps_501 = df_CheckSteps.iloc[rows_num + 1: len(df_CheckSteps)]
        if len(df_alltabs) > 0:

            df_CRFDraft.replace( np.nan,'', inplace=True)
            df_forms.replace( np.nan,'', inplace=True)
            df_fields.replace( np.nan,'', inplace=True)
            df_folders.replace( np.nan,'', inplace=True)
            df_dicts.replace( np.nan,'', inplace=True)
            df_dict_entries.replace( np.nan,'', inplace=True)
            df_checks.replace( np.nan,'', inplace=True)
            df_CheckSteps_501.replace( np.nan,'', inplace=True)
            df_CheckSteps_500.replace( np.nan,'', inplace=True)
            df_CheckActions.replace( np.nan,'', inplace=True)
            df_Derivations.replace( np.nan,'', inplace=True)
            df_DerivationSteps.replace( np.nan,'', inplace=True)
            df_LabVariableMappings.replace( np.nan,'', inplace=True)
            df_CustomFunctions.replace( np.nan,'', inplace=True)
            df_Matrices.replace( np.nan,'', inplace=True)
            df_CoderConfiguration.replace( np.nan,'', inplace=True)
            df_CoderSupplementalTerms.replace( np.nan,'', inplace=True)

    wb =  pyexcelerate.Workbook()

        # if  para1 == 'Cons_Study':
        #     pyexecelerate_to_excel(wb, df_ALS_SP,sheet_name="Summary")
        # else:
        #     print("df_standard_als", df_standard_als)
        #     pyexecelerate_to_excel(wb, df_standard_als,sheet_name="Summary")
    if (('Study_Summary' not in para2)):
        pyexecelerate_to_excel(wb, df_CRFDraft,sheet_name="CRFDraft")
        pyexecelerate_to_excel(wb, df_forms,sheet_name="Forms")
        pyexecelerate_to_excel(wb, df_fields,sheet_name="Fields")
        pyexecelerate_to_excel(wb, df_folders,sheet_name="Folders")
        pyexecelerate_to_excel(wb, df_dicts,sheet_name="DataDictionaries")
        pyexecelerate_to_excel(wb, df_dict_entries,sheet_name="DataDictionaryEntries")
        pyexecelerate_to_excel(wb, df_Matrices,sheet_name="Matrices")
        pyexecelerate_to_excel(wb, df_checks,sheet_name="Checks")
        pyexecelerate_to_excel(wb, df_CheckSteps_500,sheet_name="CheckSteps")
        pyexecelerate_to_excel(wb, df_CheckSteps_501,sheet_name="CheckSteps_1")
        pyexecelerate_to_excel(wb, df_CheckActions,sheet_name="CheckActions")
        pyexecelerate_to_excel(wb, df_Derivations,sheet_name="Derivations")
        pyexecelerate_to_excel(wb, df_DerivationSteps,sheet_name="DerivationSteps")
        pyexecelerate_to_excel(wb, df_LabVariableMappings,sheet_name="LabVariableMappings")
        pyexecelerate_to_excel(wb, df_CustomFunctions,sheet_name="CustomFunctions")
        pyexecelerate_to_excel(wb, df_CoderConfiguration,sheet_name="CoderConfiguration")
        pyexecelerate_to_excel(wb, df_CoderSupplementalTerms,sheet_name="CoderSupplementalTerms")
    elif (('Study_Summary' in para2)):
        df_masterdashboard = df_masterdashboard.replace(np.nan, '')
        df_masterdashboard = df_masterdashboard.replace(np.nan, 0)
        df_masterdashboard.to_excel(r'C:\Bhasp\NVTSonco-work\NVTSonco-work\RADA\ALS\df_masterdashboard.xlsx')
        pyexecelerate_to_excel(wb, df_masterdashboard,sheet_name="Study Summary")

        # para2 = 'Cons_Study'
        # if para2 == 'Cons_Study':
        #     DATABASE1 = "Consolidated_ALS.xlsx"
        #     DATABASE = para3 + "\\" + DATABASE1
        #     wb.save(DATABASE)
        #     print("Consolidated_ALS is done DATABASE",DATABASE)
        #     print("Consolidated_ALS is done")
        #
        # else:
        #     DATABASE1 = "Consolidated_ALS_Global.xlsx"
        #     DATABASE = para3 +  "\\" + DATABASE1
        wb.save(DATABASE)
        print("Consolidated_ALS_Global is done")
    
    return DATABASE
    #
    #
    # #

    # with pd.ExcelWriter(r'C:\Users\pillibh2\OSSUM\Consolidated_ALS_Global\Consolidated_ALS.xlsx', engine='openpyxl', mode='a') as writer:
    #     if len(df_forms) > 0:
    #         print("reading alss4")
    #         sheetname = 'Forms'
    #         df_forms.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #

    # with pd.ExcelWriter(r'C:\Users\pillibh2\OSSUM\Consolidated_ALS_Global\Consolidated_ALS.xlsx', engine='openpyxl', mode='a') as writer:
    #     if len(df_forms) > 0:
    #         print("reading alss4")
    #         sheetname = 'Forms'
    #         df_forms.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #
    #     if len(df_fields) > 0:
    #         print("reading alss5")
    #         sheetname = 'Fields'
    #         df_fields.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_dicts) > 0:
    #         sheetname = 'DataDictionaries'
    #         df_dicts.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_dict_entries) > 0:
    #         sheetname = 'DataDictionaryEntries'
    #         df_dict_entries.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_checks) > 0:
    #         sheetname = 'Checks'
    #         df_checks.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #
    #     if (len(df_CheckSteps) <= rows_num) & (len(df_CheckSteps) >=0):
    #         print(df_CheckSteps)
    #         df_CheckSteps_500 = df_CheckSteps.iloc[:rows_num+1]
    #         sheetname = 'CheckSteps'
    #         df_CheckSteps_500.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #         #index test
    #     if len(df_CheckSteps) > rows_num:
    #         df_CheckSteps_500 = df_CheckSteps.iloc[:rows_num+1]
    #
    #         sheetname = 'CheckSteps'
    #         df_CheckSteps_500.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #
    #     if len(df_CheckSteps) > rows_num:
    #         df_CheckSteps_501 = df_CheckSteps.iloc[rows_num+1: len(df_CheckSteps)]
    #         sheetname = 'CheckSteps_1'
    #         df_CheckSteps_501.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_CheckActions) > 0:
    #         sheetname = 'CheckActions'
    #         df_CheckActions.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_Derivations) > 0:
    #         sheetname = 'Derivations'
    #         df_Derivations.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_DerivationSteps) > 0:
    #         sheetname = 'DerivationSteps'
    #         df_DerivationSteps.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_LabVariableMappings) > 0:
    #         sheetname = 'LabVariableMappings'
    #         df_LabVariableMappings.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #     if len(df_CustomFunctions) > 0:
    #         sheetname = 'CustomFunctions'
    #         df_CustomFunctions.to_excel(writer, sheet_name=sheetname, index=False)
    #         worksheet = writer.sheets[sheetname]
    #
    #     writer.save()

# als_cons_prog("Cons_Study")
# als_cons_prog("Cons_Gv")