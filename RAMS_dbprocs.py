def Rams_proc(para1, **kwargs):
    study_p = kwargs.get('study_p', '')
    subgroup_p = kwargs.get('subgroup_p', '')
    activity_p = kwargs.get('activity_p', '')
    Comments_p = kwargs.get('Comments_p', '')
    timespenthrs_p = kwargs.get('timespenthrs_p', '')
    timespentmins_p = kwargs.get('timespentmins_p', '')

    import sqlalchemy
    import sqlite3, os
    import pandas as pd
    import datetime, time
    import wmi
    from sqlite3 import Error

    msgbox_p = []
    # sqlite3 database connection
    # and further processing of data

    try:
        conn = sqlite3.connect("cddra_ram.db")
        cur = conn.cursor()
    except:
        pass

    def create_table(conn, create_table_sql):
        """ create a table from the create_table_sql statement
        :param conn: Connection object
        :param create_table_sql: a CREATE TABLE statement
        :return:
        """
        try:
            c = conn.cursor()
            c.execute(create_table_sql)
        except Error as e:
            print(e)

    def assoc_act_table(conn):

        sql_create_assoc_act_table = """ CREATE TABLE IF NOT EXISTS associate_activity_details (
                                                id integer PRIMARY KEY,
                                                study_name text NOT NULL,
                                                created_date text,
                                                created_by text,
                                                updated_date text,
                                                updated_by text,
                                            ); """
        conn = sqlite3.connect("cddra_ram.db")
        if conn is not None:
            create_table(conn, sql_create_assoc_act_table)
            print("Created table.")
        else:
            print("Error! cannot create the database connection.")


    def getHCode(activityName):
        """
        activityName
        """

        # ------------------------------------------------------------------------------------------------ ##
        query_hc = "SELECT [Horizon Code] from org_activity_list  WHERE Activity = ?"
        hcode = cur.execute(query_hc, (activityName,)).fetchall()
        # print(f"[addTimeEntry]: {type(activityName)} : activityName: {activityName}")
        # print(f"[addTimeEntry]: {type(orgName)} :  orgName: {orgName}")
        # print(f"[addTimeEntry]: {type(hcode)} :  hcode: {hcode}")
        print("Horizone code: " + hcode[0][0])
        return hcode[0][0]

    def addTimeEntry():
        """
        need to update
        """
        print("timespenthrss_p",timespenthrs_p)
        print("timespentmins_p",timespentmins_p)

        msgbox_p = []
        studyName = study_p
        orgName = subgroup_p
        activityName = activity_p
        # commentsText = self.activityComments.toPlainText()
        commentsText = Comments_p
        timeSpentHr = timespenthrs_p
        timeSpentMin = timespentmins_p  # need to look
        timeSpent = timeSpentHr + ":" + timeSpentMin
        ## additional columns details, which needs to be populated as part of data entry:
        ##  1.) created_date
        ##  2.) created_by
        ##  3.) updated_date
        ##  4.) updated_by
        now = datetime.datetime.now()
        createdDate = datetime.date.today()
        path = os.path.expanduser("~")
        listFolder = path.split("\\")
        systemOwner521ID = listFolder[2]
        ## -------------------------------------------------------------------------------------------------------- ##
        ## The WMI module can be used to gain system information of a windows machine
        # cObj = wmi.WMI()
        # my_system = cObj.Win32_ComputerSystem()[0]
        username = os.environ["USERNAME"]
        machineName = username
        # machineName = my_system.Name

        ## -------------------------------------------------------------------------------------------------------- ##
        try:
            horizonCode = getHCode(str(activity_p))
            print("Main : " + horizonCode)
            raise Exception('Hello')
        except Exception as error:
            print('Caught this error: ' + repr(error))
        ## -------------------------------------------------------------------------------------------------------- ##
        # print(type(timeSpent), timeSpent, type(int(timeSpentMin)), int(timeSpentMin))
        print("timeSpent",timeSpentMin)
        print("timeSpent",timeSpentHr)
        print("timeSpent",timeSpent)

        print("timeSpentHr 12345", timeSpentHr)
        print("timeSpentMin", timeSpentMin)
        if (
                (int(timeSpentMin) > 0 and int(timeSpentHr) == 0)
                or (int(timeSpentMin) == 0 and int(timeSpentHr) > 0)
                or (int(timeSpentMin) > 0 and int(timeSpentHr) > 0)
        ):
            try:
                print("Inside if : ", timeSpentMin)
                searchQuery = (
                    "SELECT study_name, activity_name, created_date "
                    " FROM associate_activity_details WHERE study_name = ? AND activity_name = ? AND created_date =?"
                )
                statusCheck = cur.execute(
                    searchQuery, (studyName, activityName, createdDate)
                ).fetchall()
                print("statusCheck",statusCheck)
                # print("-------------------------------------------------------")
                if statusCheck == []:

                    print("Inside if : statusCheck", statusCheck)
                    # if not statusCheck:
                    # # ---------------------------------- For the first time entry/fresh entry
                    # ------------------------------------------ ##
                    try:
                        with conn:

                            print("Ranjeet Test Here")
                            query = """INSERT INTO 'associate_activity_details' 
                                        (
                                            study_name, org_name, activity_name, comments, time_spent_minutes, 
                                            machine_name, created_date, created_by, is_latest, horizon_code
                                        ) 
                                        VALUES
                                        (
                                            ?,?,?,?,?,?,?,?,?,?
                                        )"""
                            cur.execute(
                                query,
                                (
                                    studyName,
                                    orgName,
                                    activityName,
                                    commentsText,
                                    timeSpent,
                                    machineName,
                                    createdDate,
                                    systemOwner521ID,
                                    "YES",
                                    horizonCode,
                                ),
                            )
                            conn.commit()

                        msgbox_p.append("Information, Entry has been added successfully")
                        print("Information", "Entry has been added successfully")

                    except Exception as e:
                        print(e)
                        msgbox_p.append("Warning, Entry has not been added into database")
                        print("Warning", "Entry has not been added into database")

                    msgbox_p.append("Information, Entry has been added successfully")
                else:
                    ## ---- If there is time entry already present into DB ------------------------------------------ ##
                    try:
                        print("Ranjeet Test Here [Updated Version]")
                        updateQuery = """UPDATE associate_activity_details SET is_latest = 'NO' 
                                         WHERE study_name = ? AND activity_name = ? AND created_date =? """
                        cur.execute(updateQuery, (studyName, activityName, createdDate))

                        query = """INSERT INTO 'associate_activity_details' 
                                        (
                                            study_name, org_name, activity_name, comments, time_spent_minutes, 
                                            machine_name, created_date, created_by, updated_date,
                                            updated_by, is_latest, horizon_code
                                        ) 
                                        VALUES
                                        (
                                            ?,?,?,?,?,?,?,?,?,?,?,?
                                        )"""
                        cur.execute(
                            query,
                            (
                                studyName,
                                orgName,
                                activityName,
                                commentsText,
                                timeSpent,
                                machineName,
                                createdDate,
                                systemOwner521ID,
                                now,
                                systemOwner521ID,
                                "YES",
                                horizonCode,
                            ),
                        )

                        conn.commit()

                        msgbox_p.append("Information, Entry has been added successfully")
                        print("Information", "Entry has been added successfully")

                    except Exception as e:
                        print(e)
                        msgbox_p.append("Warning, Entry has not been added into database")
                        print("Warning", "Entry has not been added into database")

                    msgbox_p.append("Information, Entry has been added successfully")
            except Exception as e:
                print(e)
                print("Hello: Error is there")
            msgbox_p.append("Information, Entry has been added successfully")
        else:
            print("Warning ,time spent is not valid\n" "select a valid time for the activities.")
            msgbox_p.append("Warning, time spent is not valid\n" "select a valid time for the activities.")

        return msgbox_p



    def study_fetch():
        ## study combo implementation and data population from DB
        studyQuery = "SELECT [index], Trial FROM study"
        print("studyQuery",type(studyQuery))
        studyComboData = cur.execute(
            studyQuery,
        ).fetchall()
        # for study in studyComboData:
        #     # print(study[1])
        #     self.studyCombo.addItem(study[1], study[0])
        return studyComboData

    def subgroup_act_fetch():
        ## study combo implementation and data population from DB
        # subgrpQuery = "SELECT [index], Activity, [Horizon Code],[Sub Group] FROM org_activity_list WHERE [Sub Group] = 'CDD'"
        subgrpQuery = "SELECT [index], Activity, [Horizon Code],[Sub Group] FROM org_activity_list "

        ActivityComboData = cur.execute(
            subgrpQuery,
        ).fetchall()
        print("ActivityComboData", type(ActivityComboData))
        # subgrpQuery2 = "SELECT [index], Activity, [Horizon Code],[Sub Group] FROM org_ndp_list WHERE [Sub Group] = 'CDD'"
        subgrpQuery2 = "SELECT [index], Activity, [Horizon Code],[Sub Group] FROM org_ndp_list "
        # subgrpQuery2 = "SELECT * FROM org_ndp_list "


        ActivityComboData2 = cur.execute(
            subgrpQuery2,
        ).fetchall()
        print("ActivityComboData2", type(ActivityComboData2))
        print("ActivityComboData", (ActivityComboData))
        print("ActivityComboData2", (ActivityComboData2))
        df1 = pd.DataFrame(ActivityComboData, columns=["index", "Activity", "Horizon Code","Sub Group"])
        df2 = pd.DataFrame(ActivityComboData2, columns=["index", "Activity", "Horizon Code","Sub Group"])
        df = pd.DataFrame()
        df = df.append(df1)
        df = df.append(df2)
        ActivityComboData.append(ActivityComboData2)

        print("ActivityComboData", type(ActivityComboData))
        print("ActivityComboData", (ActivityComboData))

        # for study in studyComboData:
        #     # print(study[1])
        #     self.studyCombo.addItem(study[1], study[0])
        return df

    def displayData():
        """
        need to update
        """
        queryAssociateActivityDetails = cur.execute(
            """SELECT id, study_name, activity_name, comments, time_spent_minutes, created_date, created_by FROM 
            associate_activity_details WHERE is_latest = ? order by created_date desc""",
            ("YES",),
        )
        df1 = pd.DataFrame(queryAssociateActivityDetails, columns=["id", "study_name", "activity_name", "Horizon Code",
                                                                   "time_spent_minutes","comments",
                                                                   "created_date","created_by","Sub Group"] )
        return df1

    def call_func(queryset):
        studies = []
        studieslist = queryset

        for study in studieslist:
            studies.append(study[1])

        return studies

    assoc_act_table(conn)

    df_list = []
    if para1 == 'pre_rams':
        print("pre_rams")
        studies = []
        subgroups = []
        studies = call_func(study_fetch())
        subgroups = subgroup_act_fetch()
        # browsedata = displayData()
        # subgroups = call_func(subgroup_act_fetch())
        print("subgroups",subgroups)
        df_list.append(studies)
        df_list.append(subgroups)
        # df_list.append(browsedata)

    elif para1 == 'post_rams':
        print("acbcddd")
        print("post_rams")
        studies = []
        subgroups = []
        studies = study_fetch()
        subgroups = subgroup_act_fetch()
        # subgroups = call_func(subgroup_act_fetch())
        print("subgroups", subgroups)
        msg = addTimeEntry()
        print("msg t",type(msg))
        print("msg",msg)
        # browsedata = displayData()
        df_list.append(studies)
        df_list.append(subgroups)
        df_list.append(msg)
        # df_list.append(browsedata)


    return df_list
    # if para1 == 'pre_rams':
    #     studies = []
    #     studieslist = study_fetch()
    # 
    #     for study in studieslist:
    #         studies.append(study[1])
    # 
    #     return studies
    # 
    # 
    # if para1 == 'studylist':
    #     studies = []
    #     studieslist = study_fetch()
    # 
    #     for study in studieslist:
    #         studies.append(study[1])
    # 
    #     return studies
    # 



