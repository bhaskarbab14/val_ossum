def Rams_uploads(para1, para2,**kwargs):
    para3 = kwargs.get('para3', '')
    import sqlalchemy
    import sqlite3, os
    # from proj_util import (df_ndps,
    #                        load_activity_table,
    #                        conn,
    #                        cur,
    #                        engine)
    import pandas as pd
    import datetime, time
    import wmi


    ##################################
    # FUNCTION to load the studylist table
    def load_studylist_table(para1):
        df_ibowfile = pd.read_excel(para1)
        df_trial_name = df_ibowfile["Trial"]
        # load_studylist_table
        try:
            with conn:
                df_trial_name.to_sql(
                    "study", engine, if_exists="replace", index=True
                )

                print("studies list is refreshed")
        except:
            print("Something wrong with the studies list")

            # FUNCTION to load the activity table

    def load_activity_table(para1):

        df_activity_list = pd.read_excel(para1)
        # df_ndps_list = pd.read_excel(para3)
        df_activity_list = df_activity_list [["Sub Group", "Activity", "Horizon Code"]]
        # df_ndps_list = df_activity_list1[["Sub Group", "Activity", "Horizon Code"]]
        # df_activity_list = pd.DataFrame()
        # df_activity_list.append(df_activity_list1)
        # df_activity_list.append(df_ndps_list)
        """
        load_activity_table
        """
        try:
            with conn:
                df_activity_list.to_sql(
                    "org_activity_list", engine, if_exists="replace", index=True
                )
                print("activity table refreshed with latest data")
        except:
                print("Something wrong with the database activities")

    column_names = ["id", "study_name", "activity_name", "Horizon Code",
                    "time_spent_minutes", "comments",
                    "created_date", "created_by", "Sub Group"]

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
            with conn:
                try:
                    create_table(conn, sql_create_assoc_act_table)
                except Error as err:
                    print(err)
            print("Created table.")
        else:
            print("Error! cannot create the database connection.")




    def load_ndp_table(para1):

            # df_activity_list1 = pd.read_excel(para1)
            df_ndps_list = pd.read_excel(para1)
            print("df_ndps_list", df_ndps_list)
            # df_activity_list1 = df_activity_list1[["Sub Group", "Activity", "Horizon Code"]]
            df_ndps_list = df_ndps_list.rename(columns={"Title": "Activity"})
            df_ndps_list["Sub Group"] = 'NDPs'
            df_ndps_list = df_ndps_list[["Sub Group", "Activity", "Horizon Code"]]
            print("df_ndps_list", df_ndps_list)
            df_ndps_list_COE = df_ndps_list.loc[~df_ndps_list['Activity'].isna()]
            df_ndps_list["Sub Group"] = 'COE'
            print("df_ndps_list coe", df_ndps_list_COE)
            df_ndps_list_CDD = df_ndps_list.loc[~df_ndps_list['Activity'].isna()]
            df_ndps_list["Sub Group"] = 'CDD'
            df_ndps_list_DAP = df_ndps_list.loc[~df_ndps_list['Activity'].isna()]
            df_ndps_list["Sub Group"] = 'DAP'
            df_ndps_list_DRRA = df_ndps_list.loc[~df_ndps_list['Activity'].isna()]
            df_ndps_list["Sub Group"] = 'DRA'

            df_ndps_list1 = pd.DataFrame()
            df_ndps_list1 = df_ndps_list1.append(df_ndps_list_COE)
            df_ndps_list1 = df_ndps_list1.append(df_ndps_list_CDD)
            df_ndps_list1 = df_ndps_list1.append(df_ndps_list_DAP)
            df_ndps_list1 = df_ndps_list1.append(df_ndps_list_DRRA)

            print(len(df_ndps_list1))
            print("df_ndps_list",df_ndps_list1)
            """
            load_activity_table
            """
            try:
                with conn:
                    df_ndps_list1.to_sql(
                        "org_ndp_list", engine, if_exists="replace", index=True
                    )

                    print("ndp table refreshed with latest data")

            except:

                    print("Something wrong with the ndp activities")


    ####################################

    from sqlite3 import Error

    """ create a database connection to a database that resides
        in the memory
    """
    conn = None;
    try:
        conn = sqlite3.connect("cddra_ram.db")
        cur = conn.cursor()
        # engine creation using sqlalchemy
        engine = sqlalchemy.create_engine("sqlite:///cddra_ram.db", echo=True)
    except Error as e:
        print(e)
    finally:
        if conn:
            conn.close()
    #
    # df_ndps = pd.read_excel(para1)
    # # print(data_ndps.head(5))
    # df_activity_list = pd.read_excel(para2)
    # df_activity_list = df_activity_list[["Sub Group", "Activity", "Horizon Code"]]

    assoc_act_table(conn)

    if para2 == 'ibow':
        load_studylist_table(para1)

    if para2 == 'abcd':
        load_activity_table(para1)

    if para2 == 'ndp':
        load_ndp_table(para1)


