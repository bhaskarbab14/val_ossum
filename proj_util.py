"""
 THIS FILE CONTAINS:
 - ALL THE UTILITIES CODE
 - FUNCTIONS
 - ANY OTHER IMPORTANT CHUNC OF CODE
"""
import pandas as pd
import sqlalchemy
import sqlite3

# sqlite3 database connection
# and further processing of data
conn = sqlite3.connect("cddra_ram.db")
cur = conn.cursor()
# engine creation using sqlalchemy
engine = sqlalchemy.create_engine("sqlite:///cddra_ram.db", echo=True)

df_ndps = pd.read_excel(r"NDPs_SP.xlsx")
# print(data_ndps.head(5))
df_activity_list = pd.read_excel("ABCD_list.xlsx")
df_activity_list = df_activity_list[["Sub Group", "Activity", "Horizon Code"]]


# FUNCTION to load the activity table
def load_activity_table():
    """
    load_activity_table
    """
    try:
        with conn:
            df_activity_list.to_sql(
                "org_activity_list", engine, if_exists="replace", index=True
            )
            # QMessageBox.information(
            #     self, "Information", "activity table refreshed with latest data"
            # )
            print("activity table refreshed with latest data")
            # query_hc = """SELECT [Horizon Code] from org_activity_list
            #                WHERE Activity= 'Full Outsourced : Maintenance' and [Sub Group] = 'CDD'"""
            # hcode = cur.execute(query_hc).fetchone()
            # print(f"Horizon Code: {hcode[0]}")
    except:
        # QMessageBox.information(
        #     self, "Information", "Something wrong with the database activities"
        # )
        print("Something wrong with the database activities")
