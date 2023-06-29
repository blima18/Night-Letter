import sqlite3
from sqlite3 import Error
import base64

column_names = "FORM_ID, lead_plant, status, subject, approved, affected, created_by, issue_date, status_date, as_approved_by, bsaq, wers_alert, global, additional_vehicles, clean_point_dates, units_affected, supplier, date_repair_parts, root_cause_owner, repair_funding, date_repair_instructions, background, root_cause, interim, permanent, next_steps, new_issues, help_required, additional, other"

def create_connection(db_file):
    """ create a database connection to the SQLite database
        specified by db_file
    :param db_file: database file
    :return: Connection object or None
    """
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return conn

def run_update_case(values, targuet_id):
    list_values = values.split(",\"")
    result =    "UPDATE FORM " \
                "SET lead_plant = \"" + list_values[1] + ", " \
                "status = \"" + list_values[2] + ", " \
                "subject = \"" + list_values[3] + ", " \
                "approved = \"" + list_values[4] + ", " \
                "affected = \"" + list_values[5] + ", " \
                "created_by = \"" + list_values[6] + ", " \
                "issue_date = \"" + list_values[7] + ", " \
                "status_date = \"" + list_values[8] + ", " \
                "as_approved_by = \"" + list_values[9] + ", " \
                "bsaq = \"" + list_values[10] + ", " \
                "wers_alert = \"" + list_values[11] + ", " \
                "global = \"" + list_values[12] + ", " \
                "additional_vehicles = \"" + list_values[13] + ", " \
                "clean_point_dates = \"" + list_values[14] + ", " \
                "units_affected = \"" + list_values[15] + ", " \
                "supplier = \"" + list_values[16] + ", " \
                "date_repair_parts = \"" + list_values[17] + ", " \
                "root_cause_owner = \"" + list_values[18] + ", " \
                "repair_funding = \"" + list_values[19] + ", " \
                "date_repair_instructions = \"" + list_values[20] + ", " \
                "background = \"" + list_values[21] + ", " \
                "root_cause = \"" + list_values[22] + ", " \
                "interim = \"" + list_values[23] + ", " \
                "permanent = \"" + list_values[24] + ", " \
                "next_steps = \"" + list_values[25] + ", " \
                "new_issues = \"" + list_values[26] + ", " \
                "help_required = \"" + list_values[27] + ", " \
                "additional = \"" + list_values[28] + ", " \
                "other = \"" + list_values[29] + " " \
                "WHERE FORM_ID = " + str(targuet_id)
    return result

def convertToBinaryData(file):
    blobData = file.read()
    return blobData

def select_content(request):
    form = request.form
    database = "StopShipForm"
    conn = create_connection(database)

    cur = conn.cursor()
    cur.execute("SELECT * FROM FORM WHERE subject = \"" + form["Subject:"] + "\"")
    rows = cur.fetchall()
    update_case = False
    if len(rows) > 0:
        target_id = rows[0][0]
        update_case = True
    else:
        cur.execute("SELECT * FROM FORM")
        target_id = len(cur.fetchall()) + 1

    plant_variables_fields = ["Plant(s) Clean Point Date(s):_", "Date Repair Parts Available:_", "Date Repair Instructions Avail:_",
                              "_suspect", "_ok", "_nok", "_repaired", "_notes", "_dtg"]

    plants = {}
    values = str(target_id) + ","
    for field in form:
        found = False
        for a in plant_variables_fields:
            if a in field:
                found = True
        
        if found is False:
            values += "\"" + form[field] + "\"" + ","
        else:
            if "suspect" in field or "ok" in field or "nok" in field or "repaired" in field or "notes" in field or "dtg" in field:
                index = 0
            else:
                index = 1
            try:
                plants[field.split("_")[index]] += "\"" + form[field] + "\"" + ","
            except:
                plants[field.split("_")[index]] = "\"" + form[field] + "\"" + ","
    
    cur.execute("SELECT * FROM PLANT")
    rows = cur.fetchall()
    lenght = len(rows) + 1
    for x, plant in enumerate(plants):
        plants[plant] = str(lenght + x) + "," + str(target_id) + ",\"" + plant + "\"," + str(1 if x == 0 else 0) + "," + plants[plant][:len(plants[plant])-1]

    values = values[:len(values)-1]
    
    if not update_case:
        conn.cursor().execute("INSERT INTO FORM (" + column_names + " )" +
                          "VALUES (" + values + ")")
    else:
        conn.cursor().execute(run_update_case(values, target_id)) 

    for plant in plants:
        plant_values = plants[plant].split(",")
        cur.execute("SELECT * FROM PLANT WHERE FORM_ID = " + str(target_id) + " AND NAME = " + str(plant_values[2]))
        a = cur.fetchall()
        if len(a) > 0:
            result =    "UPDATE PLANT " \
                "SET DATE_1 = " + plant_values[4] + ", " \
                "DATE_2 = " + plant_values[5] + ", " \
                "DATE_3 = " + plant_values[6] + ", " \
                "SUSPECTED_VALUE = " + plant_values[7] + ", " \
                "NOK_VALUE = " + plant_values[8] + ", " \
                "OK_VALUE = " + plant_values[9] + ", " \
                "REPAIRED_VALUE = " + plant_values[10] + ", " \
                "DTG_VALUE = " + plant_values[11] + ", " \
                "NOTES_VALUE = " + plant_values[12] + " " \
                "WHERE FORM_ID = " + str(target_id) + " AND NAME = " + str(plant_values[2])
            conn.cursor().execute(result) 
        else:
            conn.cursor().execute("INSERT INTO PLANT (PLANT_ID, FORM_ID, NAME, MAIN_PLANT, DATE_1, DATE_2, DATE_3, SUSPECTED_VALUE, NOK_VALUE, OK_VALUE, REPAIRED_VALUE, NOTES_VALUE, DTG_VALUE)" +
                                "VALUES (" + plants[plant] + ")")

    cur.execute("SELECT * FROM PLANT WHERE FORM_ID = " + str(target_id))
    rows = cur.fetchall()
    for plant in rows:
        if plant[2] not in plants.keys():
            cur.execute("DELETE FROM PLANT WHERE FORM_ID = " + str(target_id) + " AND NAME = \"" + str(plant[2]) + "\"")
    
    cur.execute("SELECT * FROM IMAGE WHERE FORM_ID = " + str(target_id))
    rows = cur.fetchall()

    for file in request.files:
        found = False
        for r in rows:
            if r[3] in file:
                found = True
        
        if found is False:
            sqlite_insert_blob_query = """ INSERT INTO IMAGE
                                  (IMAGE_ID, FORM_ID, BLOB_IMAGE, target) VALUES (?, ?, ?, ?)"""

            image_value = (str(len(rows) + 1), str(target_id), convertToBinaryData(request.files[file]), file)
        else:
            sqlite_insert_blob_query = """ UPDATE IMAGE
                                  SET BLOB_IMAGE=? WHERE target=?"""

            image_value = (convertToBinaryData(request.files[file]), file)

        conn.cursor().execute(sqlite_insert_blob_query, image_value)
    for image in rows:
        if image[3] not in request.files.keys():
            cur.execute("DELETE FROM IMAGE WHERE FORM_ID = " + str(target_id) + " AND IMAGE_ID = \"" + str(image[0]) + "\"")

    conn.commit()

def check_in_list(target, array):
    for a in array:
        if a in target:
            return True
    return False

def convert_to_csv_structure():
    database = "StopShipForm"
    conn = create_connection(database)
    cur = conn.cursor()
    cur.execute("SELECT * FROM FORM")
    forms = cur.fetchall()
    cur.execute("SELECT * FROM PLANT")
    plants = cur.fetchall()
    cur.execute("SELECT * FROM IMAGE")
    images = cur.fetchall()
    
    plant_variables_fields = ["Plant(s) Clean Point Date(s):_", "Date Repair Parts Available:_", "Date Repair Instructions Avail:_",
                              "_suspect", "_ok", "_nok", "_repaired", "_notes", "_dtg"]

    dict_sections = {"1": "Lead Plant:",
                     "2": "Status:",
                     "3": "Subject:",
                     "4": "Approved:",
                     "5": "Plant(s) Affected:",
                     "6": "Created by:",
                     "7": "Issue Date:",
                     "8": "Status Date:",
                     "9": "As Approved By:",
                     "10": "BSAQ#:",
                     "11": "WERS Alert#(s):",
                     "12": "Global 8D or 14D #",
                     "13": "Additional Vehicles Affected:",
                     "14": "Plant(s) Clean Point Date(s):",
                     "15": "Plant(s) # of Units Affected:",
                     "16": "Supplier Name/Part #(s):",
                     "17": "Date Repair Parts Available:",
                     "18": "Root Cause Owner/Org:",
                     "19": "Repair Funding Source:",
                     "20": "Date Repair Instructions Avail:",
                     "21": "Background/Concern Description:",
                     "22": "Root Cause:",
                     "23": "Interim Corrective Action (ICA):",
                     "24": "Permanent Corrective Action (PCA):",
                     "25": "Status/Next Steps:",
                     "26": "New Issues/Updates:",
                     "27": "Help Required:",
                     "28": "Additional On-Site Support:",
                     "29": "Other:"}
    
    dict_plant_sections = {"4": "Plant(s) Clean Point Date(s):_",
                           "5": "Date Repair Parts Available:_",
                           "6": "Date Repair Instructions Avail:_",
                           "7": "_suspect",
                           "8": "_nok",
                           "9": "_ok",
                           "10": "_repaired",
                           "11": "_dtg",
                           "12": "_notes"}

    result = []
    result_fields = []
    for data in forms:
        value = []
        for x, d in enumerate(data):
            if x > 0:
                value.append(d)
                if dict_sections[str(x)] not in result_fields:
                    result_fields.append(dict_sections[str(x)])
    
        for plant in plants:
            for y, p in enumerate(plant):
                if y >= 4 and plant[1] == data[0]:
                    value.append(p)
                    if "suspect" in dict_plant_sections[str(y)] or "ok" in dict_plant_sections[str(y)] or "nok" in dict_plant_sections[str(y)] or "repaired" in dict_plant_sections[str(y)] or "dtg" in dict_plant_sections[str(y)] or "notes" in dict_plant_sections[str(y)]:
                        if (plant[2] + dict_plant_sections[str(y)]) not in result_fields:
                            result_fields.append(plant[2] + dict_plant_sections[str(y)])
                    else:
                        if (dict_plant_sections[str(y)] + plant[2]) not in result_fields:
                            result_fields.append(dict_plant_sections[str(y)] + plant[2])

        for image in images:
            if image[1] == data[0]:
                value.append(str(base64.b64encode(image[2])))
                if image[3] not in result_fields:
                    result_fields.append(image[3])

        result.append(value.copy())
                    
                    
    return result, result_fields
