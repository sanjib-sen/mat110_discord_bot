from openpyxl import load_workbook
quiz_1 = load_workbook("MAT110API.xlsx")["Quiz 1"]
def quiz1_mark(id, email=None, bux = None):
    for row in range(2, 1100):
        id_char  = "A"+str(row)
        name_char  = "B"+str(row)
        email_char  = "C"+str(row)
        bux_char  = "D"+str(row)
        id_api = str(quiz_1[id_char].value)
        name = str(quiz_1[name_char].value)
        email = str(quiz_1[email_char].value)
        bux = str(quiz_1[bux_char].value)
        if str(id) == id_api:
            mark_char  = "E"+str(row)
            mark = str(quiz_1[mark_char].value)
            return mark
    
    return "Not Found."

