import pandas as pd
from docx import Document
def infostrip (filepath):
    filepath = "JIRA.csv"
    a = pd.read_csv(filepath, sep=',')
    #a = a.replace('Fixed', 'Passed')
    todrop =['Issue Type','Issue id', 'Parent id']
    b = a.drop(todrop, axis=1)
    #print b
    return b

def maketable (do, table ):
    #do.reset_index(drop=True)
    name = do["Issue key"]
    sumer = do["Summary"]
    status = do["Status"]
    passes = do["Resolution"]
    count = 0
    for item in do["Issue key"]:
       # if do.iloc(count) == True:
            row_cells = table.add_row().cells
            row_cells[0].text = str(name[count])
            row_cells[1].text = str(sumer[count])
            row_cells[2].text = str(status[count])
            row_cells[3].text = str(passes[count])
            count = count +1

def makefiltertable(document, do, keyword):

    table = document.add_table(rows=1, cols=4)
    t1c = table.rows[0].cells
    t1c[0].text = "Name"
    t1c[1].text = "Summary"
    t1c[2].text = "Status"
    t1c[3].text = "Pass"

    name = do["Issue key"]
    sumer = do["Summary"]
    status = do["Status"]
    passes = do["Resolution"]
    count = 0
    for item in do["Status"]:
        if item == keyword:
            #print  item
            row_cells = table.add_row().cells
            row_cells[0].text = str(name[count])
            row_cells[1].text = str(sumer[count])
            row_cells[2].text = str(status[count])
            row_cells[3].text = str(passes[count])
        count = count +1