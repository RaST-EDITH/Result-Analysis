import os
import pandas as pd

def analysResult4() :

    df = pd.read_excel( data["File"] )

    row, col = df.shape
    row = row - 3
    column = df.columns
    col_name = []
    unnamed_col = []

    for i in column :
        if "Unnamed" not in i :
            col_name.append( i )
        else :
            unnamed_col.append( i )

    col_name = col_name[4:len(col_name)-4]

    facl_name = []
    try :
        for i in range( len(col_name) ) :
            facl = df[col_name[i]][0].split('/')
            if ( len(facl) == 1) :
                facl_name.extend(facl*4)
            elif ( len(facl) == 2) :
                facl_name.extend(facl*2)
            elif ( len(facl) == 3) :
                facl_name.extend(facl + [""])
            elif ( len(facl) == 4) :
                facl_name.extend(facl)
    except :
        for i in range( len(col_name) ) :
            facl_name.extend([df[col_name[i]][0]]*4)

    max_mark = []
    for i in range( len(col_name) ) :
        mark = df[col_name[i]][2]+df[unnamed_col[i]][2]
        max_mark.extend([ mark]*4)

    branch = data["Branch"]*len(col_name)

    temp = []
    for i in col_name :
        temp.extend([ i, "", "", ""])
    col_name = temp[:]

    temp = []
    for i in unnamed_col :
        temp.extend([ i, "", "", ""])
    unnamed_col = temp[:]

    sheet_structure = {
        "Subject" : col_name,
        "Faculty Name" : facl_name,
        "Branch" : branch,
        "Number of Students" : [i for i in range(len(col_name))],
        "Absent" : [i for i in range(len(col_name))],
        "Pass" : [i for i in range(len(col_name))],
        "Less than 60%" : [i for i in range(len(col_name))],
        "Between 60 to 74%" : [i for i in range(len(col_name))],
        "More than 75%" : [i for i in range(len(col_name))],
        "Maximum Score" : [i for i in range(len(col_name))],
        "Out of Mark" : max_mark,
        "Pass Percentage" : [i for i in range(len(col_name))],
    }

analysResult4()
