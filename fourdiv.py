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

    a1 = data["Start"][0]
    a2 = data["Start"][1]
    a3 = data["Start"][2]
    b1 = data["End"][0]
    b2 = data["End"][1]
    b3 = data["End"][2]
    spec1 = data["Spec"][0]
    spec2 = data["Spec"][1]
    spec3 = data["Spec"][2]

    Branch_1 =  ( df[df.columns[1]][3:] >= a1 ) & ( df[df.columns[1]][3:] <= b1 )
    Branch_2 =  ( df[df.columns[1]][3:] >= a2 ) & ( df[df.columns[1]][3:] <= b2 )
    Branch_3 =  ( df[df.columns[1]][3:] >= a3 ) & ( df[df.columns[1]][3:] <= b3 )

    if len(spec1) > 0 :
        for x in spec :
            sp = df[df.columns[1]][3:] == x
            Branch_1 = Branch_1 | sp

    if len(spec2) > 0 :
        for x in spec :
            sp = df[df.columns[1]][3:] == x
            Branch_2 = Branch_2 | sp
    
    if len(spec3) > 0 :
        for x in spec :
            sp = df[df.columns[1]][3:] == x
            Branch_3 = Branch_3 | sp

    for i in range( 0, len(col_name), 4 ) :

        # Student Count
        student_count = df[col_name[i]][3:] >= 0
        student_total = student_count.value_counts()[1]
        student_count_1 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_1
        student_count_1 = dict(student_count_1.value_counts())
        count1 = 0
        if True in student_count_1.keys() :
            count1 = student_count_1[True]
        
        student_count_2 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_2
        student_count_2 = dict(student_count_2.value_counts())
        count2 = 0
        if True in student_count_2.keys() :
            count2 = student_count_2[True]
        
        student_count_3 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_3
        student_count_3 = dict(student_count_3.value_counts())
        count3 = 0
        if True in student_count_3.keys() :
            count3 = student_count_3[True]
        
        sheet_structure["Number of Students"][i] = count1
        sheet_structure["Number of Students"][i+1] = count2
        sheet_structure["Number of Students"][i+2] = count3
        sheet_structure["Number of Students"][i+3] = student_total - count1 - count2 - count3


analysResult4()
