import os
import pandas as pd

def analysResult3() :

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
                facl_name.extend(facl*3)
            elif ( len(facl) == 2) :
                facl_name.extend(facl+[""])
            elif ( len(facl) == 3) :
                facl_name.extend(facl)
    except :
        for i in range( len(col_name) ) :
            facl_name.extend([df[col_name[i]][0]]*3)

    max_mark = []
    for i in range( len(col_name) ) :
        mark = df[col_name[i]][2]+df[unnamed_col[i]][2]
        max_mark.extend([ mark]*3)

    branch = data["Branch"]*len(col_name)

    temp = []
    for i in col_name :
        temp.extend([ i, "", ""])
    col_name = temp[:]

    temp = []
    for i in unnamed_col :
        temp.extend([ i, "", ""])
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
    b1 = data["End"][0]
    b2 = data["End"][1]
    spec1 = data["Spec"][0]
    spec2 = data["Spec"][1]

    Branch_1 =  ( df[df.columns[1]][3:] >= a1 ) & ( df[df.columns[1]][3:] <= b1 )
    Branch_2 =  ( df[df.columns[1]][3:] >= a2 ) & ( df[df.columns[1]][3:] <= b2 )

    if len(spec1) > 0 :
        for x in spec :
            sp = df[df.columns[1]][3:] == x
            Branch_1 = Branch_1 | sp

    if len(spec2) > 0 :
        for x in spec :
            sp = df[df.columns[1]][3:] == x
            Branch_2 = Branch_2 | sp

    for i in range( 0, len(col_name), 3 ) :

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
        count2student_count_2 = 0
        if True in student_count_2.keys() :
            count2 = student_count_2[True]
        
        sheet_structure["Number of Students"][i] = count1
        sheet_structure["Number of Students"][i+1] = count2
        sheet_structure["Number of Students"][i+2] = student_total - count1 - count2


        # Absent Count
        abs_count = df[col_name[i]][3:] == 0
        count = dict(abs_count.value_counts())
        abst = 0
        if True in count.keys() :
            abst = count[True]

        abs_count = abs_count & Branch_1
        count1 = dict(abs_count.value_counts())
        abst1 = 0
        if True in count1.keys() :
            abst1 = count1[True]
        
        abs_count = abs_count & Branch_2
        count2 = dict(abs_count.value_counts())
        abst2 = 0
        if True in count2.keys() :
            abst2 = count2[True]
        sheet_structure["Absent"][i] = abst1
        sheet_structure["Absent"][i+1] = abst2
        sheet_structure["Absent"][i+2] = abst - abst1 - abst2
        

        # Less than 60
        less_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) <= int(max_mark[i]*0.6)
        val = dict(less_sixty.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0
        
        less_sixty = less_sixty & Branch_1
        val1 = dict(less_sixty.value_counts())
        if True in val1.keys() :
            val1 = val1[True]
        else :
            val1 = 0
        
        less_sixty = less_sixty & Branch_2
        val2 = dict(less_sixty.value_counts())
        if True in val2.keys() :
            val2 = val2[True]
        else :
            val2 = 0
        sheet_structure["Less than 60%"][i] = val1
        sheet_structure["Less than 60%"][i+1] = val2
        sheet_structure["Less than 60%"][i+2] = val - val1 - val2


        # Between 60 and 75
        sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) > int(max_mark[i]*0.6)
        seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) < int(max_mark[i]*0.75)
        btw_sixty_seventy = sixty & seventy
        val = dict(btw_sixty_seventy.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0

        btw_sixty_seventy = btw_sixty_seventy & Branch_1
        val1 = dict(btw_sixty_seventy.value_counts())
        if True in val1.keys() :
            val1 = val1[True]
        else :
            val1 = 0
        
        btw_sixty_seventy = btw_sixty_seventy & Branch_2
        val2 = dict(btw_sixty_seventy.value_counts())
        if True in val2.keys() :
            val2 = val2[True]
        else :
            val2 = 0
        sheet_structure["Between 60 to 74%"][i] = val1
        sheet_structure["Between 60 to 74%"][i+1] = val2
        sheet_structure["Between 60 to 74%"][i+2] = val - val1 - val2


        # More Than 75
        more_seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.75
        val = dict(more_seventy.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0
        
        more_seventy = more_seventy & Branch_1
        val1 = dict(more_seventy.value_counts())
        if True in val1.keys() :
            val1 = val1[True]
        else :
            val1 = 0
        
        more_seventy = more_seventy & Branch_2
        val2 = dict(more_seventy.value_counts())
        if True in val2.keys() :
            val2 = val2[True]
        else :
            val2 = 0
        sheet_structure["More than 75%"][i] = val1
        sheet_structure["More than 75%"][i+1] = val2
        sheet_structure["More than 75%"][i+2] = val - val1 - val2


        # Maximum Score
        max_score = (df[col_name[i]][3:] + df[unnamed_col[i]][3:])
        p1 = max_score & Branch_1
        p2 = max_score & Branch_2
        p3 = p1 & p2
        sheet_structure["Maximum Score"][i] = max_score[p1].max()
        sheet_structure["Maximum Score"][i+1] = max_score[p2].max()
        sheet_structure["Maximum Score"][i+2] = max_score[~p3].max()


        # Number of Student Passed
        pass_1 = df[col_name[i]][3:] >= df[col_name[i]][2]*0.3
        pass_2 = df[unnamed_col[i]][3:] >= df[unnamed_col[i]][2]*0.3
        pass_3 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.4
        final_pass = pass_1 & pass_2
        final_pass = final_pass & pass_3
        val = dict(final_pass.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0

        final_pass1 = final_pass & Branch_1
        val1 = dict(final_pass1.value_counts())
        if True in val1.keys() :
            val1 = val1[True]
        else :
            val1 = 0
        
        final_pass2 = final_pass & Branch_2
        val2 = dict(final_pass2.value_counts())
        if True in val2.keys() :
            val2 = val2[True]
        else :
            val2 = 0
        sheet_structure["Pass"][i] = val1
        sheet_structure["Pass"][i+1] = val2
        sheet_structure["Pass"][i+2] = val - val1 - val2


        # Total Percentage Student Passed
        s1 = sheet_structure["Number of Students"][i]
        if s1 == 0 :
            s1 = 1
        s2 = sheet_structure["Number of Students"][i+1]
        if s2 == 0 :
            s2 = 1
        s3 = sheet_structure["Number of Students"][i+2]
        if s3 == 0 :
            s3 = 1
        sheet_structure["Pass Percentage"][i] = round(((sheet_structure["Pass"][i])/s1*100), 2)
        sheet_structure["Pass Percentage"][i+1] = round(((sheet_structure["Pass"][i+1])/s2*100), 2)
        sheet_structure["Pass Percentage"][i+2] = round(((sheet_structure["Pass"][i+2])/s3*100), 2)

    analysis = pd.DataFrame( sheet_structure )
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    writer = pd.ExcelWriter( destination )
    analysis.to_excel( writer, "Marks analysis", index = False )
    writer.save()
    # showinfo( title = "Done", message = "Analysis Done" )
    os.startfile( destination )

data = {
    "File" : r"C:\Users\ASUS\OneDrive\Documents\GitHub\Result-Analysis\try_multi.xlsx",
    "Start" : [2100910130001, 2100910130031],
    "End" : [2100910130029, 2100910130039],
    "Spec" : [[],[]],
    "Branch" : ["1","2","3"]
}

analysResult3()