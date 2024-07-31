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
        "Pass" : [i for i in range(len(col_name))],
        "PCP" : [i for i in range(len(col_name))],
        "Absent" : [i for i in range(len(col_name))],
        "Students with no result" : [i for i in range(len(col_name))],
        "Less than 60%" : [i for i in range(len(col_name))],
        "Between 60 to 74%" : [i for i in range(len(col_name))],
        "More than 75%" : [i for i in range(len(col_name))],
        "Maximum Score" : [i for i in range(len(col_name))],
        "Out of Mark" : max_mark,
        "Pass Percentage" : [i for i in range(len(col_name))],
    }

    start_1 = data["Start"][0]
    start_2 = data["Start"][1]
    end_1 = data["End"][0]
    end_2 = data["End"][1]
    specific_1 = data["Spec"][0]
    specific_2 = data["Spec"][1]

    Branch_1 =  ( df[df.columns[1]][3:] >= start_1 ) & ( df[df.columns[1]][3:] <= end_1 )
    Branch_2 =  ( df[df.columns[1]][3:] >= start_2 ) & ( df[df.columns[1]][3:] <= end_2 )

    if len(specific_1) > 0 :
        for roll_no in specific_1 :
            sp = df[df.columns[1]][3:] == roll_no
            Branch_1 = Branch_1 | sp

    if len(specific_2) > 0 :
        for roll_no in specific_2 :
            sp = df[df.columns[1]][3:] == roll_no
            Branch_2 = Branch_2 | sp


    for i in range( 0, len(col_name), 3 ) :

        # Absent Students Count
        total_absent = df[col_name[i]][3:].eq("ABS").sum()
        total_absent_branch_1 = df[col_name[i]][3:].eq("ABS") & Branch_1
        total_absent_branch_1 = total_absent_branch_1.sum()
        total_absent_branch_2 = df[col_name[i]][3:].eq("ABS") & Branch_2
        total_absent_branch_2 = total_absent_branch_2.sum()
        sheet_structure["Absent"][i] = total_absent_branch_1
        sheet_structure["Absent"][i+1] = total_absent_branch_2
        sheet_structure["Absent"][i+2] = total_absent - total_absent_branch_1 - total_absent_branch_2


        # Students With no Result Count
        no_result_count = df[col_name[i]][3:].eq("###").sum()
        no_result_count_1 = df[col_name[i]][3:].eq("###") & Branch_1
        no_result_count_1 = no_result_count_1.sum()
        no_result_count_2 = df[col_name[i]][3:].eq("###") & Branch_2
        no_result_count_2 = no_result_count_2.sum()
        sheet_structure["Students with no result"][i] = no_result_count_1
        sheet_structure["Students with no result"][i+1] = no_result_count_2
        sheet_structure["Students with no result"][i+2] = no_result_count - no_result_count_1 - no_result_count_2


        # Replacing ABS
        if ( abst>0 ) :
            df[col_name[i]] = df[col_name[i]].replace("ABS", 0)

        # Replacing "###"
        if ( no_result_count>0 ) :
            df[col_name[i]] = df[col_name[i]].replace("###", 0)
    
    # Checking "ABS" for overall sheet
    abst = (df == 'ABS').sum().sum()
    if ( abst>0 ) :
        df = df.replace("ABS", 0)
    
    # Checking "###" for overall sheet
    un_diclare = (df == '###').sum().sum()
    if ( un_diclare>0 ) :
        df = df.replace("###", 0)


    for i in range( 0, len(col_name), 3 ) :

        # Student Count
        student_count = df[col_name[i]][3:] >= 0
        student_total = dict(student_count.value_counts())
        student_total = student_total.get(True,0)

        student_count_1 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_1
        student_count_1 = dict(student_count_1.value_counts())
        count1 = student_count_1.get(True,0)
        
        student_count_2 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_2
        student_count_2 = dict(student_count_2.value_counts())
        count2 = student_count_2.get(True,0)
        
        sheet_structure["Number of Students"][i] = count1
        sheet_structure["Number of Students"][i+1] = count2
        sheet_structure["Number of Students"][i+2] = student_total - count1 - count2


        # Less than 60
        less_than_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) <= round(max_mark[i]*0.6,0)
        less_than_sixty_val = dict(less_than_sixty.value_counts())
        less_than_sixty_val = less_than_sixty_val.get(True,0)
        
        less_than_sixty_1 = less_than_sixty & Branch_1
        less_than_sixty_val1 = dict(less_than_sixty_1.value_counts())
        less_than_sixty_val1 = less_than_sixty_val1.get(True,0)
        
        less_than_sixty_2 = less_than_sixty & Branch_2
        less_than_sixty_val2 = dict(less_than_sixty_2.value_counts())
        less_than_sixty_val2 = less_than_sixty_val2.get(True,0)

        sheet_structure["Less than 60%"][i] = less_than_sixty_val1 - sheet_structure["Students with no result"][i] - sheet_structure["Absent"][i]
        sheet_structure["Less than 60%"][i+1] = less_than_sixty_val2 - sheet_structure["Students with no result"][i+1] - sheet_structure["Absent"][i+1]
        sheet_structure["Less than 60%"][i+2] = less_than_sixty_val - less_than_sixty_val1 - less_than_sixty_val2 - sheet_structure["Students with no result"][i+2] - sheet_structure["Absent"][i+2]


        # Between 60 and 75
        more_than_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) > round(max_mark[i]*0.6,0)
        less_than_seventyfive = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) < round(max_mark[i]*0.75,0)

        btw_sixty_seventyfive = more_than_sixty & less_than_seventyfive
        btw_sixty_seventyfive_val = dict(btw_sixty_seventyfive.value_counts())
        btw_sixty_seventyfive_val = btw_sixty_seventyfive_val.get(True,0)

        btw_sixty_seventyfive_branch_1 = btw_sixty_seventyfive & Branch_1
        btw_sixty_seventyfive_val1 = dict(btw_sixty_seventyfive_branch_1.value_counts())
        btw_sixty_seventyfive_val1 = btw_sixty_seventyfive_val1.get(True,0)
        
        btw_sixty_seventyfive_branch_2 = btw_sixty_seventyfive & Branch_2
        btw_sixty_seventyfive_val2 = dict(btw_sixty_seventyfive_branch_2.value_counts())
        btw_sixty_seventyfive_val2 = btw_sixty_seventyfive_val2.get(True,0)

        sheet_structure["Between 60 to 74%"][i] = btw_sixty_seventyfive_val1
        sheet_structure["Between 60 to 74%"][i+1] = btw_sixty_seventyfive_val2
        sheet_structure["Between 60 to 74%"][i+2] = btw_sixty_seventyfive_val - btw_sixty_seventyfive_val1 - btw_sixty_seventyfive_val2


        # More Than 75
        more_than_seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= round(max_mark[i]*0.75,0)
        more_than_seventy_val = dict(more_than_seventy.value_counts())
        more_than_seventy_val = more_than_seventy_val.get(True,0)
        
        more_than_seventy_branch_1 = more_than_seventy & Branch_1
        more_than_seventy_val1 = dict(more_than_seventy_branch_1.value_counts())
        more_than_seventy_val1 = more_than_seventy_val1.get(True,0)
        
        more_than_seventy_branch_2 = more_than_seventy & Branch_2
        more_than_seventy_val2 = dict(more_than_seventy_branch_2.value_counts())
        more_than_seventy_val2 = more_than_seventy_val2.get(True,0)

        sheet_structure["More than 75%"][i] = more_than_seventy_val1
        sheet_structure["More than 75%"][i+1] = more_than_seventy_val2
        sheet_structure["More than 75%"][i+2] = more_than_seventy_val - more_than_seventy_val1 - more_than_seventy_val2


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
        val = val.get(True,0)

        final_pass1 = final_pass & Branch_1
        val1 = dict(final_pass1.value_counts())
        val1 = val1.get(True,0)
        
        final_pass2 = final_pass & Branch_2
        val2 = dict(final_pass2.value_counts())
        val2 = val2.get(True,0)

        sheet_structure["Pass"][i] = val1
        sheet_structure["Pass"][i+1] = val2
        sheet_structure["Pass"][i+2] = val - val1 - val2

        
        # PCP -> Not Pass
        sheet_structure["PCP"][i] = sheet_structure["Number of Students"][i] - sheet_structure["Pass"][i] - sheet_structure["Absent"][i] - sheet_structure["Students with no result"][i]
        sheet_structure["PCP"][i+1] = sheet_structure["Number of Students"][i+1] - sheet_structure["Pass"][i+1] - sheet_structure["Absent"][i+1] - sheet_structure["Students with no result"][i+1]
        sheet_structure["PCP"][i+2] = sheet_structure["Number of Students"][i+2] - sheet_structure["Pass"][i+2] - sheet_structure["Absent"][i+2] - sheet_structure["Students with no result"][i+2]


        # Total Percentage Student Passed
        s1 = max(1,sheet_structure["Number of Students"][i]-sheet_structure["Students with no result"][i]-sheet_structure["Absent"][i])
        s2 = max(1,sheet_structure["Number of Students"][i+1]-sheet_structure["Students with no result"][i+1]-sheet_structure["Absent"][i+1])
        s3 = max(1,sheet_structure["Number of Students"][i+2]-sheet_structure["Students with no result"][i+2]-sheet_structure["Absent"][i+2])
        
        sheet_structure["Pass Percentage"][i] = round(((sheet_structure["Pass"][i])/s1*100), 2)
        sheet_structure["Pass Percentage"][i+1] = round(((sheet_structure["Pass"][i+1])/s2*100), 2)
        sheet_structure["Pass Percentage"][i+2] = round(((sheet_structure["Pass"][i+2])/s3*100), 2)


    analysis = pd.DataFrame( sheet_structure )
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    with pd.ExcelWriter( path=destination ) as writer:
        analysis.to_excel(writer, sheet_name="Marks analysis", index=False)
    os.startfile( destination )

if ( __name__ == "__main__" ) :

    data = {
        "File" : 'File Path' ,
        "Start" : [ 'Start Roll No 1', 'Start Roll No 2' ],
        "End" : [ 'End Roll No 1', 'End Roll No 2' ],
        "Spec" : [['Specific Roll Number'],['Specific Roll Number']],
        "Branch" : [ 'Branch 1 Name', 'Branch 2 Name', 'Branch 3 Name']
    }

    analysResult3()
