import os
import pandas as pd
from tkinter import *
import customtkinter as ctk
from PIL import Image ,ImageTk
from tkinter import ttk, filedialog
from tkinter.messagebox import showerror, showinfo

def Imgo(file,w,h) :

    # Image processing
    img=Image.open(file)
    pht=ImageTk.PhotoImage(img.resize((w,h), Image.Resampling.LANCZOS ))
    return pht

def countRange() :

    df = pd.read_excel( data["File"] )
    row, col = df.shape
    roll = []
    for _,y in enumerate(df[df.columns[1]][3:].astype('int64')) :
        roll.append(y)
    return roll

def checkValue( start_range, end_range, spec_val ) :
    
    try :
        if ( start_range.get() != "" and end_range.get() != "" and data["File"] != "" ) :
            data["Start"] = int(start_range.get())
            data["End"] = int(end_range.get())
            data["Spec"] = spec_val.get().split(",")
            if ( data["Spec"][0] == "" or data["Spec"][0] == "NA" ) :
                data["Spec"] = []
            else :
                try :
                    for i in range( len(data["Spec"])) :
                        data["Spec"][i] = int(data["Spec"][i])
                except :
                    showerror( title = "Invalid Value", message = "Invalid Value Inserted!")
                    return
            
            analysResult2()
        
        elif ( ( start_range.get() == "" or start_range.get() == "NA" ) 
                and ( end_range.get() == "" or end_range.get() == "NA" ) 
                and data["File"] != "" ) :
            data["Start"] = 0
            data["End"] = 0
            data["Spec"] = spec_val.get().split(",")
            if ( data["Spec"][0] == "" or data["Spec"][0] == "NA" ) :
                data["Spec"] = []
                analysResult1()
            else :
                try :
                    for i in range( len(data["Spec"])) :
                        data["Spec"][i] = int(data["Spec"][i])
                    analysResult2()
                except :
                    showerror( title = "Invalid Value", message = "Invalid Value Inserted!")

        else :
            showerror( title = "Empty Field", message = "Incomplete Information!!")
    
    except :
        showerror( title = "Invalid", message = "Invalid Entry!!")

def analysResult1(data, specific_sheet=None ) :

    if ( specific_sheet != None ) :
        df = pd.read_excel( data["File"] , sheet_name=specific_sheet)
    else :
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
    for i in range( len(col_name) ) :
        facl_name.append(df[col_name[i]][0])
    
    max_mark = []
    for i in range( len(col_name) ) :
        max_mark.append(df[col_name[i]][2]+df[unnamed_col[i]][2])

    sheet_structure = {
        "Subject" : col_name,
        "Faculty Name" : facl_name,
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

    for i in range( len(col_name) ) :

        # Absent Students Count
        abs_count = df[col_name[i]].apply(lambda x: isinstance(x, str))
        abst = df[abs_count][col_name[i]].eq('ABS').sum()
        sheet_structure["Absent"][i] = abst

        # Students With no Result Count
        no_result = df[col_name[i]].apply(lambda x: isinstance(x, str))
        no_result_count = df[no_result][col_name[i]].eq("###").sum()
        sheet_structure["Students with no result"][i] = no_result_count
        
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

    for i in range( len(col_name) ) :

        # Students Count
        student_count = df[col_name[i]][3:] >= 0
        store = dict(student_count.value_counts())
        sheet_structure["Number of Students"][i] = store.get(True,0)


        # Less Than 60
        less_than_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) <= int(max_mark[i]*0.6)
        less_than_sixty_val = dict(less_than_sixty.value_counts())
        sheet_structure["Less than 60%"][i] = less_than_sixty_val.get(True,0) - sheet_structure["Students with no result"][i] - sheet_structure["Absent"][i]


        # Between 60 and 75
        sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) > int(max_mark[i]*0.6)
        seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) < int(max_mark[i]*0.75)
        btw_sixty_seventy = sixty & seventy
        btw_sixty_seventy_val = dict(btw_sixty_seventy.value_counts())
        sheet_structure["Between 60 to 74%"][i] = btw_sixty_seventy_val.get(True,0)


        # More Than 75
        more_than_seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.75
        more_than_seventy_val = dict(more_than_seventy.value_counts())
        sheet_structure["More than 75%"][i] = more_than_seventy_val.get(True,0)


        # Maximum Score
        max_score = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]).max()
        sheet_structure["Maximum Score"][i] = max_score


        # Number of Student Passed
        pass_1 = df[col_name[i]][3:] >= df[col_name[i]][2]*0.3
        pass_2 = df[unnamed_col[i]][3:] >= df[unnamed_col[i]][2]*0.3
        pass_3 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.4
        final_pass = pass_1 & pass_2
        final_pass = final_pass & pass_3
        final_pass_val = dict(final_pass.value_counts())
        sheet_structure["Pass"][i] = final_pass_val.get(True,0)

        # PCP -> Not Pass
        sheet_structure["PCP"][i] = final_pass_val.get(False,0) - sheet_structure["Absent"][i] - sheet_structure["Students with no result"][i]

        # Total Percentage Of Student Passed
        sheet_structure["Pass Percentage"][i] = round(((sheet_structure["Pass"][i])/(sheet_structure["Number of Students"][i] - sheet_structure["Students with no result"][i] - sheet_structure["Absent"][i])*100),2)

    if ( specific_sheet != None ) :
        return sheet_structure
    
    analysis = pd.DataFrame( sheet_structure )
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    with pd.ExcelWriter( path=destination ) as writer:
        analysis.to_excel(writer, sheet_name="Marks analysis", index=False)
    # os.startfile( destination )

def analysResult2() :

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
                facl_name.extend([ facl[0], facl[0]])
            else :
                facl_name.extend([ facl[0], facl[1]])
    except :
        for i in range( len(col_name) ) :
            facl_name.extend([df[col_name[i]][0], df[col_name[i]][0]])

    max_mark = []
    for i in range( len(col_name) ) :
        mark = df[col_name[i]][2]+df[unnamed_col[i]][2]
        max_mark.extend([ mark, mark])

    branch = data["Branch"]*len(col_name)

    temp = []
    for i in col_name :
        temp.extend([ i, ""])
    col_name = temp[:]

    temp = []
    for i in unnamed_col :
        temp.extend([ i, ""])
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

    start = data["Start"]
    end = data["End"]
    specific = data["Spec"]

    Branch_1 =  ( df[df.columns[1]][3:] >= start ) & ( df[df.columns[1]][3:] <= end )

    if len(specific) > 0 :
        for roll_no in specific :
            temp = df[df.columns[1]][3:] == roll_no
            Branch_1 = Branch_1 | temp
    

    for i in range( 0, len(col_name), 2 ) :

        # Absent Students Count
        total_absent = df[col_name[i]][3:].eq("ABS").sum()
        total_absent_branch_1 = df[col_name[i]][3:].eq("ABS") & Branch_1
        total_absent_branch_1 = total_absent_branch_1.sum()
        sheet_structure["Absent"][i] = total_absent_branch_1
        sheet_structure["Absent"][i+1] = total_absent - total_absent_branch_1


        # Students With no Result Count
        no_result_count = df[col_name[i]][3:].eq("###").sum()
        no_result_count_1 = df[col_name[i]][3:].eq("###") & Branch_1
        no_result_count_1 = no_result_count_1.sum()
        sheet_structure["Students with no result"][i] = no_result_count_1
        sheet_structure["Students with no result"][i+1] = no_result_count - no_result_count_1


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


    for i in range( 0, len(col_name), 2 ) :

        # Student Count
        student_count = df[col_name[i]][3:] >= 0
        student_total = dict(student_count.value_counts())
        student_total = student_total.get(True,0)

        student_count_1 = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_1
        student_count_1 = dict(student_count_1.value_counts())
        count = student_count_1.get(True,0)

        sheet_structure["Number of Students"][i] = count
        sheet_structure["Number of Students"][i+1] = student_total - count


        # Less than 60
        less_than_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) <= round(max_mark[i]*0.6,0)
        less_than_sixty_val = dict(less_than_sixty.value_counts())
        less_than_sixty_val = less_than_sixty_val.get(True,0)
        
        less_than_sixty_branch_1 = less_than_sixty & Branch_1
        less_than_sixty_val1 = dict(less_than_sixty_branch_1.value_counts())
        less_than_sixty_val1 = less_than_sixty_val1.get(True,0)

        sheet_structure["Less than 60%"][i] = less_than_sixty_val1 - sheet_structure["Students with no result"][i] - sheet_structure["Absent"][i]
        sheet_structure["Less than 60%"][i+1] = less_than_sixty_val - less_than_sixty_val1 - sheet_structure["Students with no result"][i+1] - sheet_structure["Absent"][i+1]


        # Between 60 and 75
        more_than_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) > round(max_mark[i]*0.6,0)
        less_than_seventyfive = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) < round(max_mark[i]*0.75,0)
        
        btw_sixty_seventyfive = more_than_sixty & less_than_seventyfive
        btw_sixty_seventyfive_val = dict(btw_sixty_seventyfive.value_counts())
        btw_sixty_seventyfive_val = btw_sixty_seventyfive_val.get(True,0)

        btw_sixty_seventyfive_branch_1 = btw_sixty_seventyfive & Branch_1
        btw_sixty_seventyfive_val1 = dict(btw_sixty_seventyfive_branch_1.value_counts())
        btw_sixty_seventyfive_val1 = btw_sixty_seventyfive_val1.get(True,0)

        sheet_structure["Between 60 to 74%"][i] = btw_sixty_seventyfive_val1
        sheet_structure["Between 60 to 74%"][i+1] = btw_sixty_seventyfive_val - btw_sixty_seventyfive_val1


        # More Than 75
        more_than_seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= round(max_mark[i]*0.75,0)
        more_than_seventy_val = dict(more_than_seventy.value_counts())
        more_than_seventy_val = more_than_seventy_val.get(True,0)
        
        more_than_seventy_branch_1 = more_than_seventy & Branch_1
        more_than_seventy_val1 = dict(more_than_seventy_branch_1.value_counts())
        more_than_seventy_val1 = more_than_seventy_val1.get(True,0)

        sheet_structure["More than 75%"][i] = more_than_seventy_val1
        sheet_structure["More than 75%"][i+1] = more_than_seventy_val - more_than_seventy_val1


        # Maximum Score
        max_score = (df[col_name[i]][3:] + df[unnamed_col[i]][3:])
        p = max_score & Branch_1
        sheet_structure["Maximum Score"][i] = max_score[p].max()
        sheet_structure["Maximum Score"][i+1] = max_score[~p].max()


        # Number of Student Passed
        pass_1 = df[col_name[i]][3:] >= df[col_name[i]][2]*0.3
        pass_2 = df[unnamed_col[i]][3:] >= df[unnamed_col[i]][2]*0.3
        pass_3 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.4

        final_pass = pass_1 & pass_2
        final_pass = final_pass & pass_3
        final_pass_val = dict(final_pass.value_counts())
        final_pass_val = final_pass_val.get(True,0)

        final_pass_branch_1 = final_pass & Branch_1
        final_pass_val1 = dict(final_pass_branch_1.value_counts())
        final_pass_val1 = final_pass_val1.get(True,0)

        sheet_structure["Pass"][i] = final_pass_val1
        sheet_structure["Pass"][i+1] = final_pass_val - final_pass_val1

        
        # PCP -> Not Pass
        sheet_structure["PCP"][i] = sheet_structure["Number of Students"][i] - sheet_structure["Pass"][i] - sheet_structure["Absent"][i] - sheet_structure["Students with no result"][i]
        sheet_structure["PCP"][i+1] = sheet_structure["Number of Students"][i+1] - sheet_structure["Pass"][i+1] - sheet_structure["Absent"][i+1] - sheet_structure["Students with no result"][i+1]


        # Total Percentage Student Passed
        s1 = max(1,sheet_structure["Number of Students"][i]-sheet_structure["Students with no result"][i]-sheet_structure["Absent"][i])
        s2 = max(1,sheet_structure["Number of Students"][i+1]-sheet_structure["Students with no result"][i+1]-sheet_structure["Absent"][i+1])

        sheet_structure["Pass Percentage"][i] = round(((sheet_structure["Pass"][i])/s1*100), 2)
        sheet_structure["Pass Percentage"][i+1] = round(((sheet_structure["Pass"][i+1])/s2*100), 2)


    analysis = pd.DataFrame( sheet_structure )
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    with pd.ExcelWriter( path=destination ) as writer:
        analysis.to_excel(writer, sheet_name="Marks analysis", index=False)
    os.startfile( destination )

def openingFile( file_path, file_formate ) :

    # Opening File
    if ( file_path.get() != "" ) :

        # Getting path of file from entry box
        open_file = file_path.get()

    else :

        # Getting path of file from filedialog
        open_file = filedialog.askopenfilename( initialdir = os.getcwd(), title = "Open file", filetypes = file_formate )

    # Checking for empty address
    if ( open_file != "" ) :
    
        # file[0] = open_file
        data["File"] = open_file

        if ( file_path.get() != "" ) :
            file_path.delete( 0, END)
        
        file_path.insert( 0, open_file )
       
    else :

        # Showing error due to empty credientials
        showerror( title = "Empty Field", message = "No file found")

def firstPage() :

    # Defining Structure
    id_page = Canvas( root, 
                        width = wid, height = hgt, 
                         bg = "black", highlightcolor = "#3c5390", 
                          borderwidth = 0 )
    id_page.pack( fill = "both", expand = True )

    # Image on top
    back_image = Imgo( os.path.join( os.getcwd(), "images\\back.jpg"), wid+300, hgt+170)
    id_page.create_image(0, 0, image = back_image, anchor = "nw")
    
    jss_image = Imgo( os.path.join( os.getcwd(), "images\jss.png"), 135, 135)
    id_page.create_image(40, 20+30, image = jss_image, anchor = "nw")
    
    # Heading
    id_page.create_text(840,80+30,text="JSS ACADEMY OF TECHNICAL EDUCATION", 
                            font = ( font[0], 35, "bold" ), fill = "#00234f" )
    
    # Accessing the file
    file_path = ctk.CTkEntry( master = root, 
                                placeholder_text = "Enter Path", text_font = ( font[2], 20 ), 
                                 width = 550, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#494949", text_color = "#242424", 
                                   fg_color = "#ffd371", bg_color = "#faaa21", 
                                    border_color = "#00234f", border_width = 4)
    file_path_win = id_page.create_window( 80, 350+30, anchor = "nw", window = file_path )

    file_formate = [( "Excel file", "*.xlsx")]

    # Adding file path
    add_bt = ctk.CTkButton( master = root, 
                             text = "Add", text_font = ( font[1], 20 ), 
                              width = 60, height = 41, corner_radius = 14,
                               bg_color = "#faaa21", fg_color = "#e61800", text_color = "white", 
                                hover_color = "#ff5359", border_width = 0,
                                 command = lambda : openingFile( file_path, file_formate) )
    add_bt_win = id_page.create_window( 800-10, 350-1+30, anchor = "nw", window = add_bt )

    # Starting Range
    start_range = ctk.CTkEntry( master = root, 
                                placeholder_text = "Range From", text_font = ( font[2], 20 ), 
                                 width = 260, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#494949", text_color = "#242424", 
                                   fg_color = "#ffd371", bg_color = "#f7991d", 
                                    border_color = "#00234f", border_width = 4)
    start_range_win = id_page.create_window( 80, 350+30+70, anchor = "nw", window = start_range )
    
    # Ending Range
    end_range = ctk.CTkEntry( master = root, 
                                placeholder_text = "To", text_font = ( font[2], 20 ), 
                                 width = 260, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#494949", text_color = "#242424", 
                                   fg_color = "#ffd371", bg_color = "#f7991d", 
                                    border_color = "#00234f", border_width = 4)
    end_range_win = id_page.create_window( 80+363, 350+30+70, anchor = "nw", window = end_range )

    # Specific Values
    spec_val = ctk.CTkEntry( master = root, 
                                placeholder_text = "Specific Value", text_font = ( font[2], 20 ), 
                                 width = 550, height = 30, corner_radius = 14,
                                  placeholder_text_color = "#494949", text_color = "#242424", 
                                   fg_color = "#ffd371", bg_color = "#f6911d", 
                                    border_color = "#00234f", border_width = 4)
    spec_val_win = id_page.create_window( 80, 350+30+140, anchor = "nw", window = spec_val )

    # Analysing the file
    analy_bt = ctk.CTkButton( master = root, 
                             text = "Analyse", text_font = ( font[1], 20 ), 
                              width = 60, height = 40, corner_radius = 14,
                               bg_color = "#f78f1c", fg_color = "#e61800", text_color = "white", 
                                hover_color = "#ff5359", border_width = 0,
                                 command = lambda : checkValue( start_range, end_range, spec_val ) )
    analy_bt_win = id_page.create_window( 400, 520+100, anchor = "nw", window = analy_bt )

    root.mainloop()

if __name__ == "__main__" :

    global root

    # Defining Main theme of all widgets
    ctk.set_appearance_mode( "dark" )
    ctk.set_default_color_theme( "dark-blue" )
    wid = 1200
    hgt = 700
    root = ctk.CTk()
    root.title( "Result Analysis" )
    root.geometry( "1200x700+200+80" )
    root.resizable( False, False )
    data = {
        "File" : "",
        "Start" : 0,
        "End" : 0,
        "Spec" : [],
        "Branch" : [ "IT1", "IT2"]
    }
    font = [ "Tahoma", "Seoge UI", "Heloia", "Book Antiqua", "Microsoft Sans Serif"]

    firstPage()
