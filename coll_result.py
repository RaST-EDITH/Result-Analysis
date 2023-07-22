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
    for i in range( len(col_name) ) :
        facl = df[col_name[i]][0].split('/')
        if ( len(facl) == 1) :
            facl_name.extend([ facl[0], facl[0]])
        else :
            facl_name.extend([ facl[0], facl[1]])

    max_mark = []
    for i in range( len(col_name) ) :
        mark = df[col_name[i]][2]+df[unnamed_col[i]][2]
        max_mark.extend([ mark, mark])

    branch = [ "IT1", "IT2"]*len(col_name)

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
        "Absent" : [i for i in range(len(col_name))],
        "Pass" : [i for i in range(len(col_name))],
        "Less than 60%" : [i for i in range(len(col_name))],
        "Between 60 to 74%" : [i for i in range(len(col_name))],
        "More than 75%" : [i for i in range(len(col_name))],
        "Maximum Score" : [i for i in range(len(col_name))],
        "Out of Mark" : max_mark,
        "Pass Percentage" : [i for i in range(len(col_name))],
    }

    a = data["Start"]
    b = data["End"]
    spec = data["Spec"]

    Branch_1 =  ( df[df.columns[1]][3:] >= a ) & ( df[df.columns[1]][3:] <= b )

    if len(spec) > 0 :
        for x in spec :
            it_1 = df[df.columns[1]][3:] == x
            Branch_1 = Branch_1 | it_1

    for i in range( 0, len(col_name), 2 ) :

        # Student Count
        student_count = df[col_name[i]][3:] >= 0
        student_val = student_count.value_counts()[1]
        student_count = ( ~df[col_name[i]][3:].isnull() ) & student_count & Branch_1
        student_count = dict(student_count.value_counts())
        count = 0
        if True in student_count.keys() :
            count = student_count[True]
        sheet_structure["Number of Students"][i] = count
        sheet_structure["Number of Students"][i+1] = student_val - count


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
        sheet_structure["Absent"][i] = abst1
        sheet_structure["Absent"][i+1] = abst - abst1
        

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
        sheet_structure["Less than 60%"][i] = val1
        sheet_structure["Less than 60%"][i+1] = val - val1


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
        sheet_structure["Between 60 to 74%"][i] = val1
        sheet_structure["Between 60 to 74%"][i+1] = val - val1


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
        sheet_structure["More than 75%"][i] = val1
        sheet_structure["More than 75%"][i+1] = val - val1


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
        val = dict(final_pass.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0

        final_pass = final_pass & Branch_1
        val1 = dict(final_pass.value_counts())
        if True in val1.keys() :
            val1 = val1[True]
        else :
            val1 = 0
        sheet_structure["Pass"][i] = val1
        sheet_structure["Pass"][i+1] = val - val1


        # Total Percentage Student Passed
        s1 = sheet_structure["Number of Students"][i]
        if s1 == 0 :
            s1 = 1
        s2 = sheet_structure["Number of Students"][i+1]
        if s2 == 0 :
            s2 = 1
        sheet_structure["Pass Percentage"][i] = ((sheet_structure["Pass"][i])/s1*100)
        sheet_structure["Pass Percentage"][i+1] = ((sheet_structure["Pass"][i+1])/s2*100)

    analysis = pd.DataFrame( sheet_structure )
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    writer = pd.ExcelWriter( destination )
    analysis.to_excel( writer, "Marks ana", index = False )
    writer.save()
    showinfo( title = "Done", message = "Analysis Done" )
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

    # Adding file path
    analy_bt = ctk.CTkButton( master = root, 
                             text = "Analyse", text_font = ( font[1], 20 ), 
                              width = 60, height = 40, corner_radius = 14,
                               bg_color = "#f78f1c", fg_color = "#e61800", text_color = "white", 
                                hover_color = "#ff5359", border_width = 0,
                                 command = lambda : checkValue( start_range, end_range, spec_val ) )
    analy_bt_win = id_page.create_window( 400, 520+100, anchor = "nw", window = analy_bt )

    root.mainloop()

if __name__ == "__main__" :

    # Defining Main theme of all widgets
    ctk.set_appearance_mode( "dark" )
    ctk.set_default_color_theme( "dark-blue" )
    wid = 1200
    hgt = 700

    global root

    root = ctk.CTk()
    root.title( "Result Analysis" )
    root.geometry( "1200x700+200+80" )
    root.resizable( False, False )
    data = {
        "File" : "",
        "Start" : 0,
        "End" : 0,
        "Spec" : []
    }
    font = [ "Tahoma", "Seoge UI", "Heloia", "Book Antiqua", "Microsoft Sans Serif"]

    firstPage()
