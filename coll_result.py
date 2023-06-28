import os
import pandas as pd
from tkinter import *
import customtkinter as ctk
from PIL import Image ,ImageTk
from tkinter import ttk, filedialog
from tkinter.messagebox import showerror, showinfo

# Defining Main theme of all widgets
ctk.set_appearance_mode( "dark" )
ctk.set_default_color_theme( "dark-blue" )
wid = 1200
hgt = 700

def Imgo(file,w,h) :

    # Image processing
    img=Image.open(file)
    pht=ImageTk.PhotoImage(img.resize((w,h), Image.Resampling.LANCZOS ))
    return pht

def checkFile() :

    if ( file[0] == "" ) :
        showerror( title = "Empty Field", message = "No file found")

    else :
        analysResult()

def analysResult() :

    df = pd.read_excel( file[0] )

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
        "Absent" : [i for i in range(len(col_name))],
        "Pass" : [i for i in range(len(col_name))],
        "Less than 60%" : [i for i in range(len(col_name))],
        "Between 60 to 74%" : [i for i in range(len(col_name))],
        "More than 75%" : [i for i in range(len(col_name))],
        "Maximum Score" : [i for i in range(len(col_name))],
        "Out of Mark" : max_mark,
        "Pass Percentage" : [i for i in range(len(col_name))],
    }

    for i in range( len(col_name) ) :

        student_count = df[col_name[i]][3:] >= 0 
        sheet_structure["Number of Students"][i] = student_count.value_counts()[1]

        abs_count = df[col_name[i]][3:] == 0
        abs_count = dict(abs_count.value_counts())
        abst = 0
        if True in abs_count.keys() :
            abst = abs_count[True]
        sheet_structure["Absent"][i] = abst
        
        less_sixty = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) <= int(max_mark[i]*0.6)
        val = dict(less_sixty.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0
        sheet_structure["Less than 60%"][i] = val

        btw_sixty_seventy_1 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) > int(max_mark[i]*0.6)
        btw_sixty_seventy_2 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= int(max_mark[i]*0.75)
        val_1 = dict(btw_sixty_seventy_1.value_counts())
        val_2 = dict(btw_sixty_seventy_2.value_counts())
        if True in val_1.keys() :
            val = val_1[True]
            if True in val_2.keys() :
                val = val - val_2[True]
        else :
            val = 0
        sheet_structure["Between 60 to 74%"][i] = val

        more_seventy = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.75
        val = dict(more_seventy.value_counts())
        if True in val.keys() :
            val = val[True]
        else :
            val = 0
        sheet_structure["More than 75%"][i] = val

        max_score = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]).max()
        sheet_structure["Maximum Score"][i] = max_score

        pass_1 = df[col_name[i]][3:] >= df[col_name[i]][2]*0.3
        pass_2 = df[unnamed_col[i]][3:] >= df[unnamed_col[i]][2]*0.3
        pass_3 = (df[col_name[i]][3:] + df[unnamed_col[i]][3:]) >= max_mark[i]*0.4
        final_pass = pass_1 & pass_2
        final_pass = final_pass & pass_3
        val1 = 0
        val = dict(final_pass.value_counts())
        if True in val.keys() :
            val1 = val[True]
        sheet_structure["Pass"][i] = val1
        # sheet_structure["Pass"][i] = sheet_structure["Number of Students"][i] - val

        sheet_structure["Pass Percentage"][i] = ((sheet_structure["Pass"][i])/(sheet_structure["Number of Students"][i])*100)

    analysis = pd.DataFrame( sheet_structure )
    destination = file[0].split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    writer = pd.ExcelWriter( destination )
    analysis.to_excel( writer, "Marks ana")
    writer.save()
    showinfo( title = "Done", message = "Analysis Done" )

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
    
        file[0] = open_file

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
    back_image = Imgo( os.path.join( os.getcwd(), "try2a.jpg"), wid+300, hgt+170)
    id_page.create_image(0, 0, image = back_image, anchor = "nw")
    
    jss_image = Imgo( os.path.join( os.getcwd(), "jss.png"), 135, 135)
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

    # # Frame
    # mess = ctk.CTkFrame( master = id_page, 
    #                       width = 780, height = 300, corner_radius = 30,
    #                        bg_color = "#d5eafd", fg_color = "#97e1fe",
    #                         border_color = "#4d89eb", border_width = 6)
    # mess.place_configure( x = 280, y = 480, anchor = "nw")

    # Adding file path
    add_bt = ctk.CTkButton( master = root, 
                             text = "Add", text_font = ( font[1], 20 ), 
                              width = 60, height = 41, corner_radius = 14,
                               bg_color = "#faaa21", fg_color = "#e61800", text_color = "white", 
                                hover_color = "#ff5359", border_width = 0,
                                 command = lambda : openingFile( file_path, file_formate) )
    add_bt_win = id_page.create_window( 800-10, 350-1+30, anchor = "nw", window = add_bt )

    # Adding file path
    anal_bt = ctk.CTkButton( master = root, 
                             text = "Analyse", text_font = ( font[1], 20 ), 
                              width = 60, height = 40, corner_radius = 14,
                               bg_color = "#f78f1c", fg_color = "#e61800", text_color = "white", 
                                hover_color = "#ff5359", border_width = 0,
                                 command = lambda : checkFile() )
    anal_bt_win = id_page.create_window( 400, 520, anchor = "nw", window = anal_bt )

    root.mainloop()

global root

root = ctk.CTk()
root.title( "Result Analysis" )
root.geometry( "1200x700+200+80" )
root.resizable( False, False )
file = [""]
font = [ "Tahoma", "Seoge UI", "Heloia", "Book Antiqua", "Microsoft Sans Serif"]

firstPage()
