import pandas as pd
import numpy as np
import os

# sheet_structure = {
#     "S.No" : [],
#     "Roll No" : [],
#     "Students Name" : [],
# }

# Year_1 = {
#     "S.No" : [],
#     "Roll No" : [],
#     "Students Name" : [],
#     "I Sem Marks out of {sem1}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 1st sem" : [],
#     "II Sem Marks out of {sem2}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 2nd sem" : [],
#     "Marks Obtained in 1 Year" : [],
#     "Percentage" : [],
#     "Over all Status In 1 Year" : []
# }

# Year_2 = {
#     "S.No" : [],
#     "Roll No" : [],
#     "Students Name" : [],
#     "III Sem Marks out of {sem3}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 3rd sem" : [],
#     "IV Sem Marks out of {sem4}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 4th sem" : [],
#     "Marks Obtained in 2 Year" : [],
#     "Percentage" : [],
#     "Over all Status In 2 Year" : []
# }

# Year_3 = {
#     "S.No" : [],
#     "Roll No" : [],
#     "Students Name" : [],
#     "V Sem Marks out of {sem5}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 5th sem" : [],
#     "VI Sem Marks out of {sem6}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 6th sem" : [],
#     "Marks Obtained in 3 Year" : [],
#     "Percentage" : [],
#     "Over all Status In 3 Year" : []
# }

# Year_4 = {
#     "S.No" : [],
#     "Roll No" : [],
#     "Students Name" : [],
#     "VII Sem Marks out of {sem7}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 7th sem" : [],
#     "VIII Sem Marks out of {sem8}" : [],
#     "Status(Pass/Fail/Carry with subject Codes)in 8th sem" : [],
#     "Marks Obtained in 4 Year" : [],
#     "Percentage" : [],
#     "Over all Status In 4 Year" : []
# }

path = ""
xl = pd.ExcelFile( path )
all_sheets = xl.sheet_names
count = 0

if ( len(all_sheets) >=2 ) :
    
    df1 = pd.read_excel( path, all_sheets[0] )
    df2 = pd.read_excel( path, all_sheets[1] )
    col1 = df1.columns
    col2 = df2.columns
    Year_1 = {}
    count +=1

    sno = pd.merge(df1[col1[0]][3:], df2[col2[0]][3:], how='right').astype('int64').to_string(index = False)
    sno = sno.split()
    Year_1["S.No"] = [ int(i) for i in sno[1:]]

    rllno = pd.merge(df1[col1[1]][3:], df2[col2[1]][3:], how='right').astype('int64').to_string(index = False)
    rllno = rllno.split()
    Year_1["Roll No"] = [ int(i) for i in rllno[2:]]

    name = pd.merge(df1[col1[2]][3:], df2[col2[2]][3:], how='right')
    Year_1["Students Name"] = [ j for _,j in list(enumerate(name["Student Name"])) ]

    # odd_sem = int( col1[-2].split()[-1] )
    for i in col1[-2].split() :
        if i.isnumeric() :
            odd_sem = int(i)
    mrk_sem = df1[col1[-2]][3:].astype('int64').to_string(index = False)
    Year_1["I Sem Marks out of " + str(odd_sem)] = [ int(i) for i in mrk_sem.split()]

    status = df1[col1[-4]][3:].to_string(index = False)
    status = [i for i in status.split()]
    temp_status = df1[col1[-3]][3:].to_string(index = False)
    temp_status = [ i for i in temp_status.split()]
    for i in range(len(status)) :
        if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
            status[i] = temp_status[i]
    final_status = status
    Year_1["Status(Pass/Fail/Carry with subject Codes)in 1st sem"] = status

    for i in col2[-2].split() :
        if i.isnumeric() :
            even_sem = int(i)
    # even_sem = int( col2[-2].split()[-1] )
    mrk_sem = df2[col2[-2]][3:].astype('int64').to_string(index = False)
    Year_1["II Sem Marks out of " + str(even_sem)] = [ int(i) for i in mrk_sem.split()]

    status = df2[col2[-4]][3:].to_string(index = False)
    status = [i for i in status.split()]
    temp_status = df2[col2[-3]][3:].to_string(index = False)
    temp_status = [ i for i in temp_status.split()]
    for i in range(len(status)) :
        if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
            status[i] = temp_status[i]
            final_status[i] = temp_status[i]
    Year_1["Status(Pass/Fail/Carry with subject Codes)in 2nd sem"] = status

    mark = (df1[col1[-2]][3:] + df2[col2[-2]][3:])
    mark = mark.astype('int64').to_string(index = False)
    Year_1["Marks Obtained in 1 Year Out of " + str(odd_sem+even_sem)] = [ int(i) for i in mark.split()]

    per = (( df1[col1[-2]][3:] + df2[col2[-2]][3:] )/(odd_sem + even_sem))*100
    per = per.astype('float64').to_string(index = False)
    Year_1["Final Percentage in 1st Year"] = [ float(i) for i in per.split()]
    
    Year_1["Over all Status In 1 Year"] = final_status

    if ( len(all_sheets) >=4 ) :
        
        df1 = pd.read_excel( path, all_sheets[2] )
        df2 = pd.read_excel( path, all_sheets[3] )
        col1 = df1.columns
        col2 = df2.columns
        Year_2 = {}
        count +=1

        sno = pd.merge(df1[col1[0]][3:], df2[col2[0]][3:], how='right').astype('int64').to_string(index = False)
        sno = sno.split()
        Year_2["S.No"] = [ int(i) for i in sno[1:]]

        rllno = pd.merge(df1[col1[1]][3:], df2[col2[1]][3:], how='right').astype('int64').to_string(index = False)
        rllno = rllno.split()
        Year_2["Roll No"] = [ int(i) for i in rllno[2:]]

        name = pd.merge(df1[col1[2]][3:], df2[col2[2]][3:], how='right')
        Year_2["Students Name"] = [ j for _,j in list(enumerate(name["Student Name"])) ]

        # odd_sem = int( col1[-2].split()[-1] )
        for i in col1[-2].split() :
            if i.isnumeric() :
                odd_sem = int(i)
        mrk_sem = df1[col1[-2]][3:].astype('int64').to_string(index = False)
        Year_2["III Sem Marks out of " + str(odd_sem)] = [ int(i) for i in mrk_sem.split()]

        status = df1[col1[-4]][3:].to_string(index = False)
        status = [i for i in status.split()]
        temp_status = df1[col1[-3]][3:].to_string(index = False)
        temp_status = [ i for i in temp_status.split()]
        for i in range(len(status)) :
            if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                status[i] = temp_status[i]
        final_status = status
        Year_2["Status(Pass/Fail/Carry with subject Codes)in 3rd sem"] = status

        # even_sem = int( col2[-2].split()[-1] )
        for i in col2[-2].split() :
            if i.isnumeric() :
                even_sem = int(i)
        mrk_sem = df2[col2[-2]][3:].astype('int64').to_string(index = False)
        Year_2["IV Sem Marks out of " + str(even_sem)] = [ int(i) for i in mrk_sem.split()]

        status = df2[col2[-4]][3:].to_string(index = False)
        status = [i for i in status.split()]
        temp_status = df2[col2[-3]][3:].to_string(index = False)
        temp_status = [ i for i in temp_status.split()]
        for i in range(len(status)) :
            if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                status[i] = temp_status[i]
                final_status[i] = temp_status[i]
        Year_2["Status(Pass/Fail/Carry with subject Codes)in 4th sem"] = status

        mark = (df1[col1[-2]][3:] + df2[col2[-2]][3:])
        mark = mark.astype('int64').to_string(index = False)
        Year_2["Marks Obtained in 2 Year Out of "+str(odd_sem+even_sem)] = [ int(i) for i in mark.split()]

        per = (( df1[col1[-2]][3:] + df2[col2[-2]][3:] )/(odd_sem + even_sem))*100
        per = per.astype('float64').to_string(index = False)
        Year_2["Final Percentage in 2nd Year"] = [ float(i) for i in per.split()]
        
        Year_2["Over all Status In 2 Year"] = final_status

        if ( len(all_sheets) >=6 ) :
            
            df1 = pd.read_excel( path, all_sheets[4] )
            df2 = pd.read_excel( path, all_sheets[5] )
            col1 = df1.columns
            col2 = df2.columns
            Year_3 = {}
            count +=1

            sno = pd.merge(df1[col1[0]][3:], df2[col2[0]][3:], how='right').astype('int64').to_string(index = False)
            sno = sno.split()
            Year_3["S.No"] = [ int(i) for i in sno[1:]]

            rllno = pd.merge(df1[col1[1]][3:], df2[col2[1]][3:], how='right').astype('int64').to_string(index = False)
            rllno = rllno.split()
            Year_3["Roll No"] = [ int(i) for i in rllno[2:]]

            name = pd.merge(df1[col1[2]][3:], df2[col2[2]][3:], how='right')
            Year_3["Students Name"] = [ j for _,j in list(enumerate(name["Student Name"])) ]

            # odd_sem = int( col1[-2].split()[-1] )
            for i in col1[-2].split() :
                if i.isnumeric() :
                    odd_sem = int(i)
            mrk_sem = df1[col1[-2]][3:].astype('int64').to_string(index = False)
            Year_3["V Sem Marks out of " + str(odd_sem)] = [ int(i) for i in mrk_sem.split()]

            status = df1[col1[-4]][3:].to_string(index = False)
            status = [i for i in status.split()]
            temp_status = df1[col1[-3]][3:].to_string(index = False)
            temp_status = [ i for i in temp_status.split()]
            for i in range(len(status)) :
                if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                    status[i] = temp_status[i]
            final_status = status
            Year_3["Status(Pass/Fail/Carry with subject Codes)in 5th sem"] = status

            # even_sem = int( col2[-2].split()[-1] )
            for i in col2[-2].split() :
                if i.isnumeric() :
                    even_sem = int(i)
            mrk_sem = df2[col2[-2]][3:].astype('int64').to_string(index = False)
            Year_3["VI Sem Marks out of " + str(even_sem)] = [ int(i) for i in mrk_sem.split()]

            status = df2[col2[-4]][3:].to_string(index = False)
            status = [i for i in status.split()]
            temp_status = df2[col2[-3]][3:].to_string(index = False)
            temp_status = [ i for i in temp_status.split()]
            for i in range(len(status)) :
                if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                    status[i] = temp_status[i]
                    final_status[i] = temp_status[i]
            Year_3["Status(Pass/Fail/Carry with subject Codes)in 6th sem"] = status

            mark = (df1[col1[-2]][3:] + df2[col2[-2]][3:])
            mark = mark.astype('int64').to_string(index = False)
            Year_3["Marks Obtained in 3 Year Out of "+str(odd_sem+even_sem)] = [ int(i) for i in mark.split()]

            per = (( df1[col1[-2]][3:] + df2[col2[-2]][3:] )/(odd_sem + even_sem))*100
            per = per.astype('float64').to_string(index = False)
            Year_3["Final Percentage in 3rd Year"] = [ float(i) for i in per.split()]
            
            Year_3["Over all Status In 3 Year"] = final_status

            if ( len(all_sheets) >=8 ) :
                
                df1 = pd.read_excel( path, all_sheets[6] )
                df2 = pd.read_excel( path, all_sheets[7] )
                col1 = df1.columns
                col2 = df2.columns
                Year_4 = {}
                count +=1

                sno = pd.merge(df1[col1[0]][3:], df2[col2[0]][3:], how='right').astype('int64').to_string(index = False)
                sno = sno.split()
                Year_4["S.No"] = [ int(i) for i in sno[1:]]

                rllno = pd.merge(df1[col1[1]][3:], df2[col2[1]][3:], how='right').astype('int64').to_string(index = False)
                rllno = rllno.split()
                Year_4["Roll No"] = [ int(i) for i in rllno[2:]]

                name = pd.merge(df1[col1[2]][3:], df2[col2[2]][3:], how='right')
                Year_4["Students Name"] = [ j for _,j in list(enumerate(name["Student Name"])) ]

                # odd_sem = int( col1[-2].split()[-1] )
                for i in col1[-2].split() :
                    if i.isnumeric() :
                        odd_sem = int(i)
                mrk_sem = df1[col1[-2]][3:].astype('int64').to_string(index = False)
                Year_4["VII Sem Marks out of " + str(odd_sem)] = [ int(i) for i in mrk_sem.split()]

                status = df1[col1[-4]][3:].to_string(index = False)
                status = [i for i in status.split()]
                temp_status = df1[col1[-3]][3:].to_string(index = False)
                temp_status = [ i for i in temp_status.split()]
                for i in range(len(status)) :
                    if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                        status[i] = temp_status[i]
                final_status = status
                Year_4["Status(Pass/Fail/Carry with subject Codes)in 7th sem"] = status

                # even_sem = int( col2[-2].split()[-1] )
                for i in col2[-2].split() :
                    if i.isnumeric() :
                        even_sem = int(i)
                mrk_sem = df2[col2[-2]][3:].astype('int64').to_string(index = False)
                Year_4["VIII Sem Marks out of " + str(even_sem)] = [ int(i) for i in mrk_sem.split()]

                status = df2[col2[-4]][3:].to_string(index = False)
                status = [i for i in status.split()]
                temp_status = df2[col2[-3]][3:].to_string(index = False)
                temp_status = [ i for i in temp_status.split()]
                for i in range(len(status)) :
                    if (status[i] != "Pass" ) and (temp_status[i] != 'NaN') :
                        status[i] = temp_status[i]
                        final_status[i] = temp_status[i]
                Year_4["Status(Pass/Fail/Carry with subject Codes)in 8th sem"] = status

                mark = (df1[col1[-2]][3:] + df2[col2[-2]][3:])
                mark = mark.astype('int64').to_string(index = False)
                Year_4["Marks Obtained in 4 Year Out of "+str(odd_sem+even_sem)] = [ int(i) for i in mark.split()]

                per = (( df1[col1[-2]][3:] + df2[col2[-2]][3:] )/(odd_sem + even_sem))*100
                per = per.astype('float64').to_string(index = False)
                Year_4["Final Percentage in 4th Year"] = [ float(i) for i in per.split()]
                
                Year_4["Over all Status In 4 Year"] = final_status

if count>=1 :
    analysis = pd.DataFrame( Year_1 )

    if count>=2 :
        df2 = pd.DataFrame( Year_2 )
        analysis = pd.merge(analysis, df2, on=['S.No', 'Roll No', 'Students Name'], how='right')

        if count>=3 :
            df3 = pd.DataFrame( Year_3 )
            analysis = pd.merge(analysis, df3, on=['S.No', 'Roll No', 'Students Name'], how='right')

            if count==4 :
                df4 = pd.DataFrame( Year_4 )
                analysis = pd.merge(analysis, df4, on=['S.No', 'Roll No', 'Students Name'], how='right')
    
    destination = path.split(".xlsx")[0]
    destination = destination + "_analysis.xlsx"
    writer = pd.ExcelWriter( destination )
    analysis.to_excel( writer , "Marks analysis", index = False )
    writer.save()
    os.startfile( destination )

else :
    print("No data")
