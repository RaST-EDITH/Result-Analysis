from firstdiv import single_branch_analysis
import pandas as pd
import os

def faculty_base_analysis_for_all_sem(data) :

    file_object = pd.ExcelFile(data["File"])

    overall_analysis = {
        "Faculty Name" : [],
        "Semester" : [],
        "Subject" : [],
        "Number of Students" : [],
        "Absent" : [],
        "Pass" : [],
        "Less than 60%" : [],
        "Between 60 to 74%" : [],
        "More than 75%" : [],
        "Maximum Score" : [],
        "Out of Mark" : [],
        "Pass Percentage" : []
    }

    more_than_one_faculty = {
        "Faculty Name" : [],
        "Semester" : [],
        "Subject" : [],
        "Number of Students" : [],
        "Absent" : [],
        "Pass" : [],
        "Less than 60%" : [],
        "Between 60 to 74%" : [],
        "More than 75%" : [],
        "Maximum Score" : [],
        "Out of Mark" : [],
        "Pass Percentage" : []
    }

    for sem in file_object.sheet_names :
        result = single_branch_analysis(data,sem)
        result["Semester"] = [int(sem)]*len(result["Subject"])
        for key, val in result.items() :
            overall_analysis[key].extend(val)

    for i in range(len(overall_analysis["Subject"])-1,-1,-1) :
        multi_faculty = overall_analysis["Faculty Name"][i].split('/')
        if ( len(multi_faculty)>1 ) :
            overall_analysis["Faculty Name"].pop(i)
            more_than_one_faculty["Faculty Name"].extend(multi_faculty)
            for key in more_than_one_faculty.keys() :
                if ( key != "Faculty Name") :
                    cache = overall_analysis[key].pop(i)
                    more_than_one_faculty[key].extend([cache]*len(multi_faculty))
    
    for key, val in more_than_one_faculty.items() :
        overall_analysis[key].extend(val)

    df = pd.DataFrame( overall_analysis )
    sorted_df = df.sort_values( by=["Faculty Name","Semester"])
    even_df = sorted_df[sorted_df["Semester"]%2 == 0]
    odd_df = sorted_df[sorted_df["Semester"]%2 != 0]
    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_faculty_basis_analysis.xlsx"

    with pd.ExcelWriter(destination) as writer:

        sorted_df.to_excel(writer, sheet_name='Full Year', index=False)
        odd_df.to_excel(writer, sheet_name='Odd Semester', index=False)
        even_df.to_excel(writer, sheet_name='Even Semester', index=False)

if ( __name__=="__main__" ) :

    data = {
        "File" : "File path"
    }
    faculty_base_analysis_for_all_sem(data)
