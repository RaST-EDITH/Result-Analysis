from coll_result import analysResult1
import pandas as pd
import os

def faculty_base_analysis_for_all_sem(data) :

    file_object = pd.ExcelFile(data["File"])

    overall_analysis = {
        "Faculty Name" : [],
        "Semester" : [],
        "Semester and Branch" : [],
        "Subject" : [],
        "Number of Students" : [],
        "Pass" : [],
        "PCP" : [],
        "Absent" : [],
        "Students with no result" : [],
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
        "Semester and Branch" : [],
        "Subject" : [],
        "Number of Students" : [],
        "Pass" : [],
        "PCP" : [],
        "Absent" : [],
        "Students with no result" : [],
        "Less than 60%" : [],
        "Between 60 to 74%" : [],
        "More than 75%" : [],
        "Maximum Score" : [],
        "Out of Mark" : [],
        "Pass Percentage" : []
    }

    for sem_branch_name in file_object.sheet_names :

        sem = int(sem_branch_name.split()[0])
        result = analysResult1(data,sem_branch_name)

        result["Semester"] = [int(sem)]*len(result["Subject"])
        result["Semester and Branch"] = [sem_branch_name]*len(result["Subject"])

        # Unequal Length Error handling
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
    sorted_df = df.sort_values( by=["Faculty Name","Semester"], ignore_index = True )
    even_df = sorted_df[sorted_df["Semester"]%2 == 0]
    odd_df = sorted_df[sorted_df["Semester"]%2 != 0]
    del sorted_df["Semester"]
    del even_df["Semester"]
    del odd_df["Semester"]

    mask = even_df["Faculty Name"] == even_df["Faculty Name"].shift()
    even_df.loc[mask, "Faculty Name"] = " "
    mask = odd_df["Faculty Name"] == odd_df["Faculty Name"].shift()
    odd_df.loc[mask, "Faculty Name"] = " "
    mask = sorted_df["Faculty Name"] == sorted_df["Faculty Name"].shift()
    sorted_df.loc[mask, "Faculty Name"] = " "

    destination = data["File"].split(".xlsx")[0]
    destination = destination + "_faculty_basis_analysis.xlsx"

    with pd.ExcelWriter(destination) as writer:

        sorted_df.to_excel(writer, sheet_name='Full Year', index=False)
        odd_df.to_excel(writer, sheet_name='Odd Semester', index=False)
        even_df.to_excel(writer, sheet_name='Even Semester', index=False)

if ( __name__=="__main__" ) :

    # Data must be passed using this
    data = {
        "File" : "File Path"
    }
    faculty_base_analysis_for_all_sem(data)
