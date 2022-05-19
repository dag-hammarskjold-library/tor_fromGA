import pandas as pd
import argparse
import datetime

parser = argparse.ArgumentParser()

# required, positional arguments for the input data file
parser.add_argument(dest='file_name', type=str, help="File to be processed")
parser.add_argument(dest='sheet_name', type=str, help="Spreadsheet name")
parser.add_argument(dest='ga_session', type=str, help="GA session")
parser.add_argument(dest='pr_year', type=str, help="Press Release year")

#py .\process_excel.py "Resolutions_20211217.xlsx" query 76 2021

args = parser.parse_args()

dt = datetime.datetime.now()
output_file = ".\output\GA_" + str(dt.strftime("%Y%m%d_%H%M")) + ".txt"

#open file, read into pandas table
en_table = pd.read_excel(args.file_name, args.sheet_name)

lines = []
    
for ind in en_table.index:
    res_no = "A/RES/" + en_table['Number'][ind]
    ctte = en_table['Committee'][ind]
    agenda = en_table['Agenda item(s)'][ind]
    meeting = "A/" + args.ga_session + "/PV." + str(en_table['Plenary meeting'][ind])
    date = en_table['Adoption'][ind]
    pr = "XXXX"
    vote = en_table['Vote'][ind]
    draft = en_table['Draft symbol(s)'][ind].replace("and", "&")	
    topic = en_table['Title'][ind]

    if ctte == "First Committee":
        ctte_abbv = "C.1"
    elif ctte == "Second Committee":
        ctte_abbv = "C.2"
    elif ctte == "Third Committee":
        ctte_abbv = "C.3" 
    elif ctte == "Fourth Committee":
        ctte_abbv = "C.4" 
    elif ctte == "Fifth Committee":
        ctte_abbv = "C.5" 
    elif ctte == "Sixth Committee":
        ctte_abbv = "C.6"
    elif ctte == "Plenary":
        ctte_abbv = "Plen."  
    else:
        ctte_abbv = ctte 


    if vote.startswith("N"):
        vote_display = "without a vote"
    else:
        vote_display = vote


    date_display = date.strftime("%d %B %Y")

    lines.append("<tr>")
    lines.append("\n")

    #Resolution number
    lines.append("<td><a href=\"https://undocs.org/en/" + res_no + "\" target=\"_top\">" + res_no + "</a></td>")
    lines.append("\n")

    #Committee
    lines.append("<td>" + ctte_abbv + "</td>")
    lines.append("\n")

    #Agenda
    lines.append("<td>" + agenda + "</td>")
    lines.append("\n")
 
    #Meeting Record
    lines.append("<td><a href=\"https://undocs.org/en/" + meeting + "\" target=\"_top\">" + meeting + "</a><br>")
    lines.append("\n")

    #Date
    lines.append(date_display + "<br>")
    lines.append("\n")

    #Press Release
    lines.append("<a href=\"https://www.un.org/press/en/" + args.pr_year + "/ga" + pr + ".doc.htm\" target=\"_top\">GA/" + pr + "</a><br>")
    lines.append("\n")
   
    #Vote	
    lines.append(vote_display + "</td>")
    lines.append("\n")
    
    #Draft	
    lines.append("<td>" + draft + "</td>")
    lines.append("\n")
    
    #Topic
    lines.append("<td>" + topic + "</td>")
    lines.append("\n")

    lines.append("</tr>")
    lines.append("\n")

    lines.append("\n")

f = open(output_file, "w", encoding="utf-8")
f.writelines(lines)
f.close()