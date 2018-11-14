# Program    : PDFReader
# Author     : David Velez
# Date       : 07/24/18
# Description: Read PDF Files and Export Relevant Data to a CSV File


# Imports
import csv
import io
import os
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

# Global Directories
base_folder = os.path.dirname(__file__)
national_folder = os.path.join(base_folder, "data", "national")
north_regional_folder = os.path.join(base_folder, "data", "regional", "north")
south_regional_folder = os.path.join(base_folder, "data", "regional", "south")
midwest_regional_folder = os.path.join(base_folder, "data", "regional", "midwest")
west_regional_folder = os.path.join(base_folder, "data", "regional", "west")
export_folder = os.path.join(base_folder, "export")
log_folder = os.path.join(base_folder, "log")

pdf_folders = [national_folder, north_regional_folder, south_regional_folder, midwest_regional_folder,
               west_regional_folder]


# Main
def main():
    print_header()

    # Run Functions
    for index in range(len(pdf_folders)):
        column_headers = get_column_headers(pdf_folders[index])
        try:
            export_to_csv(export_folder, column_headers, "w", pdf_folders[index])
            search_through_files(pdf_folders[index])
        except Exception as e:
            print("Issue with reading file: {}".format(e))

    # End of Program
    print("\nComplete!")


# Print the Header
def print_header():
    print("-----------------------------")
    print("       PDF Reader App")
    print("-----------------------------")
    print()
    print("Starting Program...")


# Column Header for CSV Export
def get_column_headers(folder):
    if os.path.basename(folder) == "national":
        return [["University", "Rank", "USN Score", "Grad Rate Percentage",
                 "Predicted Grad Rate", "OverUnder Performance",
                 "Alumni Giving Percentage", "Alumni Giving Rank",
                 "Alumni Giving Rate", "Grad Retention Rate",
                 "Grad Retention Rank", "6 Year Grad Rate",
                 "Avg Freshman Retention Rate", "Undergrad Academic Rep",
                 "Peer Assessment Score", "HS Counselor Score",
                 "Faculty Resources Percentage", "Faculty Resources Rank",
                 "Full-Time Faculty Percentage", "Full-Time PhD Terminal Percentage",
                 "Classes Fewer than 20", "Classes More than 50",
                 "Student Faculty Ratio", "Student Selectivity Percentage",
                 "Student Selectivity Rank", "SAT/ACT Percentage",
                 "Freshmen in Top 10 of HS", "Freshmen in Top 25 of HS",
                 "Fall 2016 Acceptance Rate", "Financial Resources Percentage",
                 "Financial Resources Rank"]]

    elif os.path.basename(folder) == "north" or "south" or "midwest" or "west":
        return [["University", "Rank", "USN Score", "Grad Rate %",
                 "Grad Rate Rank", "Avg 6 Year Grade Percentage",
                 "6 Year Grad with Pell Grant",
                 "6 Year Grad No Pell Grant Percentage",
                 "Diff between Pell and No Pell Percentage",
                 "Avg First-Year Stud Retention Percentage",
                 "Grad Rate Performance Percentage", "Predicted Grad Rate Percentage",
                 "Over/Under Performance", "Expert Opinion Percentage",
                 "Peer Assessment Score", "HS Counselor Score",
                 "Faculty Resources Percentage", "Faculty Resources Rank",
                 "Full-Time Faculty Percentage", "Full-Time PhD Terminal Percentage",
                 "Classes Fewer than 20", "Class More than 50",
                 "Student Faculty Ratio", "Student Excellence Percentage",
                 "Student Excellence Rank", "SAT/ACT Percentage",
                 "Freshmen in Top 10% of HS", "Freshmen in Top 25% of HS",
                 "Financial Resources Percentage", "Financial Resources Rank",
                 "Alumni Giving Percentage", "Alumni Giving Rank",
                 "Avg Alumni Giving Percentage"]]


# Search Through PDF Files
def search_through_files(folder):
    # Get PDF Items
    items = os.listdir(folder)

    # Log Locker
    log_lock = True

    for item in items:

        if item.__contains__("Rankings Indicators") and item.__contains__(".pdf"):
            full_item = os.path.join(folder, item)
        else:
            print("Passed {}".format(item))
            print("FOLDER!!!" + folder)
            log_lock = export_to_log(log_folder, item, log_lock, folder)
            continue

        if os.path.isdir(full_item):
            continue
        else:
            print("Opening File {}.pdf...".format(item[:item.find(" _")]))
            pdf_parser(full_item)
        print("Closing File {}.pdf...".format(item[:item.find(" _")]))


# Parse the PDFs to be Human Readable
# NOTE: Credit to https://stackoverflow.com/questions/25665/python-module-for-converting-pdf-to-text
def pdf_parser(pdf_file):
    fp = open(pdf_file, 'rb')
    rsrcmgr = PDFResourceManager()
    retstr = io.StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)

    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Process each page contained in the document.
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
        data = retstr.getvalue()

    # Retrieve Values from PDF
    values = retrieve_values(data, os.path.basename(os.path.dirname(pdf_file)))

    print("Exporting to CSV...")
    export_to_csv(export_folder, values, "a", os.path.dirname(pdf_file))

    # close file
    fp.close()


# Retrieve all values from the PDF
def retrieve_values(data, dir_name):
    # Variables
    global school_name, school_rank, usn_score, grad_rate_perf, \
        predict_grad_rate, over_under_perf, alum_giving, alum_give_rank, \
        avg_alum_give_rate, grad_and_ret_rates, grad_and_ret_rank, \
        six_year_grade_rate, avg_fresh_ret_rate, undergrad_acad_rep, \
        peer_assess_score, hs_coun_score, fac_res, fac_res_rank, \
        perc_fac_fulltime, fulltime_fac_phd_term_deg, classes_fewer_twenty, \
        classes_fifty_more, stud_fac_ratio, stud_sel, stud_sel_rank, \
        sat_act_perc, fresh_top_ten_hs, fresh_top_25_hs, fall_2016_acc_rate, \
        fin_res, fin_res_rank, grad_ret_rate_perc, grad_ret_rank, \
        avg_six_grad_rate, six_year_received_pell, six_year_no_pell, \
        diff_pell_nonpell, avg_firstyear_ret_rate, pred_grad_rate, \
        expert_opinion, stud_excel, stud_excel_rank

    # Print
    print("Gathering values...")

    if dir_name == "national":

        # For every line get values
        for line in data.splitlines():

            if line.find("is ranked") >= 0:
                school_name = line[:line.find("is ranked")]
                school_rank = line[line.find("#") + 1:line.find("in National")].rstrip()
            elif line.find("Score:") >= 0:
                usn_score = extract_score(line)
            elif line.find("Graduation Rate Performance Rank") >= 0:
                grad_rate_perf = line[line.find("(") + 1:line.find("%")]
            elif line.find("Predicted graduation rate:") >= 0:
                predict_grad_rate = line[line.find(":") + 2:line.find("%")]
            elif line.find("Overperformance(+)/Underperformance(-)") >= 0:
                if line.find("-") >= 0:
                    over_under_perf = line[line.find(":") + 2:]
                else:
                    over_under_perf = extract_score(line)
            elif line.find("Alumni Giving") >= 0:
                alum_giving = line[line.find("(") + 1:line.find("%")]
            elif line.find("Alumni giving rank:") >= 0:
                alum_give_rank = extract_score(line)
            elif line.find("Average alumni giving rate:") >= 0:
                if line.find("]") >= 0:
                    avg_alum_give_rate = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    avg_alum_give_rate = line[line.find(":") + 2:line.find("%")]
                else:
                    avg_alum_give_rate = line[line.find(":") + 2:]
            elif line.find("Graduation and Retention Rates") >= 0:
                grad_and_ret_rates = line[line.find("(") + 1:line.find("%")]
            elif line.find("Graduation and retention rank") >= 0:
                grad_and_ret_rank = extract_score(line)
            elif line.find("6-year graduation rate:") >= 0:
                if line.find("]") >= 0:
                    six_year_grade_rate = line[line.find("]") + 2:line.find("%")]
                else:
                    six_year_grade_rate = line[line.find(":") + 2:line.find("%")]
            elif line.find("Average freshman retention rate:") >= 0:
                if line.find("]") >= 0:
                    avg_fresh_ret_rate = line[line.find("]") + 2:line.find("%")]
                else:
                    avg_fresh_ret_rate = line[line.find(":") + 2:line.find("%")]
            elif line.find("Undergraduate Academic Reputation") >= 0:
                undergrad_acad_rep = line[line.find("(") + 1:line.find("%")]
            elif line.find("Peer assessment score (out of 5):") >= 0:
                peer_assess_score = line[line.find(":") + 2:]
            elif line.find("High school counselor score (out of 5):") >= 0:
                hs_coun_score = line[line.find(":") + 2:]
            elif line.find("Faculty Resources (") >= 0:
                fac_res = line[line.find("(") + 1:line.find("%")]
            elif line.find("Faculty Resources Rank:") >= 0:
                fac_res_rank = extract_score(line)
            elif line.find("Percent of faculty who are full-time:") >= 0:
                if line.find("]") >= 0:
                    perc_fac_fulltime = line[line.find("]") + 2:line.find("%")]
                else:
                    perc_fac_fulltime = line[line.find(":") + 2:line.find("%")]
            elif line.find("Full-time faculty with Ph.D or terminal degree:") >= 0:
                if line.find("%") >= 0:
                    fulltime_fac_phd_term_deg = line[line.find(":") + 2:line.find("%")]
                else:
                    fulltime_fac_phd_term_deg = line[line.find(":") + 2:]
            elif line.find("Classes with fewer than 20 students: ") >= 0:
                if line.find("%") >= 0:
                    classes_fewer_twenty = line[line.find(":") + 2:line.find("%")]
                else:
                    classes_fewer_twenty = line[line.find(":") + 2:line.find(":1")]
            elif line.find("Classes with 50 or more students:") >= 0:
                if line.find("]") >= 0:
                    classes_fifty_more = line[line.find("]") + 2:line.find("%")]
                else:
                    classes_fifty_more = line[line.find(":") + 2:]
            elif line.find("Student-faculty ratio:") >= 0:
                if line.find("]") >= 0:
                    stud_fac_ratio = line[line.find("]") + 2:]
                elif line.find("/") >= 0:
                    stud_fac_ratio = line[line.find(":") + 2:]
                else:
                    stud_fac_ratio = line[line.find(":") + 2:line.find(":1")]
            elif line.find("Student Selectivity (") >= 0:
                stud_sel = line[line.find("(") + 1:line.find("%")]
            elif line.find("Student selectivity rank:") >= 0:
                stud_sel_rank = extract_score(line)
            elif line.find("SAT/ACT 25th-75th percentile:") >= 0:
                if line.find("]") >= 0:
                    sat_act_perc = line[line.find("]") + 2:]
                else:
                    sat_act_perc = line[line.find(":") + 2:]
            elif line.find("Freshmen in top 10 percent of high school class:") >= 0:
                if line.find("]") >= 0:
                    fresh_top_ten_hs = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    fresh_top_ten_hs = line[line.find(":") + 2:line.find("%")]
                else:
                    fresh_top_ten_hs = line[line.find(":") + 2:]
            elif line.find("Freshmen in top 25 percent of high school class:") >= 0:
                if line.find("]") >= 0:
                    fresh_top_25_hs = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    fresh_top_25_hs = line[line.find(":") + 2:line.find("%")]
                else:
                    fresh_top_25_hs = line[line.find(":") + 2:]
            elif line.find("Fall 2016 acceptance rate:") >= 0:
                if line.find("]") >= 0:
                    fall_2016_acc_rate = line[line.find("]") + 2:line.find("%")]
                else:
                    fall_2016_acc_rate = line[line.find(":") + 2:line.find("%")]
            elif line.find("Financial Resources (") >= 0:
                fin_res = line[line.find("(") + 1:line.find("%")]
            elif line.find("Financial resources rank:") >= 0:
                fin_res_rank = extract_score(line)

        # Putting values into a list of lists
        # NOTE: Have to have them in this format for CSV to export correctly
        values = [[school_name, school_rank, usn_score, grad_rate_perf,
                   predict_grad_rate, over_under_perf, alum_giving, alum_give_rank,
                   avg_alum_give_rate, grad_and_ret_rates, grad_and_ret_rank,
                   six_year_grade_rate, avg_fresh_ret_rate, undergrad_acad_rep,
                   peer_assess_score, hs_coun_score, fac_res, fac_res_rank,
                   perc_fac_fulltime, fulltime_fac_phd_term_deg,
                   classes_fewer_twenty, classes_fifty_more, stud_fac_ratio,
                   stud_sel, stud_sel_rank, sat_act_perc, fresh_top_ten_hs,
                   fresh_top_25_hs, fall_2016_acc_rate, fin_res, fin_res_rank]]

        return values
    elif dir_name == "north" or "south" or "midwest" or "west":

        # For every line get values
        for line in data.splitlines():

            if line.find("is ranked") >= 0:
                school_name = line[:line.find("is ranked")]
                school_rank = line[line.find("#") + 1:line.find("in Regional")].rstrip()
            elif line.find("Score:") >= 0:
                usn_score = extract_score(line)
            elif line.find("Graduation and Retention Rates") >= 0:
                grad_ret_rate_perc = line[line.find("(") + 1:line.find("%")]
            elif line.find("Graduation and retention rank") >= 0:
                grad_ret_rank = line[line.find(":") + 2:]
            elif line.find("Average 6-year graduation rate") >= 0:
                if line.find("]") >= 0:
                    avg_six_grad_rate = line[line.find("]") + 2:line.find("%")]
                elif line.find("%"):
                    avg_six_grad_rate = line[line.find(":") + 2:line.find("%")]
                else:
                    avg_six_grad_rate = line[line.find(":") + 2:]
            elif line.find("6-year graduation rate of students who received") >= 0:
                if line.find("%") >= 0:
                    six_year_received_pell = line[line.find(":") + 2:line.find("%")]
                else:
                    six_year_received_pell = line[line.find(":") + 2:]
            elif line.find("6-year graduation rate of students who did not") >= 0:
                if line.find("%") >= 0:
                    six_year_no_pell = line[line.find(":") + 2:line.find("%")]
                else:
                    six_year_no_pell = line[line.find(":") + 2:]
            elif line.find("Difference between graduation rates of Pell") >= 0:
                if line.find("%") >= 0:
                    diff_pell_nonpell = line[line.find(":") + 2:line.find("%")]
                else:
                    diff_pell_nonpell = line[line.find(":") + 2:]
            elif line.find("student retention rate") >= 0:
                if line.find("[") >= 0:
                    avg_firstyear_ret_rate = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    avg_firstyear_ret_rate = line[line.find(":") + 2:line.find("%")]
                else:
                    avg_firstyear_ret_rate = line[line.find(":") + 2:]
            elif line.find("Graduation Rate Performance") >= 0:
                grad_rate_perf = line[line.find("(") + 1:line.find("%")]
            elif line.find("Predicted graduation rate") >= 0:
                pred_grad_rate = line[line.find(":") + 2:line.find("%")]
            elif line.find("Overperformance(+)/Underperformance(-)") >= 0:
                over_under_perf = line[line.find(":") + 2:]
            elif line.find("Expert Opinion") >= 0:
                expert_opinion = line[line.find("(") + 2:line.find("%")]
            elif line.find("Peer assessment score (out of 5)") >= 0:
                peer_assess_score = line[line.find(":") + 2:]
            elif line.find("High school counselor score") >= 0:
                hs_coun_score = line[line.find(":") + 2:]
            elif line.find("Faculty Resources (") >= 0:
                fac_res = line[line.find("(") + 1:line.find("%")]
            elif line.find("Faculty Resources Rank") >= 0:
                fac_res_rank = line[line.find(":") + 2:]
            elif line.find("of faculty who are full-time") >= 0:
                if line.find("[") >= 0:
                    perc_fac_fulltime = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    perc_fac_fulltime = line[line.find(":") + 2:line.find("%")]
                else:
                    perc_fac_fulltime = line[line.find(":") + 2:]
            elif line.find("Full-time faculty with Ph.D or terminal degree") >= 0:
                if line.find("%") >= 0:
                    fulltime_fac_phd_term_deg = line[line.find(":") + 2:line.find("%")]
                else:
                    fulltime_fac_phd_term_deg = line[line.find(":") + 2:]
            elif line.find("Classes with fewer than 20 students") >= 0:
                if line.find("[") >= 0:
                    classes_fewer_twenty = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    classes_fewer_twenty = line[line.find(":") + 2:line.find("%")]
                else:
                    classes_fewer_twenty = line[line.find(":") + 2:]
            elif line.find("Classes with 50 or more students") >= 0:
                if line.find("[") >= 0:
                    classes_fifty_more = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    classes_fifty_more = line[line.find(":") + 2:line.find("%")]
                else:
                    classes_fifty_more = line[line.find(":") + 2:]
            elif line.find("Student-faculty ratio") >= 0:
                if line.find("]") >= 0:
                    stud_fac_ratio = line[line.find("]") + 2:]
                elif line.find("/") >= 0:
                    stud_fac_ratio = line[line.find(":") + 2:]
                else:
                    stud_fac_ratio = line[line.find(":") + 2:line.find(":1")]
            elif line.find("Student Excellence (") >= 0:
                stud_excel = line[line.find("(") + 1:line.find("%")]
            elif line.find("Student excellence rank") >= 0:
                stud_excel_rank = line[line.find(":") + 2:]
            elif line.find("SAT/ACT 25th-75th") >= 0:
                if line.find("[") >= 0:
                    sat_act_perc = line[line.find("]") + 2:]
                else:
                    sat_act_perc = line[line.find(":") + 2:]
            elif line.find("Freshmen in top 10 percent of high") >= 0:
                if line.find("[") >= 0:
                    fresh_top_ten_hs = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    fresh_top_ten_hs = line[line.find(":") + 2:line.find("%")]
                else:
                    fresh_top_ten_hs = line[line.find(":") + 2:]
            elif line.find("Freshmen in top 25 percent of high") >= 0:
                if line.find("[") >= 0:
                    fresh_top_25_hs = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    fresh_top_25_hs = line[line.find(":") + 2:line.find("%")]
                else:
                    fresh_top_25_hs = line[line.find(":") + 2:]
            elif line.find("Financial Resources (") >= 0:
                fin_res = line[line.find("(") + 1:line.find("%")]
            elif line.find("Financial resources rank") >= 0:
                fin_res_rank = line[line.find(":") + 2:]
            elif line.find("Alumni Giving (") >= 0:
                alum_giving = line[line.find("(") + 1:line.find("%")]
            elif line.find("Alumni giving rank") >= 0:
                alum_give_rank = line[line.find(":") + 2:]
            elif line.find("Average alumni giving rate") >= 0:
                if line.find("[") >= 0:
                    avg_alum_give_rate = line[line.find("]") + 2:line.find("%")]
                elif line.find("%") >= 0:
                    avg_alum_give_rate = line[line.find(":") + 2:line.find("%")]
                else:
                    avg_alum_give_rate = line[line.find(":") + 2:]

        values = [[school_name, school_rank, usn_score, grad_ret_rate_perc,
                   grad_ret_rank, avg_six_grad_rate, six_year_received_pell,
                   six_year_no_pell, diff_pell_nonpell,
                   avg_firstyear_ret_rate, grad_rate_perf, pred_grad_rate,
                   over_under_perf, expert_opinion, peer_assess_score,
                   hs_coun_score, fac_res, fac_res_rank, perc_fac_fulltime,
                   fulltime_fac_phd_term_deg, classes_fewer_twenty,
                   classes_fifty_more, stud_fac_ratio,
                   stud_excel, stud_excel_rank, sat_act_perc, fresh_top_ten_hs,
                   fresh_top_25_hs, fin_res, fin_res_rank, alum_giving,
                   alum_give_rank, avg_alum_give_rate
                   ]]

        return values

    print("Done gathering...")


# Extract Score from PDF
# NOTE: Does not work in some instances, using this function when applicable
def extract_score(line):
    for s in line.split():
        if s.isdigit():
            return int(s)


# Export Values to CSV path
def export_to_csv(export_dir, values, argument, folder):
    # File
    csv_file = export_dir + "\\" + os.path.basename(folder) + "Export.csv"

    # Write to File
    try:
        with open(csv_file, argument) as output:
            writer = csv.writer(output, lineterminator="\n")
            writer.writerows(values)
        output.close()
    except FileNotFoundError:
        print("File is either not found or broken.")
    except Exception as x:
        print("Unexpected error. Details: {}".format(x))


# Create Log for Files Skipped
def export_to_log(log_dir, file, lock, folder):
    # File
    log_file = log_dir + "\log" + os.path.basename(folder) + ".txt"

    # Write to File
    if lock:
        lock = False
        with open(log_file, "w") as logger:
            logger.write("Passed {}\n".format(file))
    else:
        with open(log_file, "a") as logger:
            logger.write("Passed {}\n".format(file))

    logger.close()

    return lock


# Main
if __name__ == "__main__":
    main()
