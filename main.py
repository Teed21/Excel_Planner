# Written and Developed by: Tyler Wright
# Date started: 12/13/2018
# Date when workable: 02/05/2019
# Last Updated: 04/26/2019

import ExcelPlanner
import EmailParser
# This is for open dialog window.
import tkinter as tk
from tkinter import filedialog
import os
import datetime as date


# These variables set up an invisible window that hosts the file dialog.
root = tk.Tk()
root.withdraw()

# Example of file dialog. Use when necessary for getting files.
#file_path = filedialog.askopenfilename()

spacing = "\n"

# This object handles finding the COMPANY_NAME email in user's outlook application.
up_email = EmailParser.EmailParser()

excel_request_name = "_REO_Testplan-First.xlsx"

# This function helps check if a CP exists before adding it to a list.
def check_if_cp_exists(cfa, cp_name):
    planner = ExcelPlanner.ExcelPlanner(cfa)

    cp_exists = planner.check_if_cp_exists(cp_name)

    return cp_exists


# This function gets all cp names in the selected CFA/Excel doc.
def get_cp_names(cfa):

    planner = ExcelPlanner.ExcelPlanner(cfa)

    cp_names = planner.get_cp_names()

    return cp_names


# This function will handle multiple CFA/ExcelPlanner objects.
def multiple_cfa(cfa, cp_names):
    # Create planner object using cfa name.
    planner = ExcelPlanner.ExcelPlanner(cfa)

    # Generate info based on cp names.
    list_of_gen_cp_info = planner.get_cp_cords(cp_names)
    list_of_cp_cords = list_of_gen_cp_info[0]
    list_of_cp_names = list_of_gen_cp_info[1]

    # Create list of cp objects to manipulate and combine texts from later on.
    cp_objects = planner.get_all_cp_information(list_of_cp_cords, list_of_cp_names)

    # This list is the combination of all text needed for a new CFA.
    lists = planner.return_combined_lists(cp_objects)

    return lists


# This function handles all suggested CPs by comparing what is inside the CFA, and what was
# on the email sent by COMPANY_NAME.
def filter_suggested_cps(cfa_cps, email_cps):
    suggested_cps = []
    for cp in cfa_cps:
        if cp in email_cps:
            if cp in suggested_cps:
                continue
            else:
                suggested_cps.append(cp)
    if suggested_cps:
        return suggested_cps
    else:
        return "No suggested CPs."


# This code block handles the request name, as well as parsing for the S request email.
print("What is the request number? (EX: S031)\n"
      "\tNote: This will be used to find the relevant email in your Outlook email.")
request_name = input("Request Name: ")

# Getting current year to filter emails
current_year = date.datetime.now()
current_year = current_year.year

# Finding email based on what subfolder in outlook and current year + request name.
up_emails = up_email.find_email("Requests", "Request " + str(current_year) + request_name)
new_excel_file_name = request_name
email_loop = up_emails
while email_loop is False:
    print("\nProblem finding email in Outlook to find suggested CPs.")
    correct_email = input("Would you like to:\n"
                          "1. Try to find email again.\n"
                          "2. Continue anyway.\n"
                          ":")
    if int(correct_email) == 1:
        request_name = input("Request Name: ")
        up_emails = up_email.find_email("Requests", "Request " + str(current_year) + request_name)
        email_loop = up_emails
    elif int(correct_email) == 2:
        print("Continuing...", spacing)
        email_loop = True
    else:
        print("Input not recognized. Please choose a proper option.")



# This code block gets the amount of CFAs from the user. Will continue to ask until correct input is met.
is_number = False
while is_number is not True:
    try:
        cfa_input = int(input("How many CFAs are there? 1-3\n:"))
        if cfa_input > 3 or cfa_input <= 0:
            print("Too many CFAs. Can only handle 1 to 3 CFAs.", spacing)
        else:
            print(cfa_input, "CFAs will be processed.", spacing)
            is_number = True
    except ValueError:
        print("Data Type incorrect for cfa_input. Must be a number.", spacing)


# This code block gets the names of the CFAs and the CPs related to them.
list_of_planners = []
processed = False
while processed is not True:
    # For every cfa, ask for a file name.
    for cfa_num in range(cfa_input):
        list_of_cps = []
        # User is prompted by pop-up to select a file.
        print("Please select your CFA file.", spacing)
        file_path = filedialog.askopenfilename()
        file_path_list = file_path.split("/")
        cfa_file = file_path_list[-1]
        print("File selected:", cfa_file, " in path:", file_path, spacing)
        # Checking if file exists.
        if os.path.isfile(file_path):
            # Output all possible CP names
            print("Found these CP names in file:")
            print(get_cp_names(file_path), spacing)
            # Outputting suggested cps.
            print("Suggested CPs in this CFA:")
            if up_emails is False:
                print("No suggested CPs to display.")
            else:
                print(filter_suggested_cps(get_cp_names(file_path), up_email.find_cps(up_emails)), spacing)
            # Asking for all CPs for specific CFA
            print("Type in all the CPs for this CFA file one-by-one. When done, type the letter 'd' to continue."
                  , spacing)
            user_input = ""
            while "d" not in user_input or "D" not in user_input and len(user_input) > 1:
                user_input = input("CP: ")
                if user_input == "":
                    print("Blank provided. Not sending blank to CP stack.")
                elif user_input != "d" and user_input != "D":
                    # Check if CP exists in file.
                    if check_if_cp_exists(file_path, user_input) is False:
                        print("CP name not found in file provided: ", file_path)
                        print("Please enter a different CP name...", spacing)
                        continue
                    list_of_cps.append(user_input)
                else:
                    print("Continuing...", spacing)

            list_of_planners.append(multiple_cfa(file_path, list_of_cps))
        else:
            # This should pull the user back into the while loop, and go through listing CFAs again.
            print("It looks like the file you tried to input does not exist. Try again.", spacing)
            processed = False
            break

        # This will break the user out of the while loop so they can continue.
        processed = True

# Process all planner information into one list for the writer to handle
master_list = []
first_time = False
for planner_list in list_of_planners:
    if first_time is False:
        first_time = True
        master_list += planner_list
    else:
        master_list[0] += planner_list[0]
        master_list[1] += planner_list[1]

# Create new file
excel_file_loop = True
while excel_file_loop:
    excel_file_decision = input("File is currently named: " + new_excel_file_name + "\n"
                                "Would you like to change the name?\n"
                                "1. Yes\n"
                                "2. No\n"
                                ":")
    if int(excel_file_decision) == 1:
        new_excel_file_name = input("Excel file name: ")
        excel_file_loop = False
    elif int(excel_file_decision) == 2:
        print("Excel file will be named: " + new_excel_file_name)
        excel_file_loop = False
    else:
        print("Input not recognized, please select a proper option.")
print("Please select where you want to save your new CFA file.")
new_file_path = filedialog.askdirectory()
print("Creating new file at:\n ", new_file_path)
new_file = new_file_path + "/" + request_name + excel_request_name

temp_planner = ExcelPlanner.ExcelPlanner("new")

temp_planner.process_and_write_all_cp_information(master_list, new_file)

print("Done.")
