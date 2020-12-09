# Written and Developed by: Tyler Wright
# Date started: 12/13/2018
# Date when workable: 02/05/2019
# Last Updated: 04/10/2019

import xlsxwriter

"""

This class file is meant to help write in data from Excel files.
Mainly, this will be used for COMPANY_NAME documentation sent in CFA format.

"""

"""

Programmer Notes:
    
    Left off on: Copied all data into two separate lists. Now copy to new excel file with font_index intact.
    Copy every single cell format information into a list to be passed with the cell text data as well. This will
    prevent missing any single-cell strike-through(s) on CPs that have those changes.
    
    
    3/21/2019 - changed line 121 reading "elif row <= 3:" to "elif row < 3:" - this fixed making one line bold 
    constantly.

"""


class ExcelWriter:

    def __init__(self, cp_text_data, cp_format_data, new_workbook_name):
        # Used to write CP information with appropriate formatting.
        self.cp_text_data = cp_text_data
        self.cp_format_data = cp_format_data

        self.new_workbook_name = new_workbook_name

        # Used for writing - if excel file already exists, open it and write to it.
        self.workbook = ""
        self.plan_sheet = ""

    # Resets formatting for cells to whatever is passed in the arguments.
    def reset_format(self, format_to_reset, font_name, font_color, font_type, workbook):
        format_to_reset = ""

        if "none" in font_type:
            format_to_reset = workbook.add_format({'font_color': font_color})
        else:
            format_to_reset = workbook.add_format({font_type: True, 'font_color': font_color})

        format_to_reset.set_font_name(font_name)
        format_to_reset.set_align("center")

        if font_type == "font_strikeout":
            format_to_reset.set_italic(True)

        return format_to_reset

    # Writes to the excel sheet. This function is to be used inside of a list.
    # It will help iterate through the list and write the cell data with proper formatting.
    """
        Parameters:
            workbook = book to open new excel file.
            plan_sheet = individual sheet to write on.
            cell_format_info = this is the formatting information for the current cell.
            row = current row in the list
            column = current column in the list.
    """
    def write_format_checker(self, workbook, sheet, cell_format_info, row, column):
        # Cell formatting
        # Strike
        strike_format = workbook.add_format({'font_strikeout': True, 'font_color': 'black'})
        strike_format.set_italic(True)
        strike_format.set_font_name("Arial")
        strike_format.set_align("center")

        # Bold
        bold_format = workbook.add_format({'bold': True, 'font_color': 'black'})
        bold_format.set_font_name("Arial")
        bold_format.set_align("center")

        # Default
        default_format = workbook.add_format({'font_color': 'black'})
        default_format.set_font_name("Arial")
        default_format.set_align("center")

        # Codeline Formatting
        codeline_format = workbook.add_format({'bold': True, 'font_color': 'black'})
        codeline_format.set_font_name("Arial")
        codeline_format.set_align("center")

        # If format in text_and_format_data matches one of the formatting #'s, write that cell in the format
        if "italic_strikeout" in cell_format_info:
            # Check border properties and assign accordingly.
            if "Top" in cell_format_info[1]:
                strike_format.set_top()
            if "Bottom" in cell_format_info[1]:
                strike_format.set_bottom()
            if "Right" in cell_format_info[1]:
                strike_format.set_right()
            if "Left" in cell_format_info[1]:
                strike_format.set_left()
            # Writes to new excel file as (row(x), column(y), text, format)
            sheet.write(row, column, self.cp_text_data[row][column], strike_format)
            # Reset formatting
            strike_format = self.reset_format(strike_format, "Arial", "black", "font_strikeout", workbook)
        elif "bold" in cell_format_info:
            # Check border properties and assign accordingly.
            if "Top" in cell_format_info[1]:
                bold_format.set_top()
            if "Bottom" in cell_format_info[1]:
                bold_format.set_bottom()
            if "Right" in cell_format_info[1]:
                bold_format.set_right()
            if "Left" in cell_format_info[1]:
                bold_format.set_left()
            sheet.write(row, column, self.cp_text_data[row][column], bold_format)
            # Reset formatting
            bold_format = self.reset_format(bold_format, "Arial", "black", "bold", workbook)
        # This if statement is specifically for the first three columns in the excel file. Everything will be bolded.
        elif row < 3:
            # Check border properties and assign accordingly.
            if "Top" in cell_format_info[1]:
                codeline_format.set_top()
            if "Bottom" in cell_format_info[1]:
                codeline_format.set_bottom()
            if "Right" in cell_format_info[1]:
                codeline_format.set_right()
            if "Left" in cell_format_info[1]:
                codeline_format.set_left()
            sheet.write(row, column, self.cp_text_data[row][column], codeline_format)
            # Reset formatting
            codeline_format = self.reset_format(codeline_format, "Arial", "black", "bold", workbook)
        else:
            # Check border properties and assign accordingly.
            if "Top" in cell_format_info[1]:
                default_format.set_top()
            if "Bottom" in cell_format_info[1]:
                default_format.set_bottom()
            if "Right" in cell_format_info[1]:
                default_format.set_right()
            if "Left" in cell_format_info[1]:
                default_format.set_left()
            sheet.write(row, column, self.cp_text_data[row][column], default_format)
            # Reset formatting
            default_format = self.reset_format(default_format, "Arial", "black", "none", workbook)
        return 0

    # Combines the lists of data to write to the excel sheet.
    # Perform check on lists sent over to see if they are empty or not.
    def write_to_excel(self):
        # Creating workbook to write to.
        self.workbook = xlsxwriter.Workbook(self.new_workbook_name)
        # Creating worksheets to individually write to.
        test_sheet = self.workbook.add_worksheet('Testing')
        self.plan_sheet = self.workbook.add_worksheet('Planning')

        # Setting column widths.
        self.plan_sheet.set_column(0, 0, 15.11)
        self.plan_sheet.set_column(1, 1, 11.67)
        self.plan_sheet.set_column(2, 2, 19.89)
        self.plan_sheet.set_column(3, 3, 12.33)
        self.plan_sheet.set_column(4, 12, 13.89)

        test_sheet.set_column(0, 0, 15.11)
        test_sheet.set_column(1, 1, 11.67)
        test_sheet.set_column(2, 2, 19.89)
        test_sheet.set_column(3, 3, 12.33)
        test_sheet.set_column(4, 12, 13.89)

        # For every row in the cp_text_data.
        for row in range(len(self.cp_text_data)):
            # Catch blank spot meant to separate differing CPs.
            if len(self.cp_text_data[row]) <= 1 and "" in self.cp_text_data[row]:
                # Add a special case here that writes a blank row to the excel file.
                self.write_format_checker(self.workbook, self.plan_sheet, ['regular', ['']], row,
                                          0)
                continue
            # For every column in the row.
            for column in range(len(self.cp_text_data[row])):
                #print(self.cp_text_data[row][column], " ", row, " ", column)
                cell_format_info_current_cp_cell = self.cp_format_data[row][column]
                self.cell_text = self.cp_text_data[row][column]
                #print(self.cell_text)
                self.write_format_checker(self.workbook, self.plan_sheet, cell_format_info_current_cp_cell, row,
                                          column)

        # Merge miscellaneous cells here for codeline stuff. - maybe

        self.workbook.close()

        return 0
