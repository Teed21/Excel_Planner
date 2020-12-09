# Written and Developed by: Tyler Wright
# Date started: 12/13/2018
# Date when workable: 02/05/2019
# Last Updated: 05/10/2019

import xlrd

"""

This class file is meant to help read in data from Excel files.
Mainly, this will be used for COMPANY_NAME documentation sent in CFA format.

"""

"""

    Programmer Notes:
    Row is up and down. Column is <-->.
    
    Make initialize function to handle all the CP related junk when object is created.
    
    Left off on: Copied all data into two separate lists. Now copy to new excel file with font_index intact.
    Copy every single cell format information into a list to be passed with the cell text data as well. This will
    prevent missing single-cell strike-throughs on CPs that have those changes.


    03/14/2019 - might need to update find_cp to catch if there is only one occurrence of the CP name so other code
    later on doesn't break.
    
    
    03/15/2019 - Fixed section where it could have potentially messed up with new Test CPs.
    
    03/19/2019 - Re-coded copy_sp_station and copy_font to handle a single set of cp coordinates or multiple.
    
    03/29/2019 - Re-coded find_cp so it doesn't crash when an invalid CP is input.

"""


class ExcelReader:

    def __init__(self, excel_book_name):
        self.ExcelBookName = excel_book_name
        self.is_cgc_change = False
        # 19 is the font index that correlates to a strike-through in excel.
        self.FONT_INDEX_STRIKE = 19
        # 6 is the font index that correlates to a bold in excel.
        self.FONT_INDEX_NEW_CHANGE = 6
        # Value for whether or not a station needs a full rewrite and has a strike-through version.
        self.is_full_change = False
        self.station_address = 0
        # List of rows necessary for copying CP station.
        self.list_of_rows = []
        # Column where all indications and controls are found.
        self.INDICATION_COLUMN = 3

        # Variables on Excel Sheets/Workbook
        self.first_sheet = ""
        self.workbook = ""
        self.excel_data = []
        self.excel_data_raw = []
        self.excel_format_data = []
        self.excel_format_data_raw = []
        self.cp_name = ""
        self.cp_names = []


    def test(self):
        
        # Open excel workbook
        book = xlrd.open_workbook(self.ExcelBookName, formatting_info=True)

        print(book.nsheets)

        print(book.sheet_names())

        first_sheet = book.sheet_by_index(0)

        print(first_sheet.name)

        cp_column = first_sheet.col_values(1)

        """
        for index in cp_column:
            if self.cp_name in index:
                print("Yes")
            else:
                print("NO")
        """

        # Reading in entire excel document into a list of lists.
        # Rows come first, then column.
        excel_data = [[first_sheet.cell_value(r, c) for c in range(first_sheet.ncols)] for r in range(first_sheet.nrows)]

        print(excel_data)
        print("\n\n")
        print("Data: " + excel_data[5][1])

        self.find_cell_formatting(book, first_sheet, 29, 1)

        self.find_cell_formatting(book, first_sheet, 37, 1)

        self.find_cell_formatting(book, first_sheet, 37, 4)

        return 0

    # Strike-throughs have a font_index of 19 and xf_index of 167.
    # This includes both italics and the strike-through, which is typical formatting from COMPANY_NAME.
    # This function provides output for a certain cell inside of the workbook/sheet.
    # It indicates whether or not a specific type of formatting is used.
    def find_cell_formatting(self, workbook, sheet, row, column):
        cell = sheet.cell(row, column)
        xf = workbook.xf_list[cell.xf_index]
        # print(xf.dump())
        # print(xf.font_index)
        """
        if scan_type == 1:
            # Determines based on font_index if the request needs a full change or not.
            # As a strike-through on the CP means that the entire station has been strike-throughed, this means that the
            # entire station requires a full change and rewrite.
            if xf.font_index == self.FONT_INDEX_STRIKE: # or xf.font_index == self.FONT_INDEX_NEW_CHANGE: - this will error out for new test CPs.
                self.is_full_change = True
            else:
                self.is_full_change = False
        elif scan_type == 2:
                print("Skipping scan type - not performing special check on overall request to see if full change.")
        else:
            print("Incorrect scan_type in find_cell_formatting...")
        """

        return xf.font_index

    def create_workbook(self):
        # Open excel workbook with formatting information included.
        self.workbook = xlrd.open_workbook(self.ExcelBookName, formatting_info=True)

        return self.workbook

    def create_worksheet(self, workbook):
        # Reading in first sheet.
        self.first_sheet = workbook.sheet_by_index(0)

        return self.first_sheet

    # This function helps grab all CP names in the CFA.
    def grab_all_cp_names(self, cell, row, column):
        # This appends all possible CP names to the list "cp_names" to later show to the user.
        if row > 2 and column == 1 and cell is not "" and "/" not in cell and " " not in cell:
            self.cp_names.append(cell)

        return 0

    # Reads in the excel document into a list of lists.
    # Also reads in formatting information from each cell as a number.
    def read_excel_document(self, sheet):
        # Reading in entire excel document into a list of lists.
        # Rows come first, then column.
        self.excel_data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        """
        row_iter = 0
        # Accounting for extra elements in self.excel_data.
        for row in self.excel_data:
            column_iter = 0
            for column in row:
                column_iter += 1
                if column_iter >= 13 and column == "":
                    #print("Found extra item in list.", row)
                    row.pop(column_iter-1)
                    self.excel_data[row_iter] = row
            row_iter += 1
        """

        # For every cell in the sheet, get the cell text info.
        for row in range(sheet.nrows):
            for column in range(sheet.ncols):
                self.excel_data_raw = [sheet.cell_value(row, column)]
                #self.excel_data_raw.append([sheet.cell_value(row, column)])
                # This grabs all cp names in the CFA to display for user later.
                self.grab_all_cp_names(sheet.cell_value(row, column), row, column)

        excel_format_data_raw = []

        # For every cell in the sheet, get the cell formatting info.
        for row in range(sheet.nrows):
            for column in range(sheet.ncols):
                # List of borders to be added to cell.
                borders = []
                font_style = ""
                cell = sheet.cell(row, column)
                xf = self.workbook.xf_list[cell.xf_index]

                rd_xf = self.workbook.xf_list[self.first_sheet.cell_xf_index(row, column)]
                cell_font = self.workbook.font_list[rd_xf.font_index]
                """
                # Check font_index type: if known type replace with corresponding string.
                if xf.font_index is 19:
                    font_style = "italic_strikeout"
                elif xf.font_index is 6:
                    font_style = "bold"
                elif xf.font_index is 11:
                    font_style = "regular"
                else:
                    font_style = xf.font_index
                """
                if cell_font.struck_out is 1:
                    font_style = "italic_strikeout"
                elif cell_font.bold is 1:
                    font_style = "bold"
                elif xf.font_index is 11:
                    font_style = "regular"
                else:
                    font_style = xf.font_index

                # Check if borders are on cell
                if xf.border.bottom_line_style is not 0:
                    borders.append("Bottom")
                if xf.border.diag_line_style is not 0:
                    borders.append("Diag")
                if xf.border.left_line_style is not 0:
                    borders.append("Left")
                if xf.border.right_line_style is not 0:
                    borders.append("Right")
                if xf.border.top_line_style is not 0:
                    borders.append("Top")

                self.excel_format_data_raw.append([font_style, borders])

        # Appending all the formatting information into a relevant list (how long the text columns are).
        for index in range(0, len(self.excel_format_data_raw), len(self.excel_data[0])):
            self.excel_format_data.append(self.excel_format_data_raw[index:index + len(self.excel_data[0])])

        return self.excel_data

    # Finds the row and column that a CP name, or any item, is in.
    # Gets both strike-through and regular if necessary.
    # Sets self.cp_name.
    def find_cp(self, excel_data, cp_name):
        self.cp_name = cp_name
        coordinates = []
        position_y = 0
        position_x = 0
        cp_found = False
        for index in excel_data:
            if cp_name in index:
                cp_found = True
                #position_y += 1
                coordinates.append(position_y)
                position_y += 1
                for x in index:
                    if cp_name in x and position_x > 0:
                        coordinates.append(position_x)
                        position_x = 0
                        break
                    else:
                        position_x += 1
            elif cp_name not in index and position_y >= len(excel_data):
                cp_found = False
                print("CP name not found. Returning 0 so program does not crash.")
                return False
            else:
                position_y += 1
                #print("No")

        if cp_found is False:
            return False

        # Putting all of the items from the coordinate list into a dictionary for easier parsing.
        coordinate_dictionary = {}
        items = 0
        for cord in coordinates:
            if items % 2 == 0 or items == 0:
                # add both items as a pair into the dictionary. first number = key, second in pair = value.
                coordinate_dictionary[coordinates[items]] = coordinates[items + 1]
                # coordinate_dictionary.
            items += 1
            #print(cord)

        # This snippet of code determines if the Excel data reflects a full change of REO for one station/cp.
        # This is done by comparing how many times a CP name shows up in the file. If it shows up with more than 2
        # coordinates, then it is a full change, as one MUST be a strike-through and another must be the new station.
        if len(coordinates) > 2:
            self.is_full_change = True
        else:
            self.is_full_change = False

        return coordinate_dictionary

    # Find the text value of a certain cell.
    def find_cell_text_value(self, dataset, row, column):
        return dataset[row][column]

    # Returns a list of cell format indexes to determine what kind of formatting is inside the cell.
    def send_cords_to_find_cell_formatting(self, coordinate_dictionary, workbook, sheet):
        cell_formatting = []
        # iterating over dictionary to get every coordinate and find the formatting for the cell
        for key, value in coordinate_dictionary.items():
            cell_formatting.append(self.find_cell_formatting(workbook, sheet, key, value))

        return cell_formatting

    # Find relevant data about control point and its station in CFA.
    # Sets some extra variables based on the information collected.
    def collect_data_about_cp(self, dataset, cp_position):
        # List to help determine if CGC change is needed
        ctrls_and_indics = []
        #first_time = True

        # This for loop takes the already determined position of the CP and finds extra information based on that.
        for key, value in cp_position.items():
            self.station_address = dataset[key][value+1]
            ctrls_and_indics.append(dataset[key][value+2])

            # Splitting the "IND # / CON #" into two variables, one for the number and one for the name.
            # Using split method to get IND and # on their own to easily compare the two in their own if statements.
            ctrl_and_indic_name_and_number = dataset[key][value+2].split(" ")
            ctrls_and_indics_column_name = ctrl_and_indic_name_and_number[0]
            ctrls_and_indics_column_number = ctrl_and_indic_name_and_number[1]

            """
            # if first time through loop, and if the change is a full one, determine what key/row is for the
            # strike-through
            #if first_time and self.is_full_change:
            if len(cp_position) > 1:
                # Determines where the CP is.
                strike_through_row = key
                #first_time = False
            else:
                new_row = key
            """

            # IF statements to grab all rows for a CP.
            if "CON" in ctrls_and_indics_column_name:
                if "1" in ctrls_and_indics_column_number:
                    # print("Work way down until blank space, then subtract one from that row value.")
                    self.find_rows_for_station(dataset, key, self.INDICATION_COLUMN, 1)
                else:
                    # print("Work way down until blank space, subtract one from that row value. Also,"
                    #       " subtract current # until you reach the value of 1.")
                    # Finds highest number of rows.
                    self.find_rows_for_station(dataset, key, int(ctrls_and_indics_column_number), 4)

                    # Finds rest of rows for station.
                    self.find_rows_for_station(dataset, key, self.INDICATION_COLUMN, 1)

            elif "IND" in ctrls_and_indics_column_name:
                # print("Work way up until blank, down until blank.")
                self.find_rows_for_station(dataset, key, self.INDICATION_COLUMN, 3)
            else:
                print("Unknown parameter...")   # Add logging here.

        try:
            # Checking if CGC change is necessary
            if ctrls_and_indics[0] is not ctrls_and_indics[1]:
                self.is_cgc_change = True
            else:
                self.is_cgc_change = False
        except:
            print("No CGC change required.")    # ADD LOGGING HERE INSTEAD OF PRINT OUT.
            self.is_cgc_change = False

        return 0

    # Function that finds the rows in the CP station
    def find_rows_for_station(self, dataset, row, column, scan_type):
        blank = ""
        top = "CON/IND"

        # IF STATEMENT TO DETERMINE IF MORE THAN ONE PAIR OF COORDINATES.
        if len(self.list_of_rows) == 2:
            self.list_of_rows.append("")

        # While the current row and column is not a blank, add or subtract 1 to the row to continue through the excel
        # sheet.
        # This scan type goes down from the current row and finds the lowest indication.
        if scan_type == 1:
            if "CON 1" in self.find_cell_text_value(dataset, row, column):
                #print("Highest row: ", row)
                # Adding row to list for later calculations to find rest of rows.
                self.list_of_rows.append(row)
            while self.find_cell_text_value(dataset, row, column) is not blank and \
                    top not in self.find_cell_text_value(dataset, row, column):
                row += 1
            # Setting value of current row back to an indication row.
            if self.find_cell_text_value(dataset, row, column) is blank or \
                    top in self.find_cell_text_value(dataset, row, column):
                row -= 1
                # Checking to make sure the text value of the cell is "IND #"
                if "IND" in self.find_cell_text_value(dataset, row, column):
                    #print("Last row: ", row)
                    # Adding row to list for later calculations to find rest of rows.
                    self.list_of_rows.append(row)
                else:
                    print("Incorrect row math. Cannot find IND # at bottom of station.")

        # This scan type goes up from the current row and finds the highest control.
        elif scan_type == 2:
            while self.find_cell_text_value(dataset, row, column) is not blank and \
                    top not in self.find_cell_text_value(dataset, row, column):
                row -= 1
            # Setting value of current row back to an indication row.
            if self.find_cell_text_value(dataset, row, column) is blank or \
                    top in self.find_cell_text_value(dataset, row, column):
                row += 1
                # Checking to make sure the text value of the cell is "CON #"
                if "CON 1" in self.find_cell_text_value(dataset, row, column):
                    #print("Highest row: ", row)
                    # Adding row to list for later calculations to find rest of rows.
                    self.list_of_rows.append(row)
                else:
                    print("Incorrect row math. Cannot find CON # at bottom of station.")

        # This scan type goes up and down from the current row and finds the highest and lowest control.
        elif scan_type == 3:
            reset_scan_row = row
            # Going up.
            while self.find_cell_text_value(dataset, row, column) is not blank and \
                    top not in self.find_cell_text_value(dataset, row, column):
                #if top in self.find_cell_text_value(dataset, row, column):
                #    break
                row -= 1
            # Setting value of current row back to an indication row.
            if self.find_cell_text_value(dataset, row, column) is blank or \
                    top in self.find_cell_text_value(dataset, row, column):
                row += 1
                # Checking to make sure the text value of the cell is "CON #"
                if "CON 1" in self.find_cell_text_value(dataset, row, column):
                    #print("Highest row: ", row)
                    # Adding row to list for later calculations to find rest of rows.
                    self.list_of_rows.append(row)
                else:
                    print("Incorrect row math. Cannot find CON # at bottom of station.")
            # Resetting the row to the first row scanned from to get precise row calculations and take less time doing
            # so.
            row = reset_scan_row
            # Going down.
            while self.find_cell_text_value(dataset, row, column) is not blank and \
                    top not in self.find_cell_text_value(dataset, row, column):
                row += 1
            # Setting value of current row back to an indication row.
            if self.find_cell_text_value(dataset, row, column) is blank or \
                    top in self.find_cell_text_value(dataset, row, column):
                row -= 1
                # Checking to make sure the text value of the cell is "IND #"
                if "IND" in self.find_cell_text_value(dataset, row, column):
                    #print("Last row: ", row)
                    # Adding row to list for later calculations to find rest of rows.
                    self.list_of_rows.append(row)
                else:
                    print("Incorrect row math. Cannot find IND # at bottom of station.")
            # Adding blank value between old CP and new CP for easier access.
            #self.list_of_rows.append("")

        elif scan_type == 4:
            # Finds highest row possible.
            first_row = row - int(column) + 1
            # Setting column back to regular column placement
            column = self.INDICATION_COLUMN
            # Checks if the row is the row on CON 1.
            if "CON 1" in self.find_cell_text_value(dataset, first_row, column):
                #print("Found highest row.")
                self.list_of_rows.append(first_row)
            else:
                print("Incorrect row math. Cannot find CON 1.")
        else:
            print("Incorrect scan type selected. Errors will occur. Aborting function. Please correct the scan_type"
                  " value being passed to this function.")

        return 0

    # Copies the entire station for a CP, whether it is a strike-through or not.
    # Will only copy relevant data that needs changed. - Can copy whatever you want, though. Just give it a proper CP
    # and run the CP through its paces in the other functions first.
    def copy_cp_station(self, dataset, rows):
        # These lists will contain all the data for the CP stations.
        first_station = []
        second_station = []
        if len(rows) <= 2:
            first_list = rows
            first_list_iterations = first_list[1] - first_list[0] + 1
            # Goes through each row and copies the information of each list in its entirety to a new list.
            for index in range(first_list_iterations):
                first_station.append(dataset[first_list[0] + index])
        elif len(rows) > 2:
            # split array of rows into two arrays of two values each.
            row_split = rows.index("")
            first_list = rows[:row_split]
            first_list_iterations = first_list[1] - first_list[0] + 1
            # Goes through each row and copies the information of each list in its entirety to a new list.
            for index in range(first_list_iterations):
                first_station.append(dataset[first_list[0]+index])
            # If there is a second list, repeat above and copy into new list.
            second_list = rows[row_split:]
            # Removing blank "" from second list.
            second_list.remove("")
            second_list_iterations = second_list[1] - second_list[0] + 1
            for index in range(second_list_iterations):
                second_station.append(dataset[second_list[0] + index])

        # Combining both stations.
        combined_stations = []
        blank = [""]
        for row in range(len(first_station)):
            combined_stations.append(first_station[row])

        # This blank puts a space between stations when writing to excel file.
        combined_stations.append(blank)

        # Error check on second station - making sure it doesn't fill excel sheet with blanks or give indicy error.
        if len(rows) > 2:
            for row in range(len(second_station)):
                combined_stations.append(second_station[row])

        return combined_stations

    # Copies all font information for each cell for the CPs.
    def copy_cp_station_font_info(self, rows, sheet):
        # list for appending all format info for CPs.
        cp_format_info = []
        first_list = []
        second_list = []
        if len(rows) <= 2:
            first_list = rows
        elif len(rows) > 2:
            # Split array of rows into two arrays of two values each.
            row_split = rows.index("")
            first_list = rows[:row_split]
            second_list = rows[row_split:]

        row_count = 0

        for row in self.excel_format_data:
            if row_count <= first_list[1] and row_count >= first_list[0]:
                cp_format_info.append(row)
            else:
                tmp = ""
                # print("Skipped row: ", row_count)

            if len(rows) > 2:
                if row_count == second_list[1] - 1:
                    cp_format_info.append([""])
                if row_count <= second_list[2] and row_count >= second_list[1]:
                    cp_format_info.append(row)
                else:
                    tmp = ""
                    # print("Skipped row: ", row_count)
            row_count += 1

        return cp_format_info

    # Copies all codeline info from the first three rows.
    def copy_codeline_text_info(self, dataset):
        codeline_info = []
        row_count = 0
        for row in dataset:
            # If reach end of necessary rows (3), quit and return codeline info.
            if row_count >= 3:
                break
            codeline_info.append(dataset[row_count])
            row_count += 1

        # Formatting the codeline type info - shortening it to fit in excel cell.
        codeline_type = codeline_info[1][1]

        if "placeholder1" in codeline_type:
            codeline_info[1][1] = "placeholder1"
        elif "placeholder2" in codeline_type:
            codeline_info[1][1] = "placeholder2"
        elif "placeholder3" in codeline_type:
            codeline_info[1][1] = "placeholder3"
        else:
            print("Could not find codeline type text in ExcelReader(copy_codeline_text_info)")

        return codeline_info

    # Copies all codeline formatting information.
    def copy_codeline_format_info(self, format_data):
        codeline_formatting = []

        row_count = 0

        for row in format_data:
            if row_count == 3:
                break
            codeline_formatting.append(format_data[row_count])
            row_count += 1

        return codeline_formatting

    # Combines existing CP info list with new CPs.
    def combine_cp_text_info(self, cp_info_one, cp_info_two):
        combined_cp_info = []

        for row in range(len(cp_info_one)):
            combined_cp_info.append(cp_info_one[row])

        combined_cp_info.append([""])

        for row in range(len(cp_info_two)):
            combined_cp_info.append(cp_info_two[row])

        return combined_cp_info

    # Combines existing format info with new format info.
    def combine_format_info(self, format_info_one, format_info_two):
        format_info = []

        for row in range(len(format_info_one)):
            format_info.append(format_info_one[row])

        format_info.append([""])

        for row in range(len(format_info_two)):
            format_info.append(format_info_two[row])

        return format_info

    # Returns whether CGC change is needed or not in the form of "True" or "False"
    def get_cgc_change(self):
        return self.is_cgc_change

    # Returns the station address, either a 0 for unset, or a number for the station.
    def get_station_address(self):
        return self.station_address

    # Sets the list of rows.
    def set_station_rows(self, rows):
        self.list_of_rows = rows

    # Returns the list of rows for a station(s)
    def get_station_rows(self):
        return self.list_of_rows

    # Returns "True" for full change needed, or "False" for not needed.
    def get_is_full_change(self):
        return self.is_full_change

    # Returns list of all formatting information
    def get_excel_format_data(self):
        return self.excel_format_data

    # Returns specific CP name.
    def get_cp_name(self):
        return self.cp_name

    # Returns the list of possible CP names
    def get_cp_names(self):
        filtered_list = []
        for cp in self.cp_names:
            if cp is not " ":
                filtered_list.append(cp)
        return filtered_list

    # Returns the list of possible CP names
    def get_cp_names_and_suggested(self):
        return self.cp_names
