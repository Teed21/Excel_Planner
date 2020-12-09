# Written and Developed by: Tyler Wright
# Date started: 02/05/2019
# Date when workable: 03/21/2019
# Last Updated: 04/26/2019

"""
    This class file is meant to operate as a median between ExcelReader and ExcelWriter.

    It will handle multiple CP values, instead of handling simply one run-through. It will also
    help keep track of any data specific to the overall request.

    This class also creates individual CPs based on information pulled from ExcelReader. These CPs
    should then be stored in a list, later to be called upon to conglomerate their information into
    something that ExcelWriter can read and understand.

    This class also hosts a GUI. The desired outcome is to have this be a much easier interface to
    use and understand. Notes will be displayed in a text box for the user to watch and monitor the
    program to make sure everything is working as planned. -- This was not added. Deemed unnecessary.


    Find out how to put spaces between stations in new file.


"""

import ExcelReader
import ExcelWriter
import ExcelCP


class ExcelPlanner:

    def __init__(self, excel_file):
        # This variable is the excel file name to open.
        self.excel_file = excel_file
        # This if statement helps skip the constructor if need be. - Mainly used to write to new excel file.
        if excel_file == "new":
            print("Creating new excel file...")
        else:
            # This variable is the object that controls all information from ExcelReader.py that is necessary to
            # mediate everything in this class file. Opening the excel file, getting the text data, etc.
            self.excel_reader_obj = ExcelReader.ExcelReader(excel_file)

            # This variable is a global workbook. It is meant to be used throughout the class as a way to
            # reference the workbook opened for the current task of copying from and pasting to a new workbook/file.
            self.glb_workbook = self.excel_reader_obj.create_workbook()

            # This variable grabs the relevant sheet.
            self.glb_sheet = self.excel_reader_obj.create_worksheet(self.glb_workbook)

            # This variable grabs the relevant text data from the sheet.
            self.glb_data = self.excel_reader_obj.read_excel_document(self.glb_sheet)

            # This global variable holds all format information for the data sheet.
            self.glb_format_data = self.excel_reader_obj.get_excel_format_data()

            # These lists are meant for putting spaces in between CP stations in the new Excel file.
            self.space_text = ["", "", "", "", "", "", "", "", "", "", "", ""]
            self.space_format = [["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]],
                                 ["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]],
                                 ["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]], ["regular", ["Top"]]]

    # This function gets all codeline information and combines it into one.
    def get_codeline_info(self):
        codeline_text = self.excel_reader_obj.copy_codeline_text_info(self.glb_data)
        codeline_format = self.excel_reader_obj.copy_codeline_format_info(self.glb_format_data)

        return [codeline_text, codeline_format]

    # This function checks if a CP exists before appending to any lists or permanent data structures.
    # This will cleanse any non-sense in regards to invalid CP names in later data.
    def check_if_cp_exists(self, cp_name):
        is_cp = self.excel_reader_obj.find_cp(self.glb_data, cp_name)
        return is_cp

    # This function returns all possible cp names in the selected excel file.
    def get_cp_names(self):
        cp_names = self.excel_reader_obj.get_cp_names()
        return cp_names

    # This function passes over a list of CPs required to be copied over to a new sheet. Getting all relevant
    # information from them and storing it in a large list.
    # When processing this list, remember to count whether or not this is a "full change" with a strike-out and bold
    # version of the station. This is important for copying old and new text and format information later on.
    def get_cp_cords(self, cp_list):
        list_of_cp_cords = []
        list_of_cp_names = []
        for cp in cp_list:
            list_of_cp_cords.append(self.excel_reader_obj.find_cp(self.glb_data, cp))
            list_of_cp_names.append(self.excel_reader_obj.get_cp_name())

        return list_of_cp_cords, list_of_cp_names

    # Creates a CP object by using the coordinates from get_all_cp_information and processing information from
    # ExcelReader based on those coordinates.
    def create_cp_object(self, coordinates, cp_name, occurrences):
        # Collect data from collect_data_about_cp - this will set the class var "rows" for each coordinate.
        # This will help produce relevant information (text and formatting) for each CP.
        self.excel_reader_obj.collect_data_about_cp(self.glb_data, coordinates)

        # Get CP name.
        cp_name = cp_name

        # Get CP coordinates
        cp_coordinates = coordinates

        # Get CP station address
        cp_station_address = self.excel_reader_obj.get_station_address()

        # Set rows for specific CP
        rows = self.excel_reader_obj.get_station_rows()

        # Copy station text for specific CP
        cp_text_data = self.excel_reader_obj.copy_cp_station(self.glb_data, rows)
        # Copy station font/format for specific CP
        cp_font_data = self.excel_reader_obj.copy_cp_station_font_info(rows, self.glb_sheet)

        if occurrences == 1:
            text_list_bold = []
            format_list_bold = []

            iteration = 0
            for row in range(len(cp_text_data)):
                iteration += 1
                # Catch blank spot meant to separate differing CPs.
                if len(cp_text_data[row]) <= 1 and "" in cp_text_data[row]:
                    # Split at one occurrence.
                    break
                # Append strike_rows to new list.
                text_list_bold.append(cp_text_data[row])

            # Adding blank spacing to text list.
            text_list_bold.append(self.space_text)

            iteration = 0
            for row in range(len(cp_font_data)):
                iteration += 1
                # Catch blank spot meant to separate differing CPs.
                if len(cp_font_data[row]) <= 1 and "" in cp_font_data[row]:
                    # Split at one occurrence.
                    break
                # Append strike_rows to new list.
                format_list_bold.append(cp_font_data[row])

            # Adds a blank in the format list to line up with the text lists.
            format_list_bold.append(self.space_format)

            # Setting font type to 's' for single.
            font_type = "s"
            # Create cp object.
            cp_object = ExcelCP.ExcelCP(cp_name, cp_coordinates, font_type, cp_station_address, "N/A",
                                        "N/A", text_list_bold, format_list_bold)
        elif occurrences == 2:
            # Setting font type to 'b' for both.
            font_type = "b"

            # Setting lists to split information into.
            text_list_strike = []
            text_list_bold = []
            format_list_strike = []
            format_list_bold = []

            iteration = 0
            # For every row in the cp_text_data.
            for row in range(len(cp_text_data)):
                iteration += 1
                # Catch blank spot meant to separate differing CPs.
                if len(cp_text_data[row]) <= 1 and "" in cp_text_data[row]:
                    # Split at one occurrence.
                    break
                # Append strike_rows to new list.
                text_list_strike.append(cp_text_data[row])

            # This line ensures that there are spaces between all CPs in the new Excel file.
            # Append a blank to new bold list, ensuring spaces between CPs.
            text_list_strike.append(self.space_text)

            # Appends the rest of the main cp text data into the last list.
            for row in range(iteration, len(cp_text_data)):
                text_list_bold.append(cp_text_data[row])

            # This line ensures that there are spaces between all CPs in the new Excel file.
            # Append a blank to new bold list, ensuring spaces between CPs.
            text_list_bold.append(self.space_text)

            iteration = 0
            # For every row in the cp_text_data.
            for row in range(len(cp_font_data)):
                iteration += 1
                # Catch blank spot meant to separate differing CPs.
                if len(cp_font_data[row]) <= 1 and "" in cp_font_data[row]:
                    # Split at one occurrence.
                    break
                # Append strike_rows to new list.
                format_list_strike.append(cp_font_data[row])

            # Adds a blank in the format list to line up with the text lists.
            format_list_strike.append(self.space_format)

            # Appends the rest of the main cp font data into the last list.
            for row in range(iteration, len(cp_font_data)):
                format_list_bold.append(cp_font_data[row])

            # Adds a blank in the format list to line up with the text lists.
            format_list_bold.append(self.space_format)

            # Create cp object.
            cp_object = ExcelCP.ExcelCP(cp_name, cp_coordinates, font_type, cp_station_address, text_list_strike,
                                        format_list_strike, text_list_bold, format_list_bold)
        else:
            print("Occurrences in ExcelPlanner.create_cp_object not set.")
            cp_object = ""

        return cp_object

    # This function helps create multiple CP objects based on coordinates and cp names.
    # Returns a list of all objects for later use.
    def get_all_cp_information(self, list_of_cp_cords, list_of_cp_names):
        # This list will hold all objects created from this function. This is to be returned at the end of function.
        list_of_cp_objects = []

        iteration = 0
        # This will go through each pair of coordinates and find relevant data for each CP.
        # This will produce either one CP record, or two if one is strike-through and one is bold.
        for cords in list_of_cp_cords:
            # Will find if there is only one occurrence of CP. If so, only send through one list of information.
            if len(cords) <= 1:
                # print("One occurrence of CP")
                list_of_cp_objects.append(self.create_cp_object(cords, list_of_cp_names[iteration], 1))
            else:
                list_of_cp_objects.append(self.create_cp_object(cords, list_of_cp_names[iteration], 2))
            iteration += 1
            # Resetting list of rows class variable in ExcelReader - prevents issue where rows keep incrementing for
            # each cp in the list thrown through. It messes up code inside ExcelReader for whatever reason and I can't
            # be bothered to find out why it's happening. This is the quick and dirty solution.
            del self.excel_reader_obj.list_of_rows[:]  # This sets the list to a blank string.

        return list_of_cp_objects

    # This function will "format" all the CP information into the standardized COMPANY_NAME practices.
    # Then it will write to the new excel file.
    def process_and_write_all_cp_information(self, combined_information, new_file):

        # Write to new excel file.
        # ExcelWriter object to write to new excel file.
        writer = ExcelWriter.ExcelWriter(combined_information[0], combined_information[1], new_file)

        writer.write_to_excel()

        return 0

    # This function returns two lists of the combined cp information, text and formatting.
    def return_combined_lists(self, list_of_cp_objects):
        # Main lists to send to ExcelWriter
        strike_text_list = []
        strike_format_list = []
        bold_text_list = []
        bold_format_list = []

        # Put all strike text and bold text together respectively.
        for cp in list_of_cp_objects:
            if cp.get_font_type() == "s" and "N/A" in cp.get_text_list_strike() \
                    and "N/A" in cp.get_format_list_strike():
                bold_text_list += cp.get_text_list_bold()
                bold_format_list += cp.get_format_list_bold()
            elif cp.get_font_type() == "b" and "N/A" not in cp.get_text_list_strike() \
                    and "N/A" not in cp.get_format_list_strike():
                strike_text_list += cp.get_text_list_strike()
                strike_format_list += cp.get_format_list_strike()
                bold_text_list += cp.get_text_list_bold()
                bold_format_list += cp.get_format_list_bold()
            else:
                print("cp font_type variable not set correctly.")

        # Combine text lists and format lists together for writing.
        codeline_info = self.get_codeline_info()
        combined_text = codeline_info[0] + bold_text_list + strike_text_list
        combined_formatting = codeline_info[1] + bold_format_list + strike_format_list

        return [combined_text, combined_formatting]
