# Written and Developed by: Tyler Wright
# Date started: 03/19/2019
# Date when workable: 03/21/2019
# Last Updated: 03/21/2019


"""

This class file is meant to hold CP objects for Excel purposes.
Mainly, this will be used for COMPANY_NAME documentation sent in CFA format.

"""

"""

    Programmer Notes:
    
    What does a CP Object need?
    
    1. A CP Name
    2. Coordinates of where the CP is in an excel file.
    3. A station address number.
    4. A list of the text relative to its station address (strike-through).
    4.1 A list of the formatting information for its text (strike-through).
    5. A list of text relative to its station address (bold).
    5.1 A list of the formatting information for its text (bold).
    

"""


class ExcelCP:

    def __init__(self, name, coordinates, font_type, station, text_list_strike, format_list_strike, text_list_bold,
                 format_list_bold):
        self.name = name
        self.coordinates = coordinates
        self.font_type = font_type
        self.station = station
        self.text_list_strike = text_list_strike
        self.format_list_strike = format_list_strike
        self.text_list_bold = text_list_bold
        self.format_list_bold = format_list_bold

    """********************************************** Getters ******************************************************"""

    def get_name(self):
        return self.name

    def get_coordinates(self):
        return self.coordinates

    def get_font_type(self):
        return self.font_type

    def get_station(self):
        return self.station

    def get_text_list_strike(self):
        return self.text_list_strike

    def get_format_list_strike(self):
        return self.format_list_strike

    def get_text_list_bold(self):
        return self.text_list_bold

    def get_format_list_bold(self):
        return self.format_list_bold
