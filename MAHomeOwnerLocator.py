from collections import OrderedDict
import xlsxwriter                            # For writing to xlsx file
import xlrd                                  # For reading xlsx file
from argparse import ArgumentParser          # For parsing command line arguments
from argparse import RawTextHelpFormatter
from MaDataGather import download_and_unzip_data, cleanup_tmp
from MaDataProcessor import MaDataProcessor

# Constant values for the program
MATCH_DICTIONARY_DIRECTORY = "./LastNameMatchDictionary"
DATABASE_LINK_FILE_DIRECTORY = "./DB_Links"
DATABASE_LINK_FILE = DATABASE_LINK_FILE_DIRECTORY + "/MassGIS_Parcel_Download_Links.xlsx"
DEFAULT_MIN_VALUE = 100000

# Parse the input options and setup the program
def parse_options():
    """ Leverage the argparse module to pass in command line arguments. A brief
    description of each supported parameter can be viewed by using the -h option.
    """
    global options
    parser = ArgumentParser(formatter_class=RawTextHelpFormatter)
    parser.add_argument('--min', type=int, help='The minimum property price to show', default=DEFAULT_MIN_VALUE)
    parser.add_argument('--includeland', default=False, help='Only show property that has structure value')
    parser.add_argument('--town',type=str)
    parser.add_argument('--debug', default=False)
    options = parser.parse_args()
    options.debug = options.debug in ['True', 'true', True]
    options.includeland = options.includeland in ['True', 'true', True]
    if options.town:
        options.town = options.town.upper()

# If town is not set in the command line, prompt the user for input
def prompt_for_town(links):
    for key in links.keys():
        print (key)
    while not options.town or options.town not in links:
        if options.town:
            print("Cannot find " + options.town + " in Massachusetts");
        options.town = raw_input("Enter town name from the list above : ").upper()

# Read all the database for towns from the xlsx file
def generate_db_links(dbLinkPath = DATABASE_LINK_FILE):
    wb = xlrd.open_workbook(dbLinkPath)
    sheet = wb.sheet_by_index(0)
    linksDb = OrderedDict()
    for i in range(sheet.nrows):
        if i != 0:
            linksDb[sheet.cell_value(i, 0)] = [int(sheet.cell_value(i, 1)), sheet.cell_value(i, 2)]
    return linksDb


def main():
    parse_options()
    links = generate_db_links()

    if not options.town or options.town not in links:
        prompt_for_town(links)

    print(links[options.town])
    assessFilePath = download_and_unzip_data(options.town, links[options.town][1])

    if assessFilePath:
        matchDictionaryPath = MATCH_DICTIONARY_DIRECTORY + "/PinYinWordList.txt"
        processor = MaDataProcessor(options, assessFilePath, matchDictionaryPath)
        processor.print_match_record()
    else :
        print ("Failed to generate data fpr %s" %options.town)

    # remove the tmp directory
    cleanup_tmp()
  


if __name__== "__main__":
    main()

