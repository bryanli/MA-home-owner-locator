from collections import OrderedDict
from dbfread import DBF                      # For parsing dbf file
import xlsxwriter                            # For writing to xlsx file
import xlrd                                  # For reading xlsx file
from argparse import ArgumentParser          # For parsing command line arguments
from argparse import RawTextHelpFormatter
import wget                                  # For downloading database file from MA data source
import os                                    # For removing files
import shutil                                # For removing entire non empty directory
import zipfile                               # For unziping the downloaded database file 
import fnmatch                               # For finding the Assess.dbf file

# Constant values for the program
DATABASE_LINK_FILE_DIRECTORY = "./DB_Links"
DATABASE_LINK_FILE = DATABASE_LINK_FILE_DIRECTORY + "/MassGIS_Parcel_Download_Links.xlsx"
DEFAULT_MIN_VALUE = 100000
TMP_PATH = "./tmp"
DBF_FILE_NAME_FORMAT = "*Assess.dbf"

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

# Build the database for the town
def generate_table(dbfPath):
    table = DBF(dbfPath)
    return table

# Generate the dictionary to match for names in the owner field
def generate_match_name_set(dicPath):
    nameSet = set(line.strip().upper() for line in open('./LastNameMatchDictionary/PinYinWordList.txt'))
    return nameSet

# Download the town's data from MA data source, extract and return the path to the Assess.dbf file
def download_and_unzip_data(townName, url):
    filePath = '%s/%s.zip' %(TMP_PATH, townName)
    unzipPath = '%s/%s' %(TMP_PATH, townName)
    # Create the tmp directory if not exist
    if not os.path.exists(TMP_PATH):
        os.makedirs(TMP_PATH)
    # Remove the downloaded file if existed
    if os.path.exists(filePath):
        os.remove(filePath)
    # Remove the unzipped directory if existed
    if os.path.exists(unzipPath):
        try:
            shutil.rmtree(unzipPath)
        except OSError as e:
            print("Error: %s : %s" % (unzipPath, e.strerror))
    
    # Download and unzip the town database
    wget.download(url, filePath)
    print("\n" + filePath + " download completed.")
    with zipfile.ZipFile(filePath, 'r') as zip_ref:
        zip_ref.extractall(unzipPath)
    print(unzipPath + " unzip completed.")

    # Find the assess.dbf file in the database
    filename = None
    for root, dirnames, filenames in os.walk(unzipPath):
        for f in filenames:
            if fnmatch.fnmatch(f, DBF_FILE_NAME_FORMAT):
                filename = os.path.join(root, f)
    return filename


def cleanup_tmp():
    if os.path.exists(TMP_PATH):
        try:
            shutil.rmtree(TMP_PATH)
        except OSError as e:
            print("Error: %s : %s" % (TMP_PATH, e.strerror))


# Print out all the record that match the criteria
def print_match_record(table, matchSet = None):
    if table:
        for record in table:
            owner = record['OWNER1']
            price = record['TOTAL_VAL']
            buildingValue = record['BLDG_VAL']
            ownerName =owner.replace(',', ' ').split()
            if ( (options.includeland or buildingValue > 0) and
                 price > options.min and 
                 (matchSet == None or any(name in matchSet for name in ownerName)) ):
                propertyAddress = record['SITE_ADDR']
                ownerAddr = record['OWN_ADDR']
                ownerCity = record['OWN_CITY']
                ownerState = record['OWN_STATE']
                ownerZip = record['OWN_ZIP']
                ownerInfoStr = "{}, {}, {}, {}".format(ownerAddr, ownerCity, ownerState, ownerZip)
                print (propertyAddress + "\t\t" + str(price) + "\t\t" + owner + "\t\t" + ownerInfoStr)


def main():
    parse_options()
    links = generate_db_links()

    if not options.town or options.town not in links:
        prompt_for_town(links)

    print(links[options.town])
    assessFilePath = download_and_unzip_data(options.town, links[options.town][1])

    if assessFilePath:
        table = generate_table(assessFilePath)
        nameSet = generate_match_name_set('./LastNameMatchDictionary/PinYinWordList.txt')
        print_match_record(table, nameSet)
    else :
        print ("Failed to generate data fpr %s" %options.town)

    # remove the tmp directory
    cleanup_tmp()
  


if __name__== "__main__":
    main()

