from dbfread import DBF                      # For parsing dbf file
import xlsxwriter                            # For writing to xlsx file

PROPERTY_ADDR_COLUMN = 0
PROPERTY_VALUE_COLUMN = 1
OWNER_NAMES_COLUMN = 2
OWNER_ADDRESS_COLUMN = 3
OWNER_TOWN_COLUMN = 4

class MaDataProcessor:
    def __init__(self, options, dbfPath, matchDicPath = None):
        self.options = options
        self.generate_table(dbfPath)
        if matchDicPath:
            self.generate_match_name_set(matchDicPath)

    # Build the database for the town
    def generate_table(self, dbfPath):
        self.table = DBF(dbfPath)

    # Generate the dictionary to match for names in the owner field
    def generate_match_name_set(self, dicPath):
        self.nameSet = set(line.strip().upper() for line in open(dicPath))

    # Initialize the worksheet
    def initialize_worksheet(self):
        wbPath = self.options.town + ".xlsx"
        self.workbook = xlsxwriter.Workbook(wbPath)
        self.worksheet = self.workbook.add_worksheet()
        # Write the header row
        self.currentRow = 0
        self.write_to_row("Property Address", "Total Value", "Owner Names", "Owner Address", "Owner Town")

    # Write entry to worksheet
    def write_to_row(self, propertyAddress, price, owner, ownerAddress, ownerTown):
        if self.worksheet:
            self.worksheet.write(self.currentRow, PROPERTY_ADDR_COLUMN, propertyAddress)
            self.worksheet.write(self.currentRow, PROPERTY_VALUE_COLUMN, price)
            self.worksheet.write(self.currentRow, OWNER_NAMES_COLUMN, owner)
            self.worksheet.write(self.currentRow, OWNER_ADDRESS_COLUMN, ownerAddress)
            self.worksheet.write(self.currentRow, OWNER_TOWN_COLUMN, ownerTown)
            self.currentRow += 1

    # Print out all the record that match the criteria
    def process_match_record(self):
        if self.table:
            self.initialize_worksheet()
            for record in self.table:
                owner = record['OWNER1']
                price = record['TOTAL_VAL']
                buildingValue = record['BLDG_VAL']
                ownerName =owner.replace(',', ' ').split()
                if ( (self.options.includeland or buildingValue > 0) and
                     price > self.options.min and 
                     (self.nameSet == None or any(name in self.nameSet for name in ownerName)) ):
                    propertyAddress = record['SITE_ADDR']
                    ownerAddr = record['OWN_ADDR']
                    ownerCity = record['OWN_CITY']
                    ownerState = record['OWN_STATE']
                    ownerZip = record['OWN_ZIP']
                    ownerTown = "{}, {}, {}".format(ownerCity, ownerState, ownerZip)
                    if self.options.display:
                        print (propertyAddress + "\t\t" + str(price) + "\t\t" + owner + "\t\t" + ownerAddr + ', ' + ownerTown)
                    self.write_to_row(propertyAddress, price, owner, ownerAddr, ownerTown)
            self.workbook.close()
            print("%s.xlsx file generated." %self.options.town)


