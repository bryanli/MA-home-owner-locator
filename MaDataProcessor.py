from dbfread import DBF                      # For parsing dbf file

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


    # Print out all the record that match the criteria
    def print_match_record(self):
        if self.table:
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
                    ownerInfoStr = "{}, {}, {}, {}".format(ownerAddr, ownerCity, ownerState, ownerZip)
                    print (propertyAddress + "\t\t" + str(price) + "\t\t" + owner + "\t\t" + ownerInfoStr)


