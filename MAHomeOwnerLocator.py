from dbfread import DBF
import xlsxwriter 

FILTER_TOTAL_VALUE = 100000
FILTER_BUILDING_VALUE = 0

PinYinSet = set(line.strip().upper() for line in open('./LastNameMatchDictionary/PinYinWordList.txt'))

def generate_table(dbfPath):
    table = DBF(dbfPath)
    return table

def print_match_record(table):
    if table:
        for record in table:
            owner = record['OWNER1']
            price = record['TOTAL_VAL']
            buildingValue = record['BLDG_VAL']
            ownerName =owner.replace(',', ' ').split()
            if buildingValue > FILTER_BUILDING_VALUE and price > FILTER_TOTAL_VALUE and any(name in PinYinSet for name in ownerName):
                propertyAddress = record['SITE_ADDR']
                ownerAddr = record['OWN_ADDR']
                ownerCity = record['OWN_CITY']
                ownerState = record['OWN_STATE']
                ownerZip = record['OWN_ZIP']
                ownerInfoStr = "{}, {}, {}, {}".format(ownerAddr, ownerCity, ownerState, ownerZip)
                print (propertyAddress + "\t\t" + str(price) + "\t\t" + owner + "\t\t" + ownerInfoStr)
       
def main():
    table = generate_table("Southborough.dbf")
    print_match_record(table)
  
if __name__== "__main__":
    main()

