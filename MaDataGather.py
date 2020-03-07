import wget                                  # For downloading database file from MA data source
import os                                    # For removing files
import shutil                                # For removing entire non empty directory
import zipfile                               # For unziping the downloaded database file 
import fnmatch                               # For finding the Assess.dbf file

# Constant values for the module
TMP_PATH = "./tmp"
DBF_FILE_NAME_FORMAT = "*Assess.dbf"

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


