# CODE INTENDED TO RUN ON WINDOWS OS
# Install required libraries
# This code downloads the required files from BSEE website and checks for any updates. If there are any updates,
# it refreshes the existing data

# ----------------------------------------------------------------------------------------------------------------------

# Imports
from winreg import *
from bs4 import BeautifulSoup
import requests
import re  # regular expression
import webbrowser
import os
import shutil
import time
import pandas as pd
from datetime import datetime

# Get the default Downloads directory of the computer
with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    pc_downloads_directory = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

pc_downloads_directory = pc_downloads_directory.replace('\\', '/')  # Replace to make the path consistent with the code

# Get the current working directory
cwd = os.getcwd()

# Years the data is required for
years = [2019, 2020]  # <========================= MODIFY THIS TO GET DATA FOR OTHER YEARS /////////////////////////////

# Create an empty dataframe to be used later for appending each year's data
df = pd.DataFrame()

# URL of the website from where the files have to be scraped
url = "https://www.data.bsee.gov/Main/OGOR-A.aspx"

# URL of the home page of BSEE website
url_homepage = "https://www.data.bsee.gov"

# List of column names of the production file
col_names = ['Lease Number', 'Completion Name', 'Production Date', 'Days On Prod', 'Product Code',
                 'Monthly Oil Volume', 'Monthly Gas Volume', 'Monthly Water Volume', 'Api Well Number',
                 'Well Status Code', 'Area/Block/Bottom Area/Bottom Block', 'Operator Num', 'Operator Name',
                 'Field Name Code', 'Injection Volume', 'Production Interval Code', 'First Production Date',
                 'Unit Agreement Number', 'Unit Aloc Suffix']

# Get HTML code of the URL
page = requests.get(url).text
doc = BeautifulSoup(page, "html.parser")

# -------------------------------------------- SCRAPING PART BEGINS ----------------------------------------------------

if 'Download' not in os.listdir(cwd):

    # Loop through the years to scrape the required data
    for year in years:

        # Find parts of the HTML where instances of the required year are present
        # Find the parent's parent's parent block which contains the required info of 'Last Updated' and the 'Download Link'
        # The 2nd TD tag contains the 'Last Updated' value. Split twice the final string to get the date
        # Replace "/" by "-" just for convenience
        last_updated = doc.find_all(text=re.compile(f"{year}"))[0].parent.parent.parent.find_all("td")[1].text.split(' ')[0].replace("/", "-")

        # The 4th TD tag contains the 'Download Link'. The link is embedded in the HREF part
        link = doc.find_all(text=re.compile(f"{year}"))[0].parent.parent.parent.find_all("td")[3].a['href']

        # The download link is the combination of home page URL and the link scraped
        link_to_file = url_homepage + link

        # Opening the link downloads the required .zip file
        webbrowser.open(link_to_file)

        # Time delay to make sure the file is downloaded before unzipping
        time.sleep(6)

        # Unzip the .zip file to the 'Download' sub-folder of the working directory where the python script is executed
        shutil.unpack_archive(f"{pc_downloads_directory}/ogora{year}delimit.zip", "Download")

        # Time delay to make sure the unzipping process is done before deleting the .zip file
        time.sleep(2)

        # Delete the .zip file now that the contents are unzipped
        os.remove(f"{pc_downloads_directory}/ogora{year}delimit.zip")

        # Rename the unzipped file to contain the year for which the data is scrapped and the 'Last Updated' date
        os.rename(f"{cwd}\\Download\\ogora{year}delimit.txt", f"{cwd}\\Download\\ogora{year}_{last_updated}.txt")

        # Read the unzipped .txt file and convert it into .csv
        read_file = pd.read_csv(fr'{cwd}\\Download\\ogora{year}_{last_updated}.txt')
        read_file.to_csv(fr'{cwd}\\Download\\ogora{year}_{last_updated}.csv', index=None)

        # Convert the .csv file to a temporary dataframe
        df_temp = pd.read_csv(fr'{cwd}\\Download\\ogora{year}_{last_updated}.csv', names=col_names, low_memory=False)

        # Keep the required columns only in the dataframe
        df_temp = df_temp[
            ['Lease Number', 'Production Date', 'Product Code', 'Monthly Oil Volume', 'Monthly Gas Volume', 'Operator Num']]

        # Concat the temporary dataframe to the dataframe df (initialized before the loop)
        # Hence, the df will contain data for all the required years
        df = pd.concat([df, df_temp])

        # Remove the temporary .csv file
        os.remove(fr'{cwd}\\Download\\ogora{year}_{last_updated}.csv')

    # --------------------------------------------- LOOP ENDS HERE -----------------------------------------------------

    # Data cleaning
    # Let's drop the entries with Null values.
    # The entries will Null values are negligible compared to the total number of rows, hence it is safe to drop these
    # Used Jupyter notebook to get the number of rows with Null entries
    df = df.dropna()

    # No more cleaning is required. Convert the df to the final Production.csv file
    df.to_csv(fr'{cwd}\\Download\\production.csv', index=None)

# ------------------------ CHECK IF UPDATED DATA IS PRESENT AND REFRESH THE Production.csv data ------------------------

else:

    # Loop through the required years to check if the data in the website was updated
    for year in years:

        # Get the last updated date
        last_updated = doc.find_all(text=re.compile(f"{year}"))[0].parent.parent.parent.find_all("td")[1].text.split(' ')[0].replace("/", "-")

        # Get a list of available .txt files, the data for which we have currently
        list_files = os.listdir(f"{cwd}\\Download")

        # Compare the last updated date (from file name) and last_updated date scraped from website for the required year's production file
        for file in list_files:

            # Skip the Production.csv file
            if file != "production.csv":

                # Compare to make sure we are checking the correct year's production file
                if int(file[5:9]) == year:

                    # Compare both the dates to check if any new updated data is present
                    if datetime.strptime(last_updated, '%m-%d-%Y') > datetime.strptime(file.split('_')[-1].split('.')[0], '%m-%d-%Y'):

                        # Initialize an empty dataframe so that the data does not concat with the existing data
                        df = pd.DataFrame()

                        # Get the date of last updated fetched from the website
                        last_updated = str(last_updated)[0:10]

                        # Remove the old file
                        os.remove(f"{cwd}\\Download\\ogora{year}_{file.split('_')[-1]}")

                        # Get the link to the .txt production file for the required year
                        link = doc.find_all(text=re.compile(f"{year}"))[0].parent.parent.parent.find_all("td")[3].a['href']
                        link_to_file = url_homepage + link

                        # Open the link and download the file
                        webbrowser.open(link_to_file)

                        # Time delay to make sure the file is downloaded before unzipping
                        time.sleep(6)

                        # Unzip the .zip file to the 'Download' sub-folder of the working directory where the python script is executed
                        shutil.unpack_archive(f"{pc_downloads_directory}/ogora{year}delimit.zip", "Download")

                        # Time delay to make sure the unzipping process is done before deleting the .zip file
                        time.sleep(2)

                        # Delete the .zip file now that the contents are unzipped
                        os.remove(f"{pc_downloads_directory}/ogora{year}delimit.zip")

                        # Time delay to make sure the unzipping process is done before deleting the .zip file
                        time.sleep(2)

                        # Rename the unzipped file to contain the year for which the data is scrapped and the 'Last Updated' date
                        os.rename(f"{cwd}\\Download\\ogora{year}delimit.txt",
                                  f"{cwd}\\Download\\ogora{year}_{last_updated}.txt")

                        # Delete the already existing Production.csv, if it exists
                        if 'production.csv' in list_files:
                            os.remove(f"{cwd}\\Download\\production.csv")

                        # Create a list of files present in the 'Download' sub-folder currently after the update
                        list_files = os.listdir(f"{cwd}\\Download")

                        # For each file in the 'Download' sub-folder currently
                        for new_file in list_files:

                            # Read the unzipped .txt file and convert it into .csv
                            read_file = pd.read_csv(fr'{cwd}\Download\{new_file}')
                            read_file.to_csv(fr'{cwd}\\Download\\ogora{year}_{last_updated}.csv', index=None)

                            df_temp = pd.read_csv(fr'{cwd}\Download\ogora{year}_{last_updated}.csv', names=col_names,
                                                  low_memory=False)

                            # Keep the required columns only in the dataframe
                            df_temp = df_temp[
                                ['Lease Number', 'Production Date', 'Product Code', 'Monthly Oil Volume',
                                 'Monthly Gas Volume', 'Operator Num']]

                            # Concat the temporary dataframe to the dataframe df (initialized above as an empty df)
                            # Hence, the df will contain data for all the required years
                            df = pd.concat([df, df_temp])

                            # Remove the temporary .csv file
                            os.remove(fr'{cwd}\Download\ogora{year}_{last_updated}.csv')

    # # Data cleaning
    # # Let's drop the entries with Null values.
    # # The entries will Null values are negligible compared to the total number of rows, hence it is safe to drop these
    # # Used Jupyter notebook to get the number of rows with Null entries
    df = df.dropna()

    # No more cleaning is required. Convert the df to the final Production.csv file
    df.to_csv(fr'{cwd}\\Download\\production.csv', index=None)

# ------------------------------------------ UPDATE PROCESS ENDS HERE --------------------------------------------------
