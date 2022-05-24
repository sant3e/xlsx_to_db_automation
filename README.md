# Automation script for uploading (SharePoint connected) Excel Files to a Postgre Database

Created an automation script that: 
  1. Opens local storred Excel Files 
  2. Refreshes (user-created) connection to either a SharePoint Folder or a SharepointList 
  3. Cleans the file names, column headers and sets proper data types (compliant with SQL standards)
  4. Creates the necessary (pandas) dataframes
  5. Uploads to a PostgreSQL DB using psycopg2


-------------------------
### What is a SharePoint connected Excel File?
If you have Excel files in a Sharepoint Folder or you want to pull that data from a Sharepoint List, dump them into excel files then upload to a Database:

My solution is to create a connection using Excel/Data/Get Data/ 
                                                                - either: From File/From Sharepoint File (if you have an excel)
                                                                - either: From Online Services/From Sharepoint Online list (if you have a sharepoint list)

One could just integrate another python script to pull data from Sharepoint Lists into Excel, but in my case, I'm using a corporate Sharepoint which has MultiFactorAuthentication and could not find any (working) library that deals with that. So i simply just bypassed that and create connections within (local) Excel Files

I chose excel files and not csv, due to "preservation" of text. 
While trying this with CSV, my text got scrambled (even though i used a comprehensive encoding). 
Excel files preserved everything as from the original data

-----------------------------
### Code and Resources Used
**Python Version:** 3.10
**Packages:** os, numpy, pandas, psycopg2, shutil, openpyxl, pywin32

**Source Github:** https://github.com/Strata-Scratch/csv_to_db_automation
* My gratitude to StrataScratch for wonderful tutorials and a great youtube channel:
https://www.youtube.com/channel/UCW8Ews7tdKKkBT6GdtQaXvQ
https://platform.stratascratch.com/coding

-------------------------
### Productionization
In this step, I built an interface using TKinter that allows the user to push a button to start the script. This part is not integrated here (one might not need this), but I can provide code (on request)
