import os
import sys
import json
import datetime
import logging
import pandas as pd
from pathlib import Path

import arcpy
from arcgis.gis import GIS

sys.path.insert(0,str(Path(__file__).resolve().parents[2]))

from src.functions import utility, email, arc
from src.classes.servicewrappers import PortalItem, ServiceLayer, TiledService, PortalFile
from src.constants.paths import  PORTAL_URL, LOG_DIR, BACKUPS_DIR
from src.constants.values import VALID_FILE_TYPES
#######################################################################################################################
## Globals
DATETIME_STR = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
LOG_FILE = os.path.join(LOG_DIR, "PortalBackups",f"PortalBackup_{DATETIME_STR}.log")
EXPORT_DIR = os.path.join(BACKUPS_DIR, f"PortalBackup_{DATETIME_STR}")
#######################################################################################################################
## Input Parameters
'''
Here is a quick rundown. The source_type parameter is what type of container is used to retrieve the items. This can either be 'folder', 'group' or 'catalog'.
If using Group or Folder you must use their Item Ids in the source_items list.
The same for the include exclude list. These must be the Feature Services Item IDs.
'''

source_type = 'catalog' ## This can be either 'folder', 'group', or 'catalog'
# The source_utems list has to be the id of either the Folder or Group that will be search source
# source_items = [
#     "Root Folder",
#     "60502ef2599b4d79b06b330710d1f0dc", ## Aquisitions
#     "b9da844e332e449ba1f90e501dec7c4c", ## Administrative
#     "899a1056116d4972a466ca09d454e647",  ## Hazmat
#     "3bbd6f9098e742cf95f7366ecafe79a3", ## Air Quality
#     "21f4f8980eb6424bba13342f9b102bbe", ## Archaeology
#     "da2fa4e32d1e4a98ad57d611d1f88863" ## Backups
# ]
source_items = [

]

item_types = [
    "CSV"
]
# item_types = [
#     "Feature Service",
#     "Service Definition",
#     "CSV",
#     'Microsoft Excel',
#     'Microsoft Word',
#     "Administrative Report",
#     "Shapefile",
#     "File Geodatabase",
#     'Layer Package',
#     'Vector Tile Package',
#     'Tile Package',
#     'Notebook',
#     "Desktop Style",
#     'Map Package',
#     'Project Package'
#     ]
# item_types = [
#     "Tile Package",
#     'Compact Tile Package'
# ]

include_exclude = "all" ## 'include', 'exclude', 'all
# The Include Exclude list has to be the id of the services that will be included or excluded.
include_exclude_list = [
    ]


output_excel = os.path.join(EXPORT_DIR,"PortalBackup_{}.xlsx".format(DATETIME_STR))

#######################################################################################################################
## Email Parameters
email_subject = f"Portal Backup {DATETIME_STR.split('-')[0]}"
email_from = "Edward.smith@hdrinc.com"
email_to= ["Edward.smith@hdrinc.com"]
email_attachments = [output_excel, LOG_FILE]
#######################################################################################################################
## Logging
logging.getLogger("urllib3").setLevel(logging.WARNING)
logger = logging.getLogger("main")
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler(LOG_FILE)
fh.setLevel(logging.DEBUG)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)
#######################################################################################################################



def main():

    '''
    Logic to determine whether the script is being run by the Task Scheduler or manually.
    If the script is being run manually the user will need to provide login credentials through the terminal
    If the script is being run on the scheduler than the login credentials are passed in as parameters.
    '''
    if utility.isTaskScheduler():
        username = sys.argv[1]
        scheduled = True
        if username.lower() == "pro":
            gis_connection = GIS("Pro")

        else:
            password = sys.argv[2]
            gis_connection = GIS(PORTAL_URL, username=username, password=password)

    else:
        scheduled = False
        gis_connection = arc.authenticateAgolConnection(PORTAL_URL)


    '''
    Here is the a block of statements that are adding input parameters to the log file.
    '''

    ## Logging Input Parameters
    logger.info(f"GIS Connection: {gis_connection}")
    logger.info(f"Scheduled Task: {scheduled}")
    logger.info(f"Searched Source Type: {source_type}")
    logger.info(f"Searched Source Item IDs: {source_items}")
    logger.info(f"Excel Report: {output_excel}")
    logger.info(f"Include/Exclude Flag: {include_exclude}")
    logger.info(f"Service List: {include_exclude_list}")
    logger.info(f"Email From: {email_from}")
    logger.info(f"Email To: {email_to}")
    logger.info("~~"*100)
    logger.info("~~"*100)


    '''
    Here is where we are creating a list of ArcGIS Item Objects that will be used in the "generateAppendixReport" Function
    We are using folders to determine what items to look for with a parameter specifying if there is a list of feature services that should not be included.

    '''
    ## This will be a dictionary that holds dictionaries with the key being the Item Type and the value being a list of attribute dictionaries
    ## This will be used to generate the Excel Report

    df_dict = {
        "Failed":[],
        "Parameters": {
            "Datetime": DATETIME_STR,
            "Log File": LOG_FILE,
            "GIS Connection": str(gis_connection),
            "GIS Username":gis_connection.users.me.username,
            "Local Username":os.getlogin(),
            "Source Type": source_type,
            "Source Items": ", ".join(source_items) if source_items else "None",
            "Include/Exclude Flag":include_exclude,
            "Include/Exclude List": ", ".join(include_exclude_list) if include_exclude_list else "None",
            "Excel Path":output_excel,
            "Export Directory": EXPORT_DIR,
            "Email From": email_from,
            "Email To": ", ".join(email_to) if email_to else "None"
    }}

    utility.create_directory(EXPORT_DIR)

    if "Feature Service" in item_types:
        gdb_md_dict = {
            "description": json.dumps(df_dict["Parameters"], indent=1),
            "summary": "Portal Data Backup",
            "tags": ["backup"],
            "credits":"",
            "accessConstraints":""
        }

        export_gdb = arc.create_fgdb(os.path.join(EXPORT_DIR, "FeatureService"), f"FeatureServiceBackup_{DATETIME_STR}.gdb", gdb_md_dict)

    logger.info(f"Generating Portal Item List...")
    
    item_list = arc.generateItemList(gis_conn=gis_connection,
                                        source_type=source_type,
                                        item_types=item_types,
                                        include_exclude_flag=include_exclude,
                                        source_list=source_items,
                                        include_exclude_list=include_exclude_list
                                        )
    logger.info(f"Item Count: {len(item_list)}")
    logger.info("Starting the Item Iteration...")
    for item in item_list:
        if item.type not in df_dict:
            logger.debug(f"Adding {item.type} Key to df_dict.")
            df_dict[item.type] = []
        logger.info(f"Item Object: {item}")
        logger.info(f"Item Type: {item.type}")
        if item.type == "Feature Service":
            if arc.checkLayerAccessibility(item):
                logger.info(f"Layer Count: {len(item.layers)}")
                for layer in item.layers:
                    logger.info(f"Layer Object: {layer}")
                    sl = ServiceLayer(gis_connection, item, layer)

                    logger.info(f"Exporting Layer {sl.layerName}...")
                    exported_fc, failed_dict = sl.exportLayer(export_gdb)
                    
                    if failed_dict:
                        df_dict["Failed"].append(failed_dict)

                    logger.info(f"Updating Local Metadata...")
                    md_dict = sl.getLayerMetadataDictionary()

                    local_md = arcpy.metadata.Metadata(exported_fc)
                    ## Updates the newly exported feature classes metadata to the the same as the Portal Items.
                    [setattr(local_md, k, v) for k,v in md_dict.items()]
                    local_md.save()

                    excel_dict = sl.getLayerExcelDictionary()
                    excel_dict["Feature Class Path"] = exported_fc

                    df_dict[item.type].append(excel_dict)
            else:
                logger.error(f"Item {item.id} Layers are not accessible.")
                continue

        elif item.type == "Map Service" or item.type == "Vector Tile Service":
            directory_path = utility.create_directory(os.path.join(EXPORT_DIR, item.type.replace(" ", "")))
            ts = TiledService(gis_connection, item)
            ts.exportTiles()

        elif item.type in VALID_FILE_TYPES:
            directory_path = utility.create_directory(os.path.join(EXPORT_DIR, item.type.replace(" ", "")))
            pf = PortalFile(gis_connection, item)
            
            out_file, failed_dict = pf.downloadFile(directory_path)

            if failed_dict:
                df_dict["Failed"].append(failed_dict)

            excel_dict = pf.getFileExcelDictionary()
            excel_dict["File Path"] = 'outputs\{}'.format(out_file.split("\outputs\\")[1])
            df_dict[item.type].append(excel_dict)

    logger.info(f"Exporting Excel Reports...")
    logger.debug(df_dict)
    logger.debug(f"DF Dictionary Keys: {df_dict.keys()}")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for item_type, property_list in df_dict.items():
            logger.info(f"Item Type: {item_type}")
            logger.info(f"Property List: {property_list}")
            try:
                logger.info(f"Generating DataFrame...")
                if type(property_list) is dict:
                    df = pd.DataFrame.from_dict(property_list, "index", columns=["Value"])
                    df.index.name = "Parameter"
                    idx = True
                    df.to_excel(writer, sheet_name=item_type, index=True)

                else:
                    df = pd.DataFrame(property_list)
                    idx = False

                df.to_excel(writer, sheet_name=item_type, index=idx)
                logger.info(df.head())

            except Exception as e:
                logger.error(f"Failed to Generate DataFrame!!\n{e}")
            

    if email_from:
        email.sendEmail(email_to, email_from, email_subject, "Portal Backup Complete.", "plain", attachments=email_attachments)
            


if __name__ == "__main__":
    main()

