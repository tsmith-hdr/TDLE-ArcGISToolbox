import os
import sys
import datetime 
import logging
from pathlib import Path
from importlib import reload

import arcpy
from arcgis.gis import GIS

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))

from src.functions import utility, email, appendix_report, arc
from src.classes.services import PortalItem, ServiceLayer, TiledService, PortalFile
from src.tools.backupmanagement import TOOL_AppendixReport
from src.constants.paths import  PORTAL_URL, INTRANET_APPENDIX_H_DIR, LOG_DIR, OUTPUTS_DIR, BACKUPS_DIR
#######################################################################################################################
## Globals
DATETIME_STR = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
LOG_FILE = os.path.join(LOG_DIR, "Scheduled", "AppendixReports",f"AppendixH_{DATETIME_STR}_Scheduled.log")
EXPORT_DIR = os.path.join(BACKUPS_DIR, f"PortalBackup_{DATETIME_STR}")
#######################################################################################################################
## Input Parameters 
'''
Here is a quick rundown. The source_type parameter is what type of container is used to retrieve the items. This can either be 'folder', 'group' or 'catalog'. 
If using Group or Folder you must use their Item Ids in the source_items list. 
The same for the include exclude list. These must be the Feature Services Item IDs.
'''

source_type = 'folder' ## This can be either 'folder', 'group', or 'catalog'
# The source_utems list has to be the id of either the Folder or Group that will be search source
source_items = [
    "82b74a9180f64fb8bc62b0188c368734", ## Measures
    "eb8c5a2fb2324889a289e7b51997e1ac"  ## Alternatives
]

item_types = [

]

include_exclude = "include" ## 'include', 'exclude', 'all
# The Include Exclude list has to be the id of the services that will be included or excluded.
include_exclude_list = [
    "51b54afe34354a82925463f1fa6f3889",  ## SAFER Mitigation Measures (HDR 2025)
    "f0cf92ae637a475bb8d86ae01e1b0e1e" ## Tunnels ROW (HDR 2025)
    ]

include_records = False ## 

output_excel = os.path.join(EXPORT_DIR,"PortalBackup_{}.xlsx".format(DATETIME_STR))

#######################################################################################################################
## Email Parameters
email_subject = f"Portal Backup {DATETIME_STR.split('-')[0]}"
email_from = "Edward.smith@hdrinc.com"
email_to= ["Edward.smith@hdrinc.com"]
#######################################################################################################################
## Logging
logger = logging.getLogger()
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

if __name__ == "__main__":
    
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
        gis_connection = utility.authenticateAgolConnection(PORTAL_URL)


    '''
    Here is the a block of statements that are adding input parameters to the log file.
    '''

    ## Logging Input Parameters
    logger.info(f"GIS Connection: {gis_connection}")
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

    utility.create_directory(EXPORT_DIR)

    if "Feature Service" in item_types:
        FS_DIR = utility.create_directory(os.path.join(EXPORT_DIR, "FeatureService"))
        logger.info(f"Creating Local File GDB...")
        fs_gdb_path = os.path.join(FS_DIR, f"FeatureServiceBackup_{DATETIME_STR}.gdb")
        logger.info(f"Backup GDB Path: {fs_gdb_path}")
        try:
            arcpy.management.CreateFileGDB(out_folder_path=FS_DIR, 
                                            out_name=f"FeatureServiceBackup_{DATETIME_STR}")
        except Exception as t:
            logger.error(f"Failed to Create Feature Service GDB.\n{t}")
            raise ValueError(f"Failed to Create Feature Service GDB.\n{t}")

    logger.info(f"Generating Portal Item List")
    item_list = arc.generateItemList(gis_conn=gis_connection,
                                     source_type=source_type,
                                     item_types=item_types,
                                     include_exclude_flag=include_exclude,
                                     source_list=source_items,
                                     include_exclude_list=include_exclude_list
                                     )
    
    for item in item_list:
        if item.type == "Feature Service":
            for layer in item.layers:
                sl = ServiceLayer(item, layer)

                logger.info(f"Exporting Layer: {sl.layerName}")
                exported_fc = sl.exportLayer(fs_gdb_path)

                logger.info(f"Updating Local Metadata")
                md_dict = sl.getMetadataDictionary()

                local_md = arcpy.metadata.Metadata(exported_fc)
                ## Updates the newly exported feature classes metadata to the the same as the Portal Items.
                [setattr(local_md, k, v) for k,v in md_dict.items()]
                local_md.save()

        elif item.type == "Map Service" or item.type == "Vector Tile Service":
            ts = TiledService(item)
            ts.exportTiles()

        elif item.type in 
            directory_path = utility.create_directory(os.path.join(EXPORT_DIR, item.type.replace(" ", "")))
            pf = PortalFile(item, directory_path)
            pf.downloadFile()

