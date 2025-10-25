#######################################################################################################################################################
## Logging
import logging
logger = logging.getLogger(f"main.arc")
#######################################################################################################################################################

import sys
import os
from pathlib import Path
from datetime import datetime
import getpass
import arcpy
from arcgis.gis import GIS, ItemTypeEnum

sys.path.insert(0,str(Path(__file__).resolve().parents[2]))

from src.constants.values import *
from src.functions import utility
################################################################################################################################################################

def authenticateAgolConnection(portal_url):
    """
    Allows the user of the standalone script enter their user credentials. This is to avoid having to store credentials.
    The portal url is set in the
    """
    print(f"-- If using Arcgis Pro as Authentication Credentials. Input 'Pro' for username and input nothing for password and Press 'Enter' to continue.\n** You will need to be sure you are logged in to your account and correct Portal in ArcGIS Pro")
    print(f"-- Please enter 'Pro' or your Username and Password for {portal_url} --")
    count = 0
    while True:
        if count > 2:
            print(f"Too Many Attempts !! ")
            input(f"Press Any Key to Exit...")
            sys.exit("Exiting Script...")


        username = input("Username: ")
        password = getpass.getpass()
        print(f"Authenticating...")


        try:
            if username.lower().strip() == "pro":
                gis_conn = GIS("Pro")
            else:
                gis_conn = GIS(portal_url, username, password)
            if gis_conn:
                break
        except Exception as e:
            count+=1
            print(f"Failed GIS Connection: {e} Please Re-enter Credentials...")



    return gis_conn


def checkSource(gis_conn:GIS, source_type:str, source_item:str)->bool:
    if source_type.lower() == "folder":
        return True if gis_conn.content.folders.get(source_item) else False

    elif source_type.lower() == "group":
        return True if gis_conn.groups.get(source_item) else False

    elif source_type.lower() not in ["folder","group", "catalog"]:
        logger.error(f"Incorrect 'source' variable")
        raise ValueError(f"Incorrect 'source' variable. Must be 'folder', 'group', or 'content'")


def generateItemList(gis_conn:GIS, source_type:str, item_types:list, include_exclude_flag:str="all",source_list:list=None, include_exclude_list:list=None)->list:
    '''
    Input Parameters
    - gis_conn : GIS Connection Object
    - source_type : This is the place were the items are searched for. The options are;
    ~~ "folder" : Specific Folders *These folders can only be folder that account being run
    ~~ "group" : Named Portal Groups
    ~~ "catalog" : The entire content library
    - item_types : This will be an input list of the all the item types that should be included in the portal backup.
    - include_exclude_flag : Specify if the items specified  should be included or excluded
    ~~ "all" (Default) : Using this flag, all of the items within the group or portal will be included in the Item List
    ~~ "include" : these item id's will be the only services included
    ~~ "exclude" : these item id's will not be included in the backup
    '''
    item_id_list = [] ## This list will be used to avoid duplicates when choosing multiple groups as the source.
    item_list = []

    search_method = {
        "folder":gis_conn.content.folders.get,
        "group": gis_conn.groups.get,
        'catalog': gis_conn.content.search
    }



    if source_type in ["folder", "group"]:
        for source in source_list:
            logger.info(f"Source: {source}")
            if checkSource(gis_conn, source_type, source) and source_type in ['folder', 'group']:
                content_list = search_method[source_type.lower()](source).list() if source_type.lower() == 'folder' else [c for c in search_method[source_type.lower()](source).content(max_items=5000)]

                if include_exclude_flag.lower() == "include":
                    [item_list.append(i) for i in content_list if i.id in include_exclude_list and i.id not in item_id_list and i.type in item_types]
                elif include_exclude_flag.lower() == "exclude":
                    [item_list.append(i) for i in content_list if i.id not in include_exclude_list and i.id not in item_id_list and i.type in item_types]
                else:
                    [item_list.append(i) for i in content_list if i.id not in item_id_list and i.type in item_types]
    else:
        [item_list.append(i) for i in search_method["catalog"]("", max_items=-1) if i.type in item_types]



    return item_list



def create_fgdb(directory_path:Path, gdb_name:str, metadata_dict:dict=None)->Path:
    logger.info(f"Creating File Geodatabase...")
    logger.info(f"Verifying Directory...")

    dir_path = utility.create_directory(directory_path)

    logger.info(f"Creating Local File GDB...")

    gdb_path = os.path.join(dir_path, gdb_name)#

    logger.info(f"Backup GDB Path: {gdb_path}")

    try:
        arcpy.management.CreateFileGDB(out_folder_path=dir_path, out_name=gdb_name)
        
    except Exception as t:
        logger.error(f"Failed to Create FGDB.\n{t}")
        raise ValueError(f"Failed to Create FGDB.\n{t}")
    
    try:
        if metadata_dict:
            md = arcpy.metadata.Metadata(gdb_path)
            for k,v in metadata_dict.items():
                setattr(md, k, v)
            md.save()
    except Exception as e:
        logger.warning(f"File Geodatabase Metadata was not updated!\n{e}")
    
    return gdb_path


def compressFgdbItems(gdb_path:Path)->list:
    failed_list = []
    ## Here we are compressing the file gdb this is a lossl_objess function. We want to add this process to make sure that the archived records are unable to be editied.
    logger.info(f"Compressing File Geodatabase Items...")
    with arcpy.EnvManager(workspace=gdb_path):
        arcpy.management.CompressFileGeodatabaseData(gdb_path, lossless=True)
        uncompressed = [failed_list.append({"Layer Name":f, "Error Type": "GDB Compression", "Error Message":"Failed to Compress"}) for dataset in arcpy.ListDatasets(feature_type="Feature") for f in arcpy.ListFeatureClasses(feature_dataset=dataset) if not arcpy.Describe(f).isCompressed]
        compression_status = "Successful" if len(uncompressed) == 0 else "Not Successful"
    logger.warning(f"Failed Compress Layers: {uncompressed}")
    logger.info(f"Compression Status: {compression_status}")

    return compression_status, failed_list


def checkLayerAccessibility(item_obj):
    try:
        if item_obj.layers:
            logger.debug(f"Feature Service Layers are accessible")
            return True
        
    except Exception as e:
        logger.error(f"Feature Service Layers are not Accessible!\n{e}")
        return False