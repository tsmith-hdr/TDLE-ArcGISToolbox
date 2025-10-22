#######################################################################################################################################################
## Logging
import logging
logger = logging.getLogger(f"root.arc")
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

    
    
    if source_list:
        for source in source_list:
            logger.info(f"Source: {source}")
            if checkSource(gis_conn, source_type, source) and source_type in ['folder', 'group']:
                content_list = search_method[source_type.lower()](source).list(item_type=ItemTypeEnum.FEATURE_SERVICE.value) if source_type.lower() == 'folder' else [c for c in search_method[source_type.lower()](source).content() if c.type == "Feature Service"]

                if include_exclude_flag.lower() == "include":
                    [item_list.append(i) for i in content_list if i.id in include_exclude_list and i.id not in item_id_list and i.type in item_types]
                elif include_exclude_flag.lower() == "exclude":
                    [item_list.append(i) for i in content_list if i.id not in include_exclude_list and i.id not in item_id_list and i.type in item_types]
                else:
                    [item_list.append(i) for i in content_list if i.id not in item_id_list and i.type in item_types]
    else:
        [item_list.append(i.id) for i in search_method["catalog"]("", max_items=-1) if i.id not in item_id_list and i.type in item_types]



    return item_list