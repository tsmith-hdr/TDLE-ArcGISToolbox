#######################################################################################################################################################
## Logging
import logging
logger = logging.getLogger(f"main.utility")
#######################################################################################################################################################
import json
import sys
import os
import getpass
import zipfile
from pathlib import Path
from datetime import datetime

import arcpy
from arcgis.gis import GIS

sys.path.insert(0,str(Path(__file__).resolve().parents[2]))

from src.constants.values import *
################################################################################################################################################################

def getValueFromJSON(json_file, key):
    try:
       with open(json_file) as f:
            data = json.load(f)
            return data[key]

    except Exception as e:
        logger.error(e)
        print("Error: ", e)

def isTaskScheduler()->bool:
    """
    Checks to see if the standalone file is being run from the console or run in a scheduled task.
    """
    # Check for an environment variable that might indicate Task Scheduler
    logger.debug(os.getenv("SESSIONNAME"))
    logger.debug(os.getenv('SCHEDULER_LAUNCH'))
    print(os.getenv("SESSIONNAME"))
    print(os.getenv('SCHEDULER_LAUNCH'))
    return os.getenv('SESSIONNAME') != 'Console' and os.getenv('SCHEDULER_LAUNCH') is None



def valueTableToDictionary(metadata_str:str)->dict:
    out_dict = {}
    metadata_vt = arcpy.ValueTable(2)
    metadata_vt.loadFromString(metadata_str)
    for i in range(0, metadata_vt.rowCount):
        md_item = metadata_vt.getValue(i, 0)
        md_value = metadata_vt.getValue(i, 1)
        out_dict[LOCAL_SERVICE_LOOKUP[md_item]] = md_value.strip()

    return out_dict


def epochToDate(epoch):
    timestamp = datetime.fromtimestamp(epoch/1000)
    time_string = timestamp.strftime("%m/%d/%Y")

    return (timestamp,time_string)



def zip_fgdb(input_fgdb, output_zip_dir):
    """
    Zips an ArcGIS file geodatabase folder.

    :param input_folder: Path to the file geodatabase folder.
    :param output_zip: Path to the output zip file.
    """
    output_zip = os.path.join(output_zip_dir, f"{os.path.basename(input_fgdb)}.zip")
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(input_fgdb):
            for file in files:
                if not file.endswith(".lock"):
                    file_path = os.path.join(root, file)
                    # Add file to zip, preserving folder structure
                    arcname = os.path.relpath(file_path, input_fgdb)
                    zipf.write(file_path, arcname)

    return output_zip


def create_directory(directory_path):
    logger.info(f"Creating Directory {directory_path}...")
    try:
        if os.path.exists(directory_path):
            logger.warning(f"{directory_path} Already Exists.")
        else:
            os.mkdir(directory_path)
            logger.info(f"{directory_path} Created.")
    except Exception as e:
        logger.error(e)
        raise ValueError(f"Failed to Create Directory\n{e}")

    return directory_path



def verify_filepath(file_path):
    if os.path.exists(file_path):
        logger.info(f">> Download Confirmed.")
        #logger.info(f"{file_path}")
        return True
    else:
        logger.error(f">> Download File Not Found!!")
        #logger.error(f"{file_path} Failed to Download!")
        return False


