import os
import re
import sys
import json
import logging
import shutil
import datetime
import pandas as pd
from pathlib import Path
from importlib import reload  
from zipfile import ZipFile

import arcpy
from arcpy import metadata as md

from arcgis.gis import GIS, ItemTypeEnum

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

from src.constants.paths import OUTPUTS_DIR, LOG_DIR
from src.functions import utility
from src.functions import email
from src.classes.services import ServiceLayer
########################################################################################################################################
## Environments
arcpy.env.overwriteOutput=True
########################################################################################################################################
##Globals
DATETIME_STR = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
GDB_DIR = os.path.join(OUTPUTS_DIR, "BackupServices", "gdb")
REPORT_DIR = os.path.join(OUTPUTS_DIR, "BackupServices", "report")
OUTPUT_REPORT = os.path.join(REPORT_DIR, f"BackupServices_{DATETIME_STR}.xlsx")
ZIP_DIR = os.path.join(OUTPUTS_DIR, "BackupServices", "zip")
ZIPPED_FILE = os.path.join(ZIP_DIR, f"BackupServices_{DATETIME_STR}.zip")
########################################################################################################################################
## Logging
logger = logging.getLogger(f"root.TOOL_BackupServices")
########################################################################################################################################


def main(gis_conn:GIS,spatial_reference:arcpy.SpatialReference, agol_folder_objs:list,backup_dir:str,include_exclude_flag:str, scheduled:bool,include_exclude_list:list=None)->None:
    logger.info(f"Scheduled: {scheduled}")
    logger.info(f"GIS Connection: {gis_conn}")
    logger.info(f"Geodatabase Path: {GDB_DIR}")
    logger.info(f"Spatial Reference: {spatial_reference.factoryCode}")
    logger.info(f"Excel Report: {OUTPUT_REPORT}")
    logger.info(f"Backup Directory: {backup_dir}")
    logger.info(f"Include/Exclude Flag: {include_exclude_flag}")
    logger.info(f"Service List: {include_exclude_list}")
    logger.info("~~"*100)
    logger.info("~~"*100)
#############################################################################################################################
    ## Output Data Frame Lists 
    df_list = []
    failed = []
#############################################################################################################################
    ## Creating the Local File GDB. If the Overwrite Outputs environment is set to False, this will fail
    logger.info(f"Creating Local File GDB...")
    local_gdb_path = os.path.join(GDB_DIR, f"ServiceBackup_{DATETIME_STR}.gdb")
    logger.info(f"Local GDB Path: {local_gdb_path}")
    arcpy.management.CreateFileGDB(out_folder_path=GDB_DIR, 
                                    out_name=f"ServiceBackup_{DATETIME_STR}")
    
    ## Change the default workspace to the New File GDB
    arcpy.env.workspace = local_gdb_path

    ## Iterate over the input folders and create a corresponding Feature Dataset. The Dataset Names are cleaned so as to not raise errors.
    for folder_obj in agol_folder_objs:
        logger.info(f"AGOL Folder: {folder_obj}")

        ## Retrieving the Portal Item Objects that need to be excluded or included in the backup.
        if include_exclude_flag.strip().lower()  == "include":
            item_obj_list=[item_obj for item_obj in folder_obj.list(item_type=ItemTypeEnum.FEATURE_SERVICE.value) if item_obj.title in include_exclude_list]
        elif include_exclude_flag.strip().lower() == "exclude":
            item_obj_list=[item_obj for item_obj in folder_obj.list(item_type=ItemTypeEnum.FEATURE_SERVICE.value) if item_obj.title not in include_exclude_list]
        elif include_exclude_flag.strip().lower()  == "all":
            item_obj_list=[item_obj for item_obj in folder_obj.list(item_type=ItemTypeEnum.FEATURE_SERVICE.value)]
        
        logger.info(f"AGOL Item Count: {len(item_obj_list)}")

        if len(item_obj_list) > 0:
            dataset_name = folder_obj.name.translate({ord(c): "" for c in "!@#$%^&*()[] {};:,./<>?\|`~-=_+"}) ## Removes any special characters
            logger.info(f"Formatted AGOL Folder Name: {dataset_name}")
            logger.info(f"Building Portal Item List...")

            ## Iterates over the Item Obj and creates a Service Layer Object for each of the layers in the service. 
            # This Class is stored in src/classes/servicelayer.py
            arcpy.AddMessage(f"Exporting Services...")
            for item_obj in item_obj_list:
                logger.info(f"AGOL Item: {item_obj}")
                try:
                    logger.info(f"Layer Count: {len(item_obj.layers)}")
                    for layer_obj in item_obj.layers:
                        sl_obj = ServiceLayer(gis_conn, layer_obj, item_obj)
                        logger.debug(sl_obj)
                        logger.info(f"Layer: {sl_obj.layerName}")
                        ## Exports the Service Layer to the Local GDB.
                        # The export method will only export features that intersect with the extent of the study area. 
                        # returns a dictionary of items to update the backedup feature class metadata.
                        try:
                            out_dict = sl_obj.exportLayer(out_workspace=os.path.join(local_gdb_path, dataset_name))
                        except Exception as e:
                            logger.error(f"Failed::{e}")
                            failed.append({"Layer":sl_obj.layerName, "Action":"Export Layer", "Error":e})

                        out_dict["Folder Name"] = folder_obj.name

                        df_list.append(out_dict)
                        logger.debug(f"Layer Metadata Dictionary:\n{out_dict}")
                        logger.info(f"Updating Layer Metadata...")
                        try:
                            meta = md.Metadata(out_dict["Feature Class Path"])
                            meta.summary = f"Created as part of a backup on {DATETIME_STR.split('-')[0]} performed by {os.getlogin()}"
                            meta.tags = f"Layer Name:{out_dict['Layer Name']}, Layer URL:{out_dict['Layer URL']}, Service Item Id:{out_dict['Service Item Id']}"
                            meta.credits = out_dict["Service Credits"]
                            meta.save()
                            logger.info(f"Save Successfull...")
                        except Exception as m:
                            logger.error(f"Failed To Update Layer Metadata: {m}")
                            failed.append({"Layer":sl_obj.layerName, "Action":"Update Feature Class Metadata", "Error":m})
                except Exception as p:
                    logger.error(f"Failed to access Service Info\n{p}")
                    failed.append({"Layer":item_obj.title, "Action":"Access Info", "Error":p})

    ## Here we are compressing the file gdb this is a lossl_objess function. We want to add this process to make sure that the archived records are unable to be editied.
    logger.info(f"Compressing Local GDB Items...")
    arcpy.AddMessage(f"Compressing Local GDB Items...")
    with arcpy.EnvManager(workspace=local_gdb_path):
        arcpy.management.CompressFileGeodatabaseData(local_gdb_path, lossless=True)
        uncompressed = [failed.append({"Layer":f, "Action": "GDB Compression", "Error":"Failed to Compress"}) for dataset in arcpy.ListDatasets(feature_type="Feature") for f in arcpy.ListFeatureClasses(feature_dataset=dataset) if not arcpy.Describe(f).isCompressed]
        compression_status = "Successful" if len(uncompressed) == 0 else "Not Successful"
    logger.warning(f"Failed Compress Layers: {uncompressed}")
    logger.info(f"Compression Status: {compression_status}")


    ## Updates the Backup File GDB Metadata with the input parameters and user info.
    logger.info(f"Updating Local FGDB Metadata...")
    meta = md.Metadata(local_gdb_path)
    meta.summary = f"This File GDB is a backup of the Feature Services. Run by {os.getlogin()} using AGOL User {gis_conn.users.me.username}"

    param_dict = {
    "GIS Connection":str(gis_conn),
    "GIS User":str(gis_conn.users.me.username),
    "Local User":os.getlogin(),
    "Include/Exclude Flag":include_exclude_flag,
    "Folders":[f.name for f in agol_folder_objs],
    "Service List":include_exclude_list,
    "Spatial Reference":str(spatial_reference.factoryCode),
    "Compression":compression_status
    }

    meta.description = json.dumps(param_dict, indent=1)
    try:
        meta.save()
        logger.info(f"Save Successfull")
    except Exception as d:
        logger.error(f"Metadata Failed To Save: {d}")
        failed.append({"Layer":local_gdb_path, "Action":"Failed to Update FGDB Metadata", "Error":d})

    ## 
    logger.info(f"Creating Pandas Data Frames...")
    param_df = pd.DataFrame.from_dict(param_dict, orient="index", columns=["Value"])
    update_df = pd.DataFrame(df_list, columns=["Feature Class Name","Layer Name", "Folder Name","Service Item Id","Feature Class Path", "Layer URL"])
    failed_df = pd.DataFrame(failed)

    logger.debug(param_df.head())
    logger.debug(update_df.head())
    logger.debug(failed_df.head())



    logger.info(f"Exporting Pandas Data Frames...")
    try:
        with pd.ExcelWriter(OUTPUT_REPORT) as writer:
            update_df.to_excel(writer, sheet_name="UpdatedItems", index=False)
            param_df.to_excel(writer, sheet_name="InputParams", header=True,index=True, index_label="Parameter")
            failed_df.to_excel(writer, sheet_name="FailedActions", index=False)
    except Exception as e:
        logger.error(f"Failed to Export Excel Report. {e}")


    ## Here we are compress together the excel report and the local filegdb. zipping these items will make it more efficient to send from local machine to the backup directory. 
    logger.info(f"Zipping Local GDB and Excel Reports...")
    arcpy.AddMessage(f"Zipping Local GDB and Excel Reports...")
    
    try:
        zipped_fgdb = utility.zip_fgdb(local_gdb_path, os.path.join(ZIP_DIR, "Zipped_FGDB"))
        with ZipFile(ZIPPED_FILE, 'w') as zip:
            zip.write(zipped_fgdb, os.path.basename(zipped_fgdb))
            zip.write(OUTPUT_REPORT, os.path.basename(OUTPUT_REPORT))
            logger.info('All files zipped successfully!')
    except Exception as r:
        arcpy.AddError(f"Failed to Zip Files.\n{r}")
        logger.error(f"Failed to Zip Files.\n{r}")



    ## Copies the zipped folder of the local file gdb to the designated directory to hold the weekly backups. 
    ### If we want to I can add logic to unzip the folder...
    if backup_dir:
        logger.info("Copying Zipped Folder...")
        try:
            shutil.copy(ZIPPED_FILE, os.path.join(backup_dir, os.path.basename(ZIPPED_FILE)))
        except Exception as u:
            logger.error(f"Failed Copy: {u}")

        # Checks if the zipped folder was successfully copied.
        if os.path.exists(os.path.join(backup_dir, os.path.basename(ZIPPED_FILE))):
            logger.info(f"Excel Report Has Been Exported to: {OUTPUT_REPORT}")
        else:
            logger.error(f"Excel Report Failed to Export to: {OUTPUT_REPORT}")
    else:
        logger.info("No Backup Directory Named.")

    
    return OUTPUT_REPORT, ZIPPED_FILE







def exportFeatureServices(gis_conn:GIS, item_list:list, out_gdb:Path=None, out_dir:Path=None)->pd.DataFrame:
    failed = []
    df_list = []

    if len(item_list) > 0:

        ## Iterates over the Item Obj and creates a Service Layer Object for each of the layers in the service. 
        # This Class is stored in src/classes/servicelayer.py
        logger.info(f"Exporting Services...")

        for item in item_list:
            logger.info(f"AGOL Item: {item}")
            try:
                logger.info(f"Layer Count: {len(item.layers)}")
                for layer in item.layers:
                    sl_obj = ServiceLayer(gis_conn, layer, item)
                    logger.debug(sl_obj)
                    logger.info(f"Layer: {sl_obj.layerName}")
                    ## Exports the Service Layer to the Local GDB.
                    # The export method will only export features that intersect with the extent of the study area. 
                    # returns a dictionary of items to update the backedup feature class metadata.
                    try:
                        out_dict, featureclass_path = sl_obj.exportLayer(out_workspace=out_gdb)
                    except Exception as e:
                        logger.error(f"Failed::{e}")
                        failed.append({"Layer":sl_obj.layerName, "Action":"Export Layer", "Error":e})

                    #out_dict["Folder Name"] = folder_obj.name

                    df_list.append(out_dict)
                    logger.debug(f"Layer Metadata Dictionary:\n{out_dict}")
                    logger.info(f"Updating Layer Metadata...")
                    try:
                        meta = md.Metadata(featureclass_path)
                        meta.summary = f"Created as part of a backup on {DATETIME_STR.split('-')[0]} performed by {os.getlogin()}"
                        meta.tags = f"Layer Name:{out_dict['Layer Name']}, Layer URL:{out_dict['Layer URL']}, Service Item Id:{out_dict['Service Item Id']}"
                        meta.credits = out_dict["Service Credits"]
                        meta.save()
                        logger.info(f"Save Successfull...")
                    except Exception as m:
                        logger.error(f"Failed To Update Layer Metadata: {m}")
                        failed.append({"Layer":sl_obj.layerName, "Action":"Update Feature Class Metadata", "Error":m})
            except Exception as p:
                logger.error(f"Failed to access Service Info\n{p}")
                failed.append({"Layer":item_obj.title, "Action":"Access Info", "Error":p})
    return