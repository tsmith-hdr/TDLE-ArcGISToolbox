import os
import sys
import json
import datetime

import logging
import pandas as pd
from pathlib import Path
from string import digits


import arcpy
import arcgis
from arcgis.gis import Item

if arcgis.__version__.startswith("2.3"):
    from arcgis.mapping import MapImageLayer
elif arcgis.__version__.startswith("2.4"):
    from arcgis.layers import MapImageLayer

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from src.functions import meta, utility, arc
from src.constants.paths import PORTAL_ITEM_URL
#################################################################################################################################################################################
logger = logging.getLogger("main.servicelayer")
#################################################################################################################################################################################

class PortalItem():
    def __init__(self, gis_conn ,item_obj):
        self.logger = logging.getLogger("main.servicelayer.PortalService")
        self._GIS = gis_conn
        self._ItemObj = item_obj
        self.isMultilayer = True if "Multilayer" in self._ItemObj["typeKeywords"] else False
        self.isHosted = True if "Hosted Service" in self._ItemObj["typeKeywords"] else False
        self.areLayersAccessible = arc.checkLayerAccessibility(self._ItemObj) if self._ItemObj.type == "Feature Service" else None
        self.portalUrl = f"{PORTAL_ITEM_URL}{self._ItemObj.id}"
        self.sharingLevel = self._ItemObj.sharing.shared_with["level"].value
        self.sharingGroups = [g.title for g in self._ItemObj.sharing.shared_with["groups"]]
        self.ownerFolderName = self._GIS.content.folders.get(self._ItemObj.ownerFolder).name


    def getItemMetadataDictionary(self, text_type:str="html")->dict:
        md = {}
        md["title"] = self._ItemObj.title
        md["description"] = meta.formatMdItem(self._ItemObj.description, "description", text_type)
        md["summary"] = meta.formatMdItem(self._ItemObj.snippet, "summary", text_type)
        md["tags"] = meta.formatMdItem(self._ItemObj.tags, "tags", text_type)
        md["credits"] = meta.formatMdItem(self._ItemObj.accessInformation, "accessInformation", text_type)
        md["accessConstraints"] = meta.formatMdItem(self._ItemObj.licenseInfo, "licenseInfo", text_type)

        return md


    def getItemExcelDictionary(self, md_text_type:str="plain")->dict:
        item_dict = {}
        item_dict["Item Title"] = self._ItemObj.title
        item_dict["Service Name"] = self._ItemObj.name
        item_dict["Service URL"] = self._ItemObj.url
        item_dict["Item Id"] = self._ItemObj.id
        item_dict["Item Org Id"] = self._ItemObj.orgId
        item_dict["Item Hosted"] = self.isHosted
        item_dict["Item Multilayer"] = self.isMultilayer
        item_dict["Item URL"] = self.portalUrl
        item_dict["Item Created Date"] = utility.epochToDate(self._ItemObj.created)[1]
        item_dict["Item Modified Date"] = utility.epochToDate(self._ItemObj.modified)[1]
        item_dict["Item Sharing Level"] = self.sharingLevel
        item_dict["Item Sharing Groups"] = self.sharingGroups
        item_dict["Item Owner Folder"] = self.ownerFolderName
        item_dict["Item description"] = meta.formatMdItem(self._ItemObj.description, "description", md_text_type)
        item_dict["Item summary"] = meta.formatMdItem(self._ItemObj.snippet, "summary", md_text_type)
        item_dict["Item tags"] = meta.formatMdItem(self._ItemObj.tags, "tags", md_text_type)
        item_dict["Item credits"] = meta.formatMdItem(self._ItemObj.accessInformation, "accessInformation", md_text_type)
        item_dict["Item accessConstraints"] = meta.formatMdItem(self._ItemObj.licenseInfo, "licenseInfo", md_text_type)

        return item_dict




class ServiceLayer(PortalItem):
    def __init__(self, gis_conn, portal_obj, layer_obj):
        super().__init__(gis_conn, portal_obj)
        self.logger = logging.getLogger("main.servicelayer.ServiceLayer")
        self._LayerObj = layer_obj
        self.layerProperties = self._LayerObj.properties
        self.layerName = self.layerProperties["name"]
        self.layerId = self.layerProperties["id"]
        self.layerItemId =self.layerProperties["serviceItemId"]
        self.layerPortalUrl = f"{self.portalUrl}&sublayer={self.layerId}"
        self.layerSchemaEdit = self.layerProperties["editingInfo"]["schemaLastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None ## Return Epoch
        self.layerDataEdit = self.layerProperties["editingInfo"]["dataLastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None## Return Epoch
        self.layerPropertiesEdit = self.layerProperties["editingInfo"]["lastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None## Return Epoch
        self.layerFields = self.layerProperties["fields"]  ## Returns a list of dictionaries
        self.layerDescription = self.layerProperties["description"]
        self.layerCredits = self.layerProperties["copyrightText"]
        self.layerSpatialReference = self._getSR()
        self.layerOrgId = self._getOrgID()

    def _getOrgID(self):
        layer_url = self._LayerObj.url

        striped = layer_url.replace("https://", "")
        split = striped.split("/")

        return split[1] if "arcgis.com" in layer_url else split[0]

    def _getSR(self):

        if hasattr(self.layerProperties, "spatialReference"):
            sr = self.layerProperties["spatialReference"]
            if hasattr(sr, "latestWkid"):
                return self.layerProperties["spatialReference"]["latestWkid"]
            else:
                return self.layerProperties["spatialReference"]

        elif hasattr(self.layerProperties, "sourceSpatialReference"):
            sr = self.layerProperties["sourceSpatialReference"]
            if hasattr(sr, "latestWkid"):
                return self.layerProperties["sourceSpatialReference"]["latestWkid"]
            else:
                return self.layerProperties["sourceSpatialReference"]
        else:
            return None


    def getLayerExcelDictionary(self, md_text_type:str="plain")->dict:
        layer_dict = self.getItemExcelDictionary()
        layer_dict["Layer Name"] = self.layerName
        layer_dict["Layer Id"] = self.layerId
        layer_dict["Layer Org Id"] = self.layerOrgId
        layer_dict["Layer Portal URL"] = self.layerPortalUrl
        layer_dict["Layer Service URL"] = self._LayerObj.url
        layer_dict["Layer Spatial Reference"] = self.layerSpatialReference
        layer_dict["Layer description"] = meta.formatMdItem(self.layerDescription, "description", md_text_type)
        layer_dict["Layer Credits"] = self.layerCredits
        layer_dict["Layer Schema Edit Date"] = utility.epochToDate(self.layerSchemaEdit)[1]
        layer_dict["Layer Data Edit Date"] = utility.epochToDate(self.layerDataEdit)[1]
        layer_dict["Layer Properties Edit Date"] = utility.epochToDate(self.layerPropertiesEdit)[1]

        return layer_dict



    def getLayerMetadataDictionary(self, text_type:str="html")->dict:
        """
        Generates a dictionary that
        """
        md = {}
        md["title"] = f"Item: {self._ItemObj.title} Layer: {self.layerName}"
        md["description"] = f"Item: {meta.formatMdItem(self._ItemObj.description, 'description', text_type)}\nLayer: {meta.formatMdItem(self.layerDescription, 'description', text_type)}"
        md["summary"] = meta.formatMdItem(self._ItemObj.snippet, 'summary', text_type)
        md["tags"] = meta.formatMdItem(self._ItemObj.tags, 'tags', text_type)
        md["credits"] = f"Item: {meta.formatMdItem(self._ItemObj.accessInformation, 'accessInformation', text_type)}\nLayer: {meta.formatMdItem(self.layerCredits, 'accessInformation', text_type)}"
        md["accessConstraints"] = meta.formatMdItem(self._ItemObj.licenseInfo, 'licenseInfo', text_type)

        return md

    def _formatLayerName(self, current_feature_classes:list)->str:
        """
        Formats the Layer Name so it can be exported to a File Geodatabase.
        This Replaces any special characters with underscores and removes any leading digits.
        Also checks for duplicate named feature classes and handles a numeric suffix.

        Input Parameters:
        - current_feature_classes : Use the arcpy.ListFeatureClasses() method to retrive all of the current feature classes. This will check if any duplicate feature class names exist
        """
        formatted_name = self.layerName.translate({ord(c): "_" for c in "!@#$%^&*()[] {};:,./<>?ə\|`~-=+"}) ## Replaces Special Characters with underscores
        formatted_name = formatted_name.lstrip(digits) ## Removes leading digits

        ## Handles Any Layers that Have Duplicate Names
        count = 1
        while True:
            if formatted_name not in current_feature_classes:
                break
            else:
                if count==1:
                    formatted_name = f"{formatted_name}_{count}"
                else:
                    formatted_name = f"{formatted_name.rsplit('_',1)[0]}_{count}"

                count+=1
        return formatted_name

    def exportLayer(self, out_gdb:Path)->dict:
        failed_dict = {}

        with arcpy.EnvManager(workspace=out_gdb):
            feature_names = arcpy.ListFeatureClasses()
            logger.debug(f"Current Feature Classes:\n{feature_names}")
            logger.info

        formatted_layer_name = self._formatLayerName(feature_names)



        self.logger.info(f"Formatted Name: {formatted_layer_name}")

        try:
            self.logger.info("Exporting...")
            featureclass_path = os.path.join(out_gdb, formatted_layer_name)
            self.logger.info(f"Feature Class Path: {featureclass_path}")
            with arcpy.EnvManager(preserveGlobalIds=True, maintainAttachments=True):
                arcpy.conversion.ExportFeatures(in_features=self._LayerObj.url,
                                                out_features=featureclass_path)

        except Exception as f:
            self.logger.error(f"{self.layerName:30s} {self._ItemObj.id:30s}\n{f}")
            failed_dict["Item ID"] = self._ItemObj.id
            failed_dict["Layer Name"] = self.layerName
            failed_dict["Formatted Layer Name"] = formatted_layer_name
            failed_dict["Intended Path"] = featureclass_path
            failed_dict["Error Type"] = "Failed Export"
            failed_dict["Error Message"] = f



        return featureclass_path, failed_dict



class TiledService(PortalItem):
    def __init__(self, gis_conn, portal_obj):
        super().__init__(gis_conn, portal_obj)
        self._MapImageLayer = MapImageLayer(self._ItemObj.url, gis_conn)
        self._Service = self._MapImageLayer.service

        self.milProperties = self._MapImageLayer.properties
        self.serProperties = self._Service.properties
        self.name = self.milProperties["name"]
        self.allowExport = self.milProperties["exportTilesAllowed"]
        self.maxExport = self.milProperties["maxExportTilesCount"]
        self.minLod = self.milProperties["minLOD"]
        self.maxLod = self.milProperties["maxLOD"]
        self.lodRange = f"{self.minLod}-{self.maxLod}"
        self.spatialReference = self.milProperties["spatialReference"]
        self.layerCount = len(self.milProperties["layers"])
        self.documentInfo = self.milProperties["documentInfo"]
        self.size = serProperties["size"]
        self.tileCount = serProperties["count"]
        self.serviceUrl = serProperties["url"]
        self.extent = self.extent


    def _determineExportExtents(self)->dict:
        extents_dict = {}
        initial_extent = arcpy.Extent(
            XMin=self.extent["xmin"],
            YMin=self.extent["ymin"],
            XMax=self.extent["xmax"],
            YMax=self.extent["ymax"],
            spatial_referernce=arcpy.SpatialReference(serProperties["spatialReference"]["latestWkid"])
        )
        initial_polygon = initial_extent.polygon


        split_angle = self._determineSplitAngle(self)
        number_of_splits = self._determineNumberOfSplits

        split_polygons = arcpy.management.SubdividePolygon(
            in_polygons=initial_polygon,
            out_feature_class="memory/temp_polys",
            method="NUMBER_OF_EQUAL_PARTS",
            num_areas=number_of_splits,
            split_angle=split_angle,
            subdivision_type="STRIPS"
        )

        ## Need to Finish
        return extents_dict


    def _determineSplitAngle(self)->int:

        return split_angle

    def _determineNumberOfSplits(self)->int:

        return number_of_splits









    def exportTiles(self, output_directory, levels):

        temp_path = self._MapImageLayer.export_tiles(self.lodRange, "levelId", True, storage_format="tpkx")

        tpkx_path=os.path.join(output_dir, f"{self.milName}.tpkx")

        shutil.move(src=temp_path, dst=tpkx_path)

        return tpk_path





class PortalFile(PortalItem):
    def __init__(self, gis_conn, item_obj):
        super().__init__(gis_conn, item_obj)

    def getFileExcelDictionary(self):
        file_dict = self.getItemExcelDictionary()
        file_dict["File Type"] = self._ItemObj.type

        return file_dict



    def downloadFile(self, output_directory):
        failed_dict = {}
        logger.info(f"Downloading {self._ItemObj.title}...")
        try:
            file_path = self._ItemObj.download(output_directory, self._ItemObj.name)
            utility.verify_filepath(file_path)
            logger.info(f"Output File Path: {file_path}")
        except Exception as e:
            self.logger.error(f"{self._ItemObj.name:30s} {self._ItemObj.id:30s}")
            failed_dict["Item ID"] = self._ItemObj.id
            failed_dict["Layer Name"] = self._ItemObj.name
            failed_dict["Formatted Layer Name"] = "N/A"
            failed_dict["Intended Path"] = file_path
            failed_dict["Error Type"] = "Failed Download"
            failed_dict["Error Message"] = e


        return file_path, failed_dict



