import os
import sys
import json
import datetime
import logging
import pandas as pd
from pathlib import Path

import arcpy
from arcgis.gis import Item

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from src.functions import meta
from src.constants.paths import SHP_DIR, PORTAL_ITEM_URL
from src.constants.values import PROJECT_SPATIAL_REFERENCE
#################################################################################################################################################################################
logger = logging.getLogger("root.servicelayer")
#################################################################################################################################################################################

class PortalItem():
    def __init__(self, item_obj):
        self.logger = logging.getLogger("root.servicelayer.PortalService")
        self._ItemObj = item_obj
        self.isMultilayer = True if "Multilayer" in self._ItemObj["typeKeywords"] else False
        self.isHosted = True if "Hosted Service" in self._ItemObj["typeKeywords"] else False
        self.portalUrl = f"{PORTAL_ITEM_URL}{self._ItemObj.id}"


    def getMetadataDictionary(self, text_type:str="html")->dict:
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
        item_dict["Item Hosted"] = self.isHosted
        item_dict["Item Multilayer"] = self.isMultilayer
        item_dict["Item URL"] = self.portalUrl
        item_dict["Item Created Date"] = utility.epochToDate(self._ItemObj.created)[1]
        item_dict["Item Modified Date"] = utility.epochToDate(self._ItemObj.modified)[1]
        item_dict["Item description"] = meta.formatMdItem(self._ItemObj.description, "description", text_type)
        item_dict["Item summary"] = meta.formatMdItem(self._ItemObj.snippet, "summary", text_type)
        item_dict["Item tags"] = meta.formatMdItem(self._ItemObj.tags, "tags", text_type)
        item_dict["Item credits"] = meta.formatMdItem(self._ItemObj.accessInformation, "accessInformation", text_type)
        item_dict["Item accessConstraints"] = meta.formatMdItem(self._ItemObj.licenseInfo, "licenseInfo", text_type)

        return item_dict

class ServiceLayer(PortalItem):
    def __init__(self, portal_obj, layer_obj):
        super().__init__(portal_obj)
        self.logger = logging.getLogger("root.servicelayer.ServiceLayer")
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
        self.layerSpatialReferenceWkid = self.layerProperties["spatialReference"]["latestWkid"] if hasattr(self.layerProperties, "spatialReference") else self.layerProperties["sourceSpatialReference"]["latestWkid"]


    def getLayerExcelDictionary(self, md_text_type:str="plain")->dict:
        layer_dict = self.getItemExcelDictionary()
        layer_dict["Layer Name"] = self.layerName
        layer_dict["Layer Id"] = self.layerId
        layer_dict["Layer Portal URL"] = self.layerName
        layer_dict["Layer Service URL"] = self._LayerObj.url
        layer_dict["Layer Spatial Reference"] = self.layerSpatialReferenceWkid
        layer_dict["Layer description"] = self.layerDescription
        layer_dict["Layer Credits"] = self.layerCredits
        layer_dict["Layer Schema Edit Date"] = utility.epochToDate(self.layerSchemaEdit)[1]
        layer_dict["Layer Data Edit Date"] = utility.epochToDate(self.layerDataEdit)[1]
        layer_dict["Layer Properties Edit Date"] = utility.epochToDate(self.layerPropertiesEdit)[1]


        return layer_dict



    def getMetadataDictionary(self)->dict:
        md = {}
        md["title"] = self.layerName
        md["description"] = self.layerDescription if self.layerDescription else self.portalItemDescription
        md["summary"] = self.portalItemSummary
        md["tags"] = self.portalItemTags
        md["credits"] = self.layerCredits if self.layerCredits else self.portalItemCredits
        md["accessConstraints"] = self.portalItemTermsOfUse



    def exportLayer(self, out_gdb:Path)->dict:

        metadata_dict = {}
        feature_names = [f for f in arcpy.ListFeatureClasses()]
        formatted_layer_name = self.layerName.translate({ord(c): "_" for c in "!@#$%^&*()[] {};:,./<>?\|`~-=+"})
        ## Need to add logic to replace leading digits ##


        ## Handles Any Layers that Have Duplicate Names
        count = 1
        while True:
            if formatted_layer_name not in feature_names:
                break
            else:
                if count==1:
                    formatted_layer_name = f"{formatted_layer_name}_{count}"
                else:
                    formatted_layer_name = f"{formatted_layer_name.rsplit('_',1)[0]}_{count}"

                count+=1


        self.logger.info(f"Formatted Name: {formatted_layer_name}")

        try:
            self.logger.info("Exporting...")
            featureclass_path = os.path.join(out_gdb, formatted_layer_name)
            with arcpy.EnvManager(preserveGlobalIds=True):
                arcpy.conversion.ExportFeatures(in_features=self.layerUrl,
                                                out_features=featureclass_path)
        except Exception as f:
            self.logger.error(f"{self.layerName:30s} {self.parentId:30s}")
            arcpy.AddWarning(f"Layer Failed to Export:\n{self.layerName:30s} {self.parentId:30s}")



        return featureclass_path



class TiledService(PortalItem):
    def __init__(self, portal_obj):
        super().__init__(portal_obj)


    def exportTiles(self, output_directory, levels):



        return tpk_path

class PortalFile(PortalItem):
    def __init__(self, item_obj):
        super().__init__(item_obj)

    def getFileExcelDictionary(self):
        file_dict = self.getItemExcelDictionary()


        return file_dict



    def downloadFile(self, output_directory):
        logger.info(f"Downloading {self._ItemObj.title}...")
        try:
            file_path = self._ItemObj.download(output_directory, self._ItemObj.name)
            utility.verify_filepath(file_path)
        except Exception as e:
            logger.error(f"Failed to download {self._ItemObj.title} ({self._ItemObj.id})\n{e}")


        return file_path



