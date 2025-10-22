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
    def __init__(self, portal_obj):
        self.logger = logging.getLogger("root.servicelayer.PortalService")
        self._PortalObj = portal_obj
        self.isMultilayer = True if "Multilayer" in self._PortalObj["typeKeywords"] else False
        self.isHosted = True if "Hosted Service" in self._PortalObj["typeKeywords"] else False
        self.portalUrl = f"{PORTAL_ITEM_URL}{self.ServiceId}"
        # self.portalServiceName = self._PortalObj.name
        # self.portalServiceUrl = self._PortalObj.url
        # self._PortalObjType
        # self._PortalObjId = self._PortalObj.id
        
        # self._PortalObjTitle = self._PortalObj.title
        # self._PortalObjCategories = self._PortalObj.categories
        # self._PortalObjCreatedDate = self._PortalObj.created ## Return Epoch
        # self._PortalObjModifiedDate = self._PortalObj.modified ## Return Epoch
        # self._PortalObjDescription = self._PortalObj.description
        # self._PortalObjSummary = self._PortalObj.snippet
        # self._PortalObjTags = self._PortalObj.tags
        # self._PortalObjCredits = self._PortalObj.accessInformation
        # self._PortalObjTermsOfUse = self._PortalObj.licenseInfo

    def getMetadataDictionary(self, text_type:str="html")->dict:
        md = {}
        md["title"] = self._PortalObj.title
        md["description"] = meta.formatMdItem(self._PortalObj.description, "description", text_type)
        md["summary"] = meta.formatMdItem(self._PortalObj.snippet, "summary", text_type)
        md["tags"] = meta.formatMdItem(self._PortalObj.tags, "tags", text_type)
        md["credits"] = meta.formatMdItem(self._PortalObj.accessInformation, "accessInformation", text_type)
        md["accessConstraints"] = meta.formatMdItem(self._PortalObj.licenseInfo, "licenseInfo", text_type)

        return md



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
        self.layerSchemaEditEpoch = self.layerProperties["editingInfo"]["schemaLastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None ## Return Epoch
        self.layerDataEditEpoch = self.layerProperties["editingInfo"]["dataLastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None## Return Epoch
        self.layerPropertiesEditEpoch = self.layerProperties["editingInfo"]["lastEditDate"] if hasattr(self.layerProperties,"editingInfo") else None## Return Epoch
        self.layerFields = self.layerProperties["fields"]  ## Returns a list of dictionaries 
        self.layerDescription = self.layerProperties["description"]
        self.layerCredits = self.layerProperties["copyrightText"]
        self.layerSpatialReferenceWkid = self.layerProperties["spatialReference"]["latestWkid"] if hasattr(self.layerProperties, "spatialReference") else self.layerProperties["sourceSpatialReference"]["latestWkid"]


    

    def getMetadataDictionary(self)->dict:
        md = {}
        md["title"] = self.layerName
        md["description"] = self.layerDescription if self.layerDescription else self.portalItemDescription
        md["summary"] = self.portalItemSummary
        md["tags"] = self.portalItemTags
        md["credits"] = self.layerCredits if self.layerCredits else self.portalItemCredits
        md["accessConstraints"] = self.portalItemTermsOfUse


    
    def dataCatalogDictionary(self)->dict:
        out_dict = {}
        out_dict["Service Name"] = self.portalServiceName if self.portalServiceName else self.layerUrl.split("/")[-3] ## This is to handle if the Portal Item is referencing a Map Service 
        out_dict["Item ID"] = self.layerItemId
        out_dict["Item Title"] = self.portalItemTitle
        out_dict["Item Created Date"] = self.epochToString(self.portalItemCreatedDate)
        out_dict["Item Last Edited Date"] = self.epochToString(self.portalItemModifiedDate)
        out_dict["Item Categories"] = self.portalItemCategories
        out_dict["Layer Name"] = self.layerName
        out_dict["Layer ID"] = self.layerId
        out_dict["Layer Last Edited Date"] = self.epochToString(self.layerDataEditDate)
        out_dict["Metadata - Description"] = meta.formatMdItem(self.layerDescription, "description", "plain")
        out_dict["Metadata - Summary"] = self.portalSummary
        out_dict["Metadata - Tags"] = meta.formatMdItem(self.portalTags, "tags", "plain")
        out_dict["Metadata - Credits"] = meta.formatMdItem(self.layerCredits, "accessconstraints", "plain")
        out_dict["Metadata - License Information"] = meta.formatMdItem(self.portalTermsOfUse, "licenseinfo", "plain")
        out_dict["AGOL URL"] = self.layerPortalUrl

        return out_dict
    
    def propertyDictionary(self, metadata_text_type:str="html")->dict:
        out_dict = {}
        out_dict["Layer Name"]=self.layerName
        out_dict["Feature Service Name"]=self.parentServiceName if self.parentServiceName else self._getServiceName()
        out_dict["Is Hosted"] = self.isHosted
        out_dict["Is Multilayer"] = self.isMultilayer
        #out_dict["Feature Service REST URL"]=self.parentServiceUrl
        out_dict["Portal Item ID"]=self.parentId
        out_dict["Portal Item Created Date"]=self.epochToString(self.portalCreatedDate)
        #out_dict["Portal Item Last Edit Date"]=self.epochToString(self.portalModifiedDate)
        out_dict["Layer Schema Edit Date"]=self.epochToString(self.layerSchemaEditDate)
        out_dict["Layer Data Edit Date"]=self.epochToString(self.layerDataEditDate)
        out_dict["Layer Properties Edit Date"]=self.epochToString(self.layerPropertiesEditDate)
        out_dict["Layer Field Names"]=", ".join([f["name"] for f in self.layerFields])
        out_dict["Layer Description"] = meta.formatMdItem(self.layerDescription, 'description', metadata_text_type)
        out_dict["Layer Credits"] = meta.formatMdItem(self.layerCredits, 'accessConstraints', metadata_text_type)
        out_dict["Portal Item Title"]=self.portalTitle
        out_dict["Portal Item Description"]=meta.formatMdItem(self.portalDescription, 'description', metadata_text_type)
        out_dict["Portal Item Summary"]=self.portalSummary
        out_dict["Portal Item Tags"]=meta.formatMdItem(self.portalTags, "tags", metadata_text_type)
        out_dict["Portal Item Credits"]=meta.formatMdItem(self.portalCredits, 'accessconstraints', metadata_text_type)
        out_dict["Portal Item Terms of Use"]=meta.formatMdItem(self.portalTermsOfUse, 'licenseinfo', metadata_text_type)


        return out_dict

    def recordDf(self)->pd.DataFrame:
        self._handleProjectBoundaryShp()
        shp_geom = [i[0] for i in arcpy.da.SearchCursor(self.projectBoundaryPath, ["SHAPE@"])][0]
        spatial_rel = "ENVELOPE_INTERSECTS"
        fields = [field["name"] for field in self.layerFields]
        df_columns = [f'{field["name"]} ({field["alias"]})' for field in self.layerFields]
        records = [list(row) for row in arcpy.da.SearchCursor(self.layerUrl.replace("'",""), [fields], spatial_filter=shp_geom, spatial_relationship=spatial_rel)]
        df = pd.DataFrame(data=records, columns=df_columns)

        return df
    

    

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


    def exportTiles(self, tpk_path, levels):

        return tpk_path

class PortalFile(PortalItem):
    def __init__(self, portal_obj, output_directory):
        super().__init__(portal_obj)
        self.output_directory = output_directory

    def _checkFolders(self):
        if not os.path.exists(self.output_directory):
            try:
                logger.info(f"Output Directory Does Not Exist. Creating Directory...")
                os.makedirs(self.output_directory)
            except Exception as e:
                logger.error(f"Failed To create {self.output_directory}\n{e}")
        elif
        return
    
    def downloadFile(self, folder_path):
        
        return 
        


