from pathlib import Path
import urllib
import arcpy
import sys
import os

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))





VALID_FILE_TYPES = [
    "Service Definition",
    "CSV",
    'Microsoft Word',
    "Administrative Report",
    "Shapefile",
    "File Geodatabase",
    'Layer Package',
    'Vector Tile Package',
    'Tile Package',
    'Notebook',
    "Desktop Style"
    ]