from pathlib import Path
import urllib
import arcpy
import sys
import os

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

## Local Paths
ROOT_DIR = Path(__file__).resolve().parents[2]

LOG_DIR = Path(ROOT_DIR, "logs")

OUTPUTS_DIR = Path(ROOT_DIR, "outputs")

BACKUPS_DIR = os.path.join(OUTPUTS_DIR, "backups")

## AGOL Paths
PORTAL_URL = "https://arcgis.com/"

PORTAL_ITEM_URL = urllib.parse.urljoin(PORTAL_URL, "home/item.html?id=")


## Email SMTP Path

SMTP_PATH = 'smtp.hdrinc.com'