#######################################################################################################################################################
## Logging
import logging
logger = logging.getLogger("main.meta")
#######################################################################################################################################################
## Libraries
import sys
import re
from pathlib import Path
from bs4 import BeautifulSoup
from bs4.element import Tag

import arcpy.metadata as md

if str(Path(__file__).resolve().parents[2]) not in sys.path:
    sys.path.insert(0,str(Path(__file__).resolve().parents[2]))


from src.constants.values import *


#######################################################################################################################################################
## Functions



        
def formatMdItem(text:str, md_item:str, text_type:str)->str:
    """
    Using ReGex to strip HTML Tags from the Description, Summary, and Access Constraints.
    The Item Tags are also formatted from a string to a list and all spaces are striped and the list is sorted.
    md_items can be "description", accessInformation", "licenseInfo", "tags"
    """

    if text and text_type.lower() == "plain" and md_item.lower().replace(" ","") in ["description", "accessinformation", "licenseinfo"]:
        bs = BeautifulSoup(text,'html.parser')

        a_tags = bs.find_all('a', href=True)
        for a_tag in a_tags:
            if a_tag.contents:
                a_tag_contents = a_tag.contents[0]

                if isinstance(a_tag.contents[0], Tag):
                    a_tag_contents = a_tag.contents[0].text
                
                if not a_tag_contents:
                    new_content = f"{a_tag['href']}"
                elif a_tag_contents.startswith("http"):
                    new_content = f"{a_tag['href']}"
                else:
                    new_content = f"{a_tag_contents} ({a_tag['href']})"

                a_tag.contents[0].replace_with(new_content)

            else:
                logger.warning(f"!! HTML Description/AccessConstraint/LicenseInfo is missing <a> tag content.")

        clean_text = bs.get_text()
        #clean_text = re.sub(r'<.*?>', '', new_text)
        
        
    elif text and md_item == 'tags':
        if type(text) == str:
            tag_list = text.split(",")

        if type(text) == list:
            tag_list = text

        cleaned_list = [t.strip() for t in tag_list]

        sorted_list = sorted(cleaned_list)

        clean_text = ",".join(sorted_list)

        
    else:
        clean_text = text


    return clean_text
    

def _cleanCheckText(text):
    notags_text = re.sub(r'<.*?>', '', text)
    lower_text = notags_text.lower()
    
    #nofollow_text = lower_text.replace(" rel='nofollow ugc'", "")
    standard_quotes = lower_text.replace('"',"'")
    text_strip = standard_quotes.strip()
    clean_text = text_strip
    
    return clean_text

