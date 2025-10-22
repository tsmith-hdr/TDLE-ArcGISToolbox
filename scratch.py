from arcgis.gis import GIS

gis = GIS()
content = gis.content.get("c4cba79ad1194d7e9a08ea8f0e35c8e2")
print(content)