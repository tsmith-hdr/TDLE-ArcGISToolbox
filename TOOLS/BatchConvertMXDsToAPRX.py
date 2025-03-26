"""
Script documentation

- Tool parameters are accessed using arcpy.GetParameter() or 
                                     arcpy.GetParameterAsText()
- Update derived parameter values using arcpy.SetParameter() or
                                        arcpy.SetParameterAsText()
"""
import arcpy
import os


def getMxdList(batch_type, mxd_directory, mxd_list):
    if batch_type == "Folder":
        mxd_list = [os.path.join(mxd_directory, i) for i in os.listdir(mxd_directory) if i.endswith(".mxd")]
        
    elif batch_type == "File":
        mxd_list = [i.replace("'", "") for i in mxd_files.split(";")]
        
    return mxd_list

def renameLayout(layout_obj, mxd_path):
    arcpy.AddMessage(f"Old Layout Name: {layout_obj.name}")
    new_layout_name = f"{layout_obj.name} {_formatLongName(mxd_path)}"
    layout_obj.name = new_layout_name
    arcpy.AddMessage(f"Updated Layout Name: {layout_obj.name}")
    
    return 

def renameMaps(map_objs:list, mxd_path, layout_name):
    for map_obj in map_objs:
        arcpy.AddMessage(f"Old Map Name: {map_obj.name}")
        new_map_name = f"{map_obj.name} {_formatLongName(mxd_path, layout_name)}"
        map_obj.name = new_map_name
        arcpy.AddMessage(f"Updated Map Name: {map_obj.name}")
        
    return

def _formatLongName(path, layout_name=None):
    if layout_name:
        if not layout_name.endswith(".mxd)"):
            suffix_number = layout_name.split(")")[-1]
        else:
            suffix_number = ""
    else:
        suffix_number = ""
        
    last_two = path.split("\\")[-2:]
    joined_name = "--".join(last_two)
    
    return f"({joined_name}){suffix_number}"



def main(batch_type, aprx_path, mxd_directory=None, mxd_files=None):
    aprx = arcpy.mp.ArcGISProject(aprx_path)
    arcpy.AddMessage(f"APRX: {aprx.filePath}")
    
    mxd_list = getMxdList(batch_type, mxd_directory, mxd_files)
    arcpy.AddMessage(f"Number of MXDs: {len(mxd_list)}")
    
    for mxd in mxd_list:
        arcpy.AddMessage(f"MXD Path: {mxd}")
        layout = aprx.importDocument(mxd, True)
        
        arcpy.AddMessage(f"-- Renaming Layout...")
        renameLayout(layout, mxd)
        
        map_obj_list = [m.map for m in layout.listElements("MAPFRAME_ELEMENT")]
        
        arcpy.AddMessage(f"-- Renaming Maps ({len(map_obj_list)})... ")
        renameMaps(map_obj_list, mxd, layout.name)

    
    if aprx_path != "CURRENT":
        try:
            aprx.save()
            del aprx
        except Exception as e:
            arcpy.AddError(f"Save Failed; The APRX is probably open somewhere.\n{e}")

if __name__ == "__main__":

    batch_type = arcpy.GetParameterAsText(0)
    mxd_directory = arcpy.GetParameterAsText(1)
    mxd_files = arcpy.GetParameterAsText(2)
    current_flag = arcpy.GetParameter(3)
    if not current_flag:
        aprx_path = arcpy.GetParameterAsText(4)
    else:
        aprx_path = "CURRENT"
    arcpy.AddMessage(aprx_path)
    main(batch_type, aprx_path, mxd_directory, mxd_files)