import os
import zip



class FolderManager():
  def __init__(self, export_dir):
    self.base_dir = export_dir





  return
  def zipDirectory(self):

    return zipped_path

  def listFolders(self)->list:
    folder_list = []
    for folder_name in os.listdir(self.base_dir):
      folder_path = os.path.join(self.base_dir, folder_name)

      folder_list.append(LocalFolder(folder_path))
    return folders_list




class LocalFolder(FolderManager):
  def __init__(self):


    return


