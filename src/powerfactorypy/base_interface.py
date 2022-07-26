import sys
#sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10')
#import powerfactory as pf
sys.path.append(r".\src")
import powerfactorypy
from os import path as os_path
import math

# ToDo: get_active_networks, copy_graphics_pages

class PFBaseInterface:
  """Base interface for data manipulation
  """

  def __init__(self,app):  
    
    self.app = app
    self.export_path = ""
    self.csv_export_file_name = "SimulationResults.csv"    
    self.results = "All calculations.ElmRes"
    self.project_language = "English" 

  def get_obj(self,path,project_folder=None,error_if_non_existent=True):
    """
    Returns the PowerFactory object under path in the project folder.
    The path must be specified relative to the project folder.
    If argument `project_folder` is not specified, the path is
    searched within the currently active project.
    ToDo: Does specifying a project folder object enhance performance? 

    Examples:
      pfbi.get_obj("Network Model\\Network Data\\Grid\\Terminal 1")
    The path can also start with "\\":
      get_obj("\\Network Model\\Network Data\\Grid\\Terminal 1")
    Note that you can also use r" at the beginning of the string
    argument to use single "\".   
    """

    try:
      if path[0] == "\\":
        path = path[1:] 
      splitted_path = path.split("\\")
    except(TypeError):
      raise TypeError("The argument 'path' must be of type string.")
    if project_folder:
      folder_obj = [project_folder]
    else:
      project_folder = self.app.GetActiveProject()
      folder_obj = [project_folder]
    try:
      for folder_name in splitted_path:
        folder_obj = folder_obj[0].GetContents(folder_name + '.*')
      return folder_obj[0]
    except(IndexError):
      if error_if_non_existent:
        raise powerfactorypy.PFPathError(path,project_folder)
      else:
        return None

  def get_by_attribute(self,objects,attr,attr_lambda):
    """From a list of objects, get those for whom the the attr_lambda
    function returns 'True' when given the attribute attr of an
    object as input.
    Example:
      pfbi.get_by_attribute(list_of_objects,"uknom",lambda x : x==110)
    """
    objects_true = []
    for obj in objects:
      try:
        # This lambda function is problematic because it does
        # not always throw an error when it the user provided
        # a lambda function that does not make sense.
        if attr_lambda(obj.GetAttribute(attr)):
          objects_true.append(obj)
      except(AttributeError) as e:
        raise powerfactorypy.PFAttributeError(obj,attr,e,self)
    return objects_true

  def get_from_folder(self,folder,obj_name=None,attr=None,attr_lambda=None,include_subfolders=False):
    """Get objects from a folder according to name and attribute.
    Arguments:
      folder: Path to folder or folder object
      obj_name: Name of the object(s), can contain wildcards ('*'). 
        If not specified, all objects are returned.
      attr: Name of attribute.
      attr_lambda: Lambda function with condition for the attribute.
        Example: attr_lambda = lambda x : x>110
        This checks if the value of attr is larger than 110.
    """
    folder = self.return_obj_if_path_is_provided(folder)
    int_subfolders = 0 if not include_subfolders else 1
    if obj_name:  
      objects = folder.GetChildren(1,obj_name,int_subfolders) # 1 means hidden objects are included
    else:
      objects = folder.GetContents(1,subfolders = int_subfolders)
    if attr:
      try:
        objects = self.get_by_attribute(objects,attr,attr_lambda)
      except(TypeError) as e:
        raise ValueError(f"Please specify method argument 'attr_lambda': {e}.")
    elif attr_lambda:
      raise ValueError("Please specify method argument \'attr\'.")
    return objects

  # ToDo functions get_from_global_library(path) and
  
  def get_path_of_object(self,obj):
    return PFStringManipuilation.format_full_path(str(obj))

  def get_attr(self,obj,attr):
    """
    Get the value of an attribute of an object.
    'obj' can be a path (string) or a Powerfactory object.
    Example:
     pfbi.get_attr(terminal_1,"systype")
    """
    obj = self.return_obj_if_path_is_provided(obj)
    attr_values = []
    try:
      if not isinstance(attr, list):
        return obj.GetAttribute(attr)
      else:
        for attribute in attr:
          attr_values.append(obj.GetAttribute(attribute))
        return attr_values
    except(AttributeError) as e:
      if not isinstance(attr, list):
        attribute = attr
      raise powerfactorypy.PFAttributeError(obj,attribute,e,self)

  def set_attr(self,obj,params):
    """
    Set the attribute(s) of an object. 
    obj: PowerFactory object or its path (string).
    params: dictionary {parameter1:value1, parameter2:value2,..}.
    """
    obj = self.return_obj_if_path_is_provided(obj)
    for attr, value in params.items():
      try:
        obj.SetAttribute(attr,value)
      except(TypeError) as e:
        print(obj)
        raise powerfactorypy.PFAttributeTypeError(obj,attr,e,self)
      except(AttributeError) as e:
        raise powerfactorypy.PFAttributeError(obj,attr,e,self)

  def set_attr_by_path(self,path_with_attr,value):
    """
    path_with_attr: path of object plus the attribute name
    Example:
      pfbi.set_attr_by_path(self,
        "Library\\Dynamic Models\\Linear_interpolation\\desc",["description"])
      Here 'desc' is the name of the attribute.  
    """
    head_tail = os_path.split(path_with_attr)
    self.set_attr(head_tail[0],{head_tail[1]:value})

  def create_by_path(self,path,overwrite=True):
    """
    Create an object by specifying its path including its class.
    If overwrite is true, objects with the same name will be overwritten.
    Example:
      pfbi.create_by_path(r"Library\Dynamic Models\dummy.BlkDef") 
    """
    try:
      folder_path, obj_name_incl_class = path.rsplit('\\',1)
    except(AttributeError):
      raise TypeError("The argument 'path' must be of type string.")
    folder = self.get_obj(folder_path)
    return self.create_in_folder(folder,obj_name_incl_class,overwrite=overwrite)
  
  def create_in_folder(self,folder,obj,overwrite=True):
    """
    Creates an obj inside a folder.
    If overwrite is true, objects with the same name will be overwritten.
    Example:
      pfbi.create_in_folder("Library\\Dynamic Models","dummy2.BlkDef")
    """
    folder = self.return_obj_if_path_is_provided(folder)
    try:
      obj_name, class_name = obj.split('.')
    except(AttributeError):
      raise TypeError("The argument 'obj' must be of type string.")
    if overwrite:
      self.delete_obj_from_folder(folder,obj,error_when_nonexistent=False)
    return folder.CreateObject(class_name, obj_name)

  def delete_obj(self,obj):
    obj = self.return_obj_if_path_is_provided(obj)
    obj.Delete()

  def delete_obj_from_folder(self,folder,obj_name=None,error_when_nonexistent=True,include_subfolders=False):
    """Deletes object(s) in a folder.
    Arguments:
      folder: PF folder (or its path)
      obj_name: Can contain wildcards ('*')
      error_when_nonexistent: Throw an error if the object does not exist.
      include_subfolders: Search also in subfolders.
    """
    folder = self.return_obj_if_path_is_provided(folder)
    objects = self.get_from_folder(folder,obj_name,include_subfolders=include_subfolders)  
    if objects:
      for obj in objects:
        obj.Delete()
    elif error_when_nonexistent:
      raise powerfactorypy.PFNonExistingObjectError(folder,obj_name)

  def copy_obj(self,obj_to_be_copied,new_obj_folder,new_name=None,overwrite=True):
    """Copy single object.
    Arguments:
      obj_to_be_copied: object or path
      new_obj_folder: folder object or its path
      new_name: If not specified, the original name is used
      overwrite: Overwrite if already exists in target folder
    """
    obj_to_be_copied = self.return_obj_if_path_is_provided(obj_to_be_copied)
    new_obj_folder = self.return_obj_if_path_is_provided(new_obj_folder)
    if overwrite:
      if not new_name:
        self.delete_obj_from_folder(new_obj_folder,obj_to_be_copied.GetAttribute("loc_name"),error_when_nonexistent=False)
      else:
        self.delete_obj_from_folder(new_obj_folder,new_name,error_when_nonexistent=False)
    if not new_name:
      new_obj = new_obj_folder.AddCopy(obj_to_be_copied)
    else:
      new_obj = new_obj_folder.AddCopy(obj_to_be_copied,new_name)
    return new_obj

  def copy_multiple_objects(self, objects_or_folder_to_be_copied, target_folder, overwrite=True):
    """Copy multiple objects or a whole folder.
    Arguments:
      objects_or_folder_to_be_copied: PF objects or a folder path
      target_folder: Target folder object or path
    """
    if isinstance(objects_or_folder_to_be_copied,str):
      objects_or_folder_to_be_copied = self.get_from_folder(objects_or_folder_to_be_copied)
    # If the content of a container object should be copied 
    elif (not isinstance(objects_or_folder_to_be_copied,list) and 
      True):
      objects_or_folder_to_be_copied = self.get_from_folder(objects_or_folder_to_be_copied)
    target_folder = self.return_obj_if_path_is_provided(target_folder)
    if overwrite:
      for obj in objects_or_folder_to_be_copied:
        obj_to_delete_list = target_folder.GetContents(obj.GetAttribute("loc_name") + ".*")
        if obj_to_delete_list:
          obj_to_delete_list[0].Delete()
    return target_folder.AddCopy(objects_or_folder_to_be_copied)

  def is_container(self,obj):
    """Checks whether a PF object is a container. It is assumed
    that an object is a container if it has the attribute "contents.
    """
    obj = self.return_obj_if_path_is_provided(obj)
    return obj.HasAttribute("contents")

  def return_obj_if_path_is_provided(self,obj):
    """Returns obj if obj is not a string. 
    Else it returns the PowerFactory object under obj (string path).
    """
    if not isinstance(obj, str):
      return obj  
    else:
      return self.get_obj(obj)

  def activate_study_case(self, path):
    """Activate study case under path.
    """
    study_case = self.get_obj(path)
    study_case.Activate()

class PFStringManipuilation:
  
  @staticmethod
  def replace_between_characters(char1,char2,replacement,string):
    new_string = ""
    is_between_chars = False
    for c in string:
      if c == char1:
        is_between_chars = True
      elif c == char2 and is_between_chars:
        is_between_chars = False
        new_string = new_string + replacement
      elif not is_between_chars:
        new_string = new_string + c
    return new_string   

  @staticmethod
  def delete_classes(path):
    return PFStringManipuilation.replace_between_characters('.','\\','\\',path)

  @staticmethod
  def format_full_path(path,pf_base_interface):
    """
    Takes the full path (including user and project) and returns the path 
    relative to the currently active project.
    Example:
      input path:  \\username.IntUser\\powerfactorypy_base.IntPrj\\Network Model.IntPrjfolder\\Network Data.IntPrjfolder\\Grid.ElmNet\\Terminal 1.ElmTerm
      output: Network Model\\Network Data\\Grid\\Terminal 1 
    """
    project_name = pf_base_interface.app.GetActiveProject().loc_name + '.IntPrj\\'
    path = path[path.find(project_name)+len(project_name):]
    return PFStringManipuilation.delete_classes(path)

def replace_in_string(string,translation):
  for old_string,replacement in translation.items():
    string = string.replace(old_string,replacement)
  return string
  
def get_R_and_X_from_RX_ratio(RX_ratio,Z_abs):
  """
  sqrt(R^2 + X^2) = Z_abs
  --> sqrt((X*RX_ratio)^2 + X^2) = Z_abs
  --> X = sqrt(Z_abs^2/(1 + RX_ratio^2))
  """
  X = math.sqrt(Z_abs**2/(1 + RX_ratio**2))
  R = X*RX_ratio
  return R,X

if __name__ == "__main__":
  print("ok")
  