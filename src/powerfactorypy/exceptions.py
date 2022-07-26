import sys
sys.path.append(r".\src")
import powerfactorypy

class PFInterfaceError(Exception):
  """There should always be a base class (that inherits
  from 'Exception') for all custom errors/exceptions.
  """
  pass

class PFAttributeError(PFInterfaceError):
  """Attempt to access an invalid attribute of a PF object.
  """
  def __init__(self,obj,attr,msg_raised,pf_base_interface):
    object_str = powerfactorypy.PFStringManipuilation.format_full_path(str(obj),pf_base_interface)
    self.message = (f"The attribute '{attr}' of the object '{object_str}'"
      f" is unexpected: {msg_raised}.")
    super().__init__(self.message)

class PFAttributeTypeError(PFInterfaceError):
  """Attempt to set an invalid type for the attribute of a PF object.
  """
  def __init__(self,obj,attr,msg_raised,pf_base_interface):
    object_str = powerfactorypy.PFStringManipuilation.format_full_path(str(obj),pf_base_interface)
    self.message = (f"The attribute '{attr}' of the object '{object_str}' "
      f"is of unexpected type: {msg_raised}.")
    super().__init__(self.message)

class PFPathError(PFInterfaceError):
  """Attempt to access invalid path in PF database.
  """
  def __init__(self,path,project_folder):
    project_folder_str = powerfactorypy.PFStringManipuilation.delete_classes(str(project_folder))
    self.message = (f"The path '{path}' does not exist "
     f"within the project '{project_folder_str}'")
    super().__init__(self.message)

class PFNonExistingObjectError(PFInterfaceError):
  """Attempt to access PF object that does not exist.
  """
  def __init__(self,folder,obj):
    folder_str = powerfactorypy.PFStringManipuilation.delete_classes(str(folder))
    self.message = (f"The folder '{folder_str}' does not contain "
      f"any object named '{obj}'")
    super().__init__(self.message)

