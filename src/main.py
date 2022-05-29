import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10')
import powerfactory as pf
import pandas as pd
from os import path as os_path
from os import remove
import collections 

class PowerFactoryInterface:

  def __init__(self,app, export_path=""):  
    
    self.app = app
    self.export_path = export_path
    self.csv_export_file_name = "SimulationResults.csv"    
    self.active_graphics_page = ""
    self.active_plot = ""
    self.results = "All calculations.ElmRes"
  
  def get_obj(self,path):
    """
    Returns the PowerFactory object under path.
    The path must be specified relativ to the project folder.
    Example:
    *ToDo
    """
    if path[0] == "\\":
      path = path[1:] 
    splitted_path = path.split("\\")
    folder_obj = [self.app.GetActiveProject()]
    for folder_name in splitted_path:
      folder_obj = folder_obj[0].GetContents(folder_name + '.*')
    return folder_obj[0]

  def activate_study_case(self, path):
    if not path.startswith("Study Cases\\"):
      path = "Study Cases\\" + path
    study_case = self.get_obj(path)
    study_case.Activate()

  def set_active_graphics_page(self,obj):
    """Accepts a graphics page object or its name."""
    if isinstance(obj, str):
      grb = self.app.GetGraphicsBoard()
      self.active_graphics_page = grb.GetPage(obj,1,"GrpPage")
    else:
      self.active_graphics_page = obj

  def set_active_plot(self,name,graphics_page=None):
    """Accepts the name of the plot."""
    if not graphics_page==None:
      self.set_active_graphics_page(graphics_page)
    self.active_plot = self.active_graphics_page.GetOrInsertCurvePlot(name)

  def return_obj_if_string_path_is_specified(self,obj):
    """
    Returns the input object if  it is not a string. 
    Else it returns the PowerFactory object from the path.
    """
    if not isinstance(obj, str):
      return obj  
    else:
      return self.get_obj(obj)

  def set_param(self,obj,params):
    """
    obj: PowerFactory object or its path.
    params: dictionary {parameter:value,..}.
    """
    obj = self.return_obj_if_string_path_is_specified(obj)
    for param, value in params.items():
      obj.SetAttribute(param,value)

  def set_param_by_path(self,param_path,values):
    """
    param_path: path of object plus the parameter/attribute name
      - example: *ToDo
    values: list of values
      * ToDo: does the list make sense? 
    """
    head_tail = os_path.split(param_path)
    obj = self.get_obj(head_tail[0])
    if not isinstance(values, collections.abc.Iterable):
      values =  [values]
    for value in values:
      obj.SetAttribute(head_tail[1],value)

  def add_results_variable(self,obj,variables):
    """
    Adds variables of the object to the PowerFactory results object (ElmRes)
    in self.results. *ToDo: should it really be a string?
    obj: PowerFactory object or its path
    """
    results_storage_obj = self.app.GetFromStudyCase(self.results)
    obj = self.return_obj_if_string_path_is_specified(obj)
    if isinstance(variables, str):
      variables = [variables]
    for var in variables:
      results_storage_obj.AddVariable(obj,var)
    results_storage_obj.Load()

  def plot_monitored_variables(self,obj,variables,**kwargs):
    """
    obj: PowerFactory object or its path
    variable: string or list of variable names 
    graphics_page: Name of graphics page
    plot: Name of plot
    """
    if "graphics_page" in kwargs:
      self.set_active_graphics_page(kwargs['graphics_page'])
    if "plot" in kwargs:
      self.set_active_plot(kwargs['plot'])
    data_series = self.active_plot.GetDataSeries()
    obj = self.return_obj_if_string_path_is_specified(obj)
    if isinstance(variables, str):
     variables = [variables]
    for var in variables:
      data_series.AddCurve(obj,var)
      self.set_curve_attributes(data_series,**kwargs)
    self.active_graphics_page.Show()
     
  def plot(self,obj,variables,graphics_page=None,plot=None,**kwargs):
    """
    Plots the variables of 'obj' to the currently active plot.
    Includes adding the variables to the results object.
    The active plot can be set with the conditional arguments.
    """
    obj = self.return_obj_if_string_path_is_specified(obj)
    self.add_results_variable(obj,variables)
    self.plot_monitored_variables(obj,variables,**kwargs) 
  
  @staticmethod
  def set_curve_attributes(data_series,**kwargs):
    if  "linestyle" in kwargs:
      list_curveTableAttr = data_series.GetAttribute("curveTableLineStyle")
      list_curveTableAttr[-1] = kwargs['linestyle']
      data_series.SetAttribute("curveTableLineStyle",list_curveTableAttr)
    if "linewidth" in kwargs:
      list_curveTableAttr = data_series.GetAttribute("curveTableLineWidth")
      list_curveTableAttr[-1] = kwargs['linewidth']
      data_series.SetAttribute("curveTableLineWidth",list_curveTableAttr)
    if "color" in kwargs:
      list_curveTableAttr = data_series.GetAttribute("curveTableColor")
      list_curveTableAttr[-1] = kwargs['color']
      data_series.SetAttribute("curveTableColor",list_curveTableAttr)
    # The label must be handled differently because PF returns an empty list
    # if there haven't been any labels specified yet for any of the curves.
    if "label" in kwargs:
      list_curveTableAttr = data_series.GetAttribute("curveTableLabel")
      if list_curveTableAttr:
        list_curveTableAttr[-1] = kwargs['label']
      else:
        list_curveTableAttr = [kwargs['label']]
      data_series.SetAttribute("curveTableLabel",list_curveTableAttr)

  def autoscale(self):
    self.active_graphics_page.DoAutoScale()

  def clear_all_graphics_pages(self):
    """
    Deletes all graphics pages from the graphics board of 
    the active study case. 
    """
    grb = self.app.GetGraphicsBoard()
    graphics = grb.GetContents()
    for graphic in graphics:
      if graphic.GetClassName() == "GrpPage":    
        graphic.RemovePage()
        
  def clear_curves_from_all_plots(self): 
    """
    Clears data (i.e. curves) from all plots of the active study case.
    """     
    grb = self.app.GetGraphicsBoard()
    graphics = grb.GetContents()
    for graphic in graphics:
      if graphic.GetClassName() == "GrpPage":    
        for child in graphic.GetContents(): 
          if child.GetClassName() == "PltLinebarplot":
            data_series =child.GetDataSeries()
            data_series.ClearCurves()
                
  def export_to_csv(self):
    """
    Exports the simulation results in self.results to csv.
    *ToDo better explanation ftarget file path
    """
    comRes = self.app.GetFromStudyCase("ComRes")
    comRes.pResult = self.app.GetFromStudyCase(self.results)
    comRes.iopt_exp = 6 # to export as csv
    comRes.f_name = self.export_path + "\\" + self.csv_export_file_name
    comRes.iopt_sep = 1 # to use the system seperator
    comRes.iopt_honly = 0 # to export data and not only the header
    comRes.iopt_csel = 0 # export all variables 
    comRes.iopt_locn = 3 # column header includes path
    comRes.ciopt_head = 1 # full variable name
    comRes.Execute()
    self.format_csv()
  
  def format_full_path(self,path):
    """
    Takes the full path and returns the path relative to the currently active project.
    Example:
      input:  Network Data.IntPrjfolder\Grid.ElmNet\With Selflim.ElmComp\Control.ElmDsl.2
      output: Network Data\Grid\With Selflim\Control  
    """
    project_name = self.app.GetActiveProject().loc_name + '.IntPrj\\'
    path = path[path.find(project_name)+len(project_name):]
    path = PowerFactoryInterface.delete_chars_between_dot_and_slash(path)
    return path
    
  @staticmethod  
  def delete_chars_between_dot_and_slash(string):
    """
    Deletes all characters between '.' and ''.
    Example:
      input:  "User.IntUser\\Project name.IntPrj\\Network Data.IntPrjfolder\\Grid.ElmNet\\With Selflim.ElmComp\\Control.ElmDsl.2"
      output: "Network Data\\Grid\With Selflim\\Control"
    Characters after the last '.' are also deleted.
    Note that '.' are deleted, wheras '\\' are kept.
    """
    is_between_dot_and_slash = False
    string_new = ""
    for c in string:
        if c == '.':
          is_between_dot_and_slash = True
        elif c == '\\':
          is_between_dot_and_slash = False
          string_new = string_new + c
        elif not is_between_dot_and_slash:
          string_new = string_new + c  
    return string_new
          
  def format_csv(self):
    """
    Formats the csv file as exported from PowerFactory.
    PF uses two columns at the top for object and variabel name.
    This is reduced to one column that contains all the information.
    The first time column is named 'Time'. 
    """
    csv_path = self.export_path + '\\' + self.csv_export_file_name
    with open(csv_path) as file:
      csv_file = pd.read_csv(file) 
      new_headers = []
      for header in csv_file: 
        variable = csv_file[header][0]
        new_header = self.format_full_path(header)
        new_headers.append(new_header + '.' + variable)
      new_headers[0] = 'Time'
      csv_file.columns = new_headers
      csv_file.drop(0,axis=0,inplace=True) # delete first column containing indexes
      csv_file.to_csv(csv_path+".csv",index=False)
    remove(csv_path)
  
  def initialize_dyn_sim(self,param=None):
    """
    Initialize time domain simulation.
    Parameters for 'ComInc' command object can be specified in 'param' dictionary.
    """
    cominc = self.app.GetFromStudyCase("ComInc")
    if param is not None:
      self.set_param(cominc,param)
    cominc.Execute()

  def sim(self,param=None):
    """
    Perform dynamic simulation.
    Parameters for 'ComSim' command object can be specified in 'param' dictionary.
    """
    comsim = self.app.GetFromStudyCase("ComSim")
    if param != None:
      self.set_param(comsim,param)
    comsim.Execute()

  def initialize_and_sim(self):
    """Initialize and perform time domain simulation."""
    self.initialize_dyn_sim()
    self.sim()

    
