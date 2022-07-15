import pytest
import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10')
import powerfactory
sys.path.append(r'.\src')
#sys.path.append(r'.\src\powerfactorypy')
import powerfactorypy 


@pytest.fixture(scope='session')
def pf_app():
    return powerfactory.GetApplication()

@pytest.fixture
def pfbi(pf_app):
    # Return PFBaseInterface instance
    return powerfactorypy.PFBaseInterface(pf_app)   

@pytest.fixture
def activate_test_project(pfbi):
    pfbi.app.ActivateProject(r"\seberlein\powerfactorypy_base")

def test_get_obj(pfbi,activate_test_project):
    terminal_1 = pfbi.get_obj(r"Network Model\Network Data\Grid\Terminal MV 1")
    assert isinstance(terminal_1,powerfactory.DataObject)
    with pytest.raises(powerfactorypy.PFPathError):
        terminal_1 = pfbi.get_obj(r"Stretchwork Model\Stretchwork Data\Grid\Termalamala")
    with pytest.raises(powerfactorypy.PFPathError):
        terminal_1 = pfbi.get_obj(r"N")    
    with pytest.raises(TypeError):
        terminal_1 = pfbi.get_obj(terminal_1)

def test_get_obj_with_project_folder_argument(pfbi,activate_test_project):
    project_folder = pfbi.app.GetActiveProject()
    terminal_1 = pfbi.get_obj(r"Network Model\Network Data\Grid\Terminal MV 1",project_folder)
    assert isinstance(terminal_1,powerfactory.DataObject)

def test_set_attr(pfbi,activate_test_project):
    test_string_1 = "TestString1"
    test_string_2 = "TestString2"
    pfbi.set_attr(r"Library\Dynamic Models\Linear_interpolation",{"sTitle":test_string_1})
    pfbi.set_attr(r"Library\Dynamic Models\Linear_interpolation",{"sTitle":test_string_2,
        "desc":["dummy description"]})
    stitle = pfbi.get_attr(r"Library\Dynamic Models\Linear_interpolation","sTitle")
    assert stitle == test_string_2

def test_set_attr_exceptions(pfbi,activate_test_project):
    with pytest.raises(powerfactorypy.exceptions.PFAttributeTypeError):
        pfbi.set_attr(r"Library\Dynamic Models\Linear_interpolation",{"sTitle":"dummy",
        "desc":2}) # "desc" should be a list with one string item
    with pytest.raises(powerfactorypy.exceptions.PFAttributeError):
        pfbi.set_attr(r"Library\Dynamic Models\Linear_interpolation",{"sTie":"dummy",
        "desc":["dummy description"]}) # 'sTie' is not a valid attribute 
    with pytest.raises(powerfactorypy.exceptions.PFPathError):
        terminal_1 = pfbi.get_obj(r"Network Model\Network Data\Grid\Termalamala")

def test_set_attr_by_path(pfbi,activate_test_project):
    pfbi.set_attr_by_path(r"Library\Dynamic Models\Linear_interpolation\desc",["description"])
    with pytest.raises(powerfactorypy.exceptions.PFPathError):
        pfbi.set_attr_by_path(r"Stretchwork Model\Stretchwork Data\Grid\Termalamala",["description"])


def test_get_attr(pfbi,activate_test_project):
    terminal_1 = pfbi.get_obj(r"Network Model\Network Data\Grid\Terminal MV 1")
    systype = pfbi.get_attr(terminal_1,"systype")
    assert systype == 0
    with pytest.raises(powerfactorypy.exceptions.PFAttributeError):
        systype = pfbi.get_attr(terminal_1,"trixi")

def test_create_by_path(pfbi,activate_test_project):
    pfbi.create_by_path(r"Library\Dynamic Models\dummy.BlkDef")   
    with pytest.raises(powerfactorypy.exceptions.PFPathError):
        pfbi.create_by_path(r"ry\Dynamic Models\dummy.BlkDef")
    with pytest.raises(TypeError):
        pfbi.create_by_path(4)

def test_create_in_folder(pfbi,activate_test_project):
    pfbi.create_in_folder(r"Library\Dynamic Models","dummy2.BlkDef")
    with pytest.raises(TypeError):
        pfbi.create_in_folder(r"Library\Dynamic Models",2)

def test_get_from_folder(pfbi,activate_test_project):
    dsl_obj_angle = pfbi.get_from_folder(r"Network Model\Network Data\Grid\Voltage source ctrl",
        obj_name="Angle")
    assert len(dsl_obj_angle) == 1

    folder = r"Network Model\Network Data"
    terminals = pfbi.get_from_folder(folder, obj_name="*.ElmTerm", attr="uknom", 
        attr_lambda=lambda x : x==110, include_subfolders=True)
    assert len(terminals) == 2

    terminals = pfbi.get_from_folder(folder, obj_name="*.ElmTerm", attr="uknom", 
        attr_lambda=lambda x : x==110) # Does not search in subfolders.
    assert len(terminals) == 0

def test_get_from_folder_exceptions(pfbi,activate_test_project):
    folder = r"Network Model\Network Data"
    with pytest.raises(ValueError):
        terminals = pfbi.get_from_folder(folder, obj_name="Terminal*", attr="uknom", 
            include_subfolders=True)    

    with pytest.raises(powerfactorypy.exceptions.PFAttributeError):
        terminals = pfbi.get_from_folder(folder, obj_name="Terminal*", attr="wrong_attr", 
            attr_lambda=lambda x : x==110, include_subfolders=True)

def test_get_by_attribute(pfbi,activate_test_project):
    folder = r"Network Model\Network Data\Grid"
    all_terminals = pfbi.get_from_folder(folder, obj_name="*.ElmTerm")
    
    mv_terminals = pfbi.get_by_attribute(all_terminals,"uknom",lambda x:x > 100)
    assert len(mv_terminals) == 2

    with pytest.raises(powerfactorypy.exceptions.PFAttributeError):
        mv_terminals = pfbi.get_by_attribute(all_terminals,
            "wrong_attr",lambda x:x > 100)

def test_delete_obj_from_folder(pfbi,activate_test_project):
    folder = r"Library\Dynamic Models\TestDelete"
    pfbi.create_in_folder(folder,"dummy_to_be_deleted_1.BlkDef")
    pfbi.create_in_folder(folder,"dummy_to_be_deleted_2.BlkDef")
    pfbi.delete_obj_from_folder(folder,"dummy_to_be_deleted*")
    objects_in_folder = pfbi.get_from_folder(folder)
    assert len(objects_in_folder) == 0

    pfbi.create_in_folder(folder,"dummy_to_be_deleted_1.BlkDef")
    pfbi.create_in_folder(folder,"dummy_to_be_deleted_2.BlkDef")
    pfbi.delete_obj_from_folder(folder,"dummy_to_be_deleted_1.BlkDef")
    objects_in_folder = pfbi.get_from_folder(folder)
    assert len(objects_in_folder) == 1

    pfbi.create_in_folder(folder,"dummy_to_be_deleted_1.BlkDef")
    pfbi.create_in_folder(folder,"dummy_to_be_deleted_2.BlkDef")
    pfbi.delete_obj_from_folder(r"Library\Dynamic Models","dummy_to_be_deleted*",include_subfolders=True)
    objects_in_folder = pfbi.get_from_folder(folder)
    assert len(objects_in_folder) == 0

    with pytest.raises(powerfactorypy.exceptions.PFNonExistingObjectError):
        pfbi.delete_obj_from_folder(folder,"dummy_to_be_deleted_1.BlkDef",
            error_when_nonexistent=True)

def test_copy_obj(pfbi,activate_test_project):
    obj_to_be_copied = r"Library\Dynamic Models\Linear_interpolation"
    folder_copy_to = r"Library\Dynamic Models\TestCopy"
    pfbi.delete_obj_from_folder(folder_copy_to,"*",error_when_nonexistent=False)
    pfbi.copy_obj(obj_to_be_copied,folder_copy_to)
    pfbi.copy_obj(obj_to_be_copied,folder_copy_to,new_name="dummy_new_name")

def test_copy_multiple_objects(pfbi,activate_test_project):
    folder_copy_from = r"Library\Dynamic Models\TestDummyFolder"
    folder_copy_to = r"Library\Dynamic Models\TestCopyMultiple"
    pfbi.copy_multiple_objects(folder_copy_from,folder_copy_to)

    pfbi.delete_obj_from_folder(folder_copy_to,"*",error_when_nonexistent=False)
    objects_to_copy = pfbi.get_from_folder(folder_copy_from)
    pfbi.copy_multiple_objects(objects_to_copy,folder_copy_to)



if __name__ == "__main__":
    pytest.main()
