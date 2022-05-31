# %%
import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10')
import powerfactory as pf

sys.path.append(r'D:\User\seberlein\Code\powerfactorypy\src\powerfactorypy')
import powerfactorypy_base 
import importlib
importlib.reload(powerfactorypy_base)

# %%
app = pf.GetApplication()
pfi = powerfactorypy_base.PowerFactoryInterface(app)
pfi.app.Show()

# %%
pfi.app.ActivateProject(r'\seberlein\powerfactorypy\tests.IntPrj')

# %%
pfi.set_param(r"Library\Dynamic Models\Unit test frame\Input signals",{"sOutput":["aber"]})

# %%
composite_model = pfi.create(r"Network Model\Network Data\Grid\Unit test.ElmComp")

# %% Create an object
composite_model 
composite_frame = pfi.get_obj(r"Library\Dynamic Models\Unit test frame")
composite_model.SetAttribute("typ_id",composite_frame)

# s[s.rindex('-')+1:]
# %%
