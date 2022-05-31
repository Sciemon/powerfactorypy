import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2022 SP1\Python\3.10')
import powerfactory as pf

sys.path.append(r'..\src')
import powerfactorypy
import importlib
importlib.reload(powerfactorypy)


app = pf.GetApplication()
pfi = powerfactorypy.PowerFactoryInterface(app)
pfi.app.Show()

pfi.app.ActivateProject(r'\seberlein\SelfSyncToSimulink\SelfSyncFromSimulink_WithPWMAndScripts.IntPrj')

