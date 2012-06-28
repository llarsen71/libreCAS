#import uno
from OpenOfficeTools import RunScript
#from copy import copy

OOTools = RunScript("OOTools.py","Python","share")
CASCore = RunScript("CASCore.py","Python","share")
CAS     = RunScript("CAS.py","Python","share")
#basic   = RunScript("Standard.CAS","Basic","application")

def dlgPyBlock_ExecuteBtn(event):
    try:
        CASCore.ReloadOpenOfficeTools()
        dlg = event.Source.Context
        CASCore.ModifyBlock(dlg)
    except Exception, e:
        OOTools.MsgBox(e)

def dlgPyBlock_ScopeField(event):
    """
    Set the name of the new block.
    """
    label = event.Source.Context.getControl('BlockNameLabel')
    if event.Selected == 0:
        label.Text = CASCore.BlockNameFactory("Local")
    elif event.Selected == 1:
        label.Text = CASCore.BlockNameFactory("Global")

def dlgPyBlock_Next(event):
    # Get an ordered list and find where this fits
    pass

def dlgPyBlock_Prev(event):
    # Get an ordered list and find where this fits
    pass