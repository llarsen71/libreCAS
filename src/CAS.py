import uno
from OpenOfficeTools import RunScript, BlockNameHandlerFactory

masterNamer  = BlockNameHandlerFactory("Master")
CASCore = RunScript("CASCore.py","Python","share")
OOTools = RunScript("OOTools.py","Python","share")

def InsertPyCodeGlobalBlock():
    return CASCore.InsertPyCodeBlock("Global")

def InsertPyCodeLocalBlock():
    return CASCore.InsertPyCodeBlock("Local")

def InsertPyCodeBlock():
    dlg  = OOTools.GetDialog("Standard.PyCodeDialog")
    data = OOTools.GetDialogData("Standard.PyCodeDialog")
    data.clear()
    # Add a new undo_stack
    undo_stack = OOTools.UndoStackFactory()
    data["undo_stack"] = undo_stack

    label = dlg.getControl('BlockNameLabel')
    label.Text = CASCore.BlockNameFactory("Local")

    result = dlg.execute()
    if result == 1:
        # OK was pressed.
        CASCore.ModifyBlock(dlg)
    elif result == 0:
        # Cancel was pressed - if execute was used, delete block and items created.
        OOTools.ExecuteUndoStack(undo_stack)

def CASUpdate():
    blocks = masterNamer.GetNames(XSCRIPTCONTEXT)
    for block in blocks:
        try:
            CASCore.ExecuteCASBlock(block)
        except Exception, e:
            OOTools.MsgBox(str(e))
