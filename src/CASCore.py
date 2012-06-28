import uno
import re
from com.sun.star.beans import PropertyValue
from OpenOfficeTools import RunScript, StackRepositoryFactory, BlockNameHandlerFactory, Cursor, CreateCodeBlockFactory
from copy import copy

OOTools = RunScript("OOTools.py","Python","share")
#basic   = RunScript("Standard.CAS","Basic","application")
stack_repo   = StackRepositoryFactory()

def CodeBlockFactory():
    return CreateCodeBlockFactory(XSCRIPTCONTEXT)

def InsertPyCodeBlock(typ, code=None):
    undo = []
    return CodeBlockFactory().CreateCodeBlock(typ, code, undo)

def BlockNameFactory(typ="Local", dummy=None):
    """
    type = Global, Local
    """
    return CodeBlockFactory().GetNewBlockName(typ)

def GetCASBlockCode(block_name):
    """
    List of CAS block names to extract code from.
    """
    undo = []
    return CodeBlockFactory().GetCodeBlock(block_name, undo).GetCode()

def ExecuteCASBlock(block, undo_stack=None):
    if undo_stack is None: undo_stack = []
    code_block = CodeBlockFactory().GetCodeBlock(name, undo_stack)
    code_block.Execute()

def ModifyBlock(dlg):
    cf  = dlg.getControl('PyCodeField')
    sf  = dlg.getControl('ScopeField')
    bl  = dlg.getControl('BlockNameLabel')
    hb  = dlg.getControl('HideCB')

    # Keep an undo stack for the dialog in case a cancel is called.
    undo = OOTools.GetDialogData("Standard.PyCodeDialog")['undo_stack']

    # Check the sf field to see if this is a new or an already defined block.
    # TODO: Come up with a better way to mark whether the current field is created.
    if not sf.getModel().Enabled:
        name = bl.Text
        code_block = CodeBlockFactory().GetCodeBlock(name, undo)
        code_block.SetCode(cf.Text)
    else:
        code_block = CodeBlockFactory().CreateCodeBlock(sf.Text, cf.Text, undo)
        name    = code_block.GetBlockName()
        bl.Text = name
        # Since we have added a block, we disable to code type field.
        sf.getModel().Enabled = False

    if hb.State == 0:
        code_block.Show()
    else:
        code_block.Hide()

    code_block.Execute()

def ReloadOpenOfficeTools():
    import OpenOfficeTools
    reload(OpenOfficeTools)