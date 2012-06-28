from com.sun.star.awt import Rectangle
from com.sun.star.awt import WindowDescriptor

from com.sun.star.awt.WindowClass import MODALTOP, TOP
from com.sun.star.awt.VclWindowPeerAttribute import OK, OK_CANCEL, YES_NO, YES_NO_CANCEL, \
                          RETRY_CANCEL, DEF_OK, DEF_CANCEL, DEF_RETRY, DEF_YES, DEF_NO

from OpenOfficeTools import var_bin

import uno
from com.sun.star.beans import PropertyValue

# Show a message box with the UNO based toolkit
def MsgBox(MsgText="", MsgTitle="Message", MsgType="messbox", MsgButtons=OK):
	doc = XSCRIPTCONTEXT.getDocument()
	ParentWin = doc.CurrentController.Frame.ContainerWindow

	MsgType = MsgType.lower()

	#available msg types
	MsgTypes = ("messbox", "infobox", "errorbox", "warningbox", "querybox")

	if not ( MsgType in MsgTypes ): MsgType = "messbox"

	#describe window properties.
	aDescriptor                   = WindowDescriptor()
	aDescriptor.Type              = MODALTOP
	aDescriptor.WindowServiceName = MsgType
	aDescriptor.ParentIndex       = -1
	aDescriptor.Parent            = ParentWin
	#aDescriptor.Bounds           = Rectangle()
	aDescriptor.WindowAttributes  = MsgButtons

	tk = ParentWin.getToolkit()
	msgbox = tk.createWindow(aDescriptor)

	msgbox.setMessageText(str(MsgText))
	if MsgTitle: msgbox.setCaptionText(MsgTitle)

	return msgbox.execute()

def executeDispatch(command, args=(), context=None):
    if context is None:
        context = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    ctxt = uno.getComponentContext()
    sm = ctxt.ServiceManager
    dispatcher = sm.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctxt)
    dispatcher.executeDispatch(context, command, "", 0, makeProperties(args))

def makeProperties(dict_):
    if isinstance(dict_,tuple): return dict_
    if isinstance(dict_,list):  return tuple(dict_)
    properties = []
    for k,v in dict_.items():
        #p = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        p = PropertyValue()
        p.Name = k
        p.Value = v
        properties.append(p)
    return tuple(properties)

def inspect(object, name = "Inspect Object"):
    sm = uno.getComponentContext().ServiceManager
    inspector = sm.createInstance("org.openoffice.InstanceInspector")
    inspector.inspect(object,name)

def GetDialog(name, location="application"):
    sm = uno.getComponentContext().ServiceManager
    dp = sm.createInstance("com.sun.star.awt.DialogProvider")
    url = "vnd.sun.star.script:%s?location=%s" % (name,location)
    dlg = dp.createDialog(url)
    return dlg

def GetDialogData(name, location="application"):
    url = "vnd.sun.star.script:%s?location=%s" % (name,location)
    if url in var_bin:
        return var_bin[url]
    else:
        data = {}
        var_bin[url] = data
        return data

def UndoStackFactory():
    """
    Creates an undo stack. The undo stack should have the 'append' function that
    recieves a function that takes no parameters but executes
    an undo command.
    """
    return []

def ExecuteUndoStack(undo_stack):
    while len(undo_stack) > 0:
        try:
            command = undo_stack.pop()
            command()
        except Exception, e:
            MsgBox(e)
