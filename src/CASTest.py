from OpenOfficeTools import RunScript

OOTools  = RunScript("OOTools.py","Python","share")
CASUtil  = RunScript("CASUtil.py","Python","share")
CAS      = RunScript("CAS.py","Python","share")
  
def TestSetStarMath():
    CASUtil.ShowFormula("Fred1","p+q")
    CASUtil.ShowFormula("Fred2","m+z")
    CASUtil.ShowFormula("Fred3","yay")

def TestSetVarField():
    CASUtil.ShowVar("Var_1",15.0)
    #OOTools.MsgBox("test")
    CASUtil.ShowVar("Var_2",20.0)
    
def GetPath():
    import sys
    OOTools.MsgBox("\n".join(sys.path))

def CASCodeView():
    OOTools.MsgBox(CAS.GetCASCodeBlocks())
    OOTools.MsgBox(CAS.GetCASCodeBlocks(["CAS_O","CAS_L"]))
    
    
#def runScript(script, args=None, language="Python", location="share"):
#    # location options ('application','share')
#    # language options ('Basic','Python',etc.)
#    url = "vnd.sun.star.script:%(script)s?language=%(language)s&location=%(location)s" % {'script':script,'language':language,'location':location}
#    ctx = uno.getComponentContext()
#    scripts = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
#    scriptProvider = scripts.createScriptProvider(" ")
#    myScript = scriptProvider.getScript(url)
#    if args is None: args = ()
#    print myScript.invoke(args,(),())
    