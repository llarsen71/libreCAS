Sub InsertPyCodeBlockLocal()
	InsertPyCodeBlock("Local")
End Sub

Sub InsertPyCodeBlock(typ As String)
	dim Args(5) as new com.sun.star.beans.PropertyValue
	dim dummy()
	dim arg(1)
	doc = ThisComponent
	
	namer = GetScript("CASCore.py","BlockNameFactory")
	arg(0) = typ
	name1 = namer.invoke(arg, dummy, dummy)
	usermodel = doc.createInstance("com.sun.star.text.FieldMaster.User")
	usermodel.Name  = name1
	usermodel.Content = "# Insert Python Code Here"
	
	disp = createUnoService("com.sun.star.frame.DispatchHelper")
    
    Args(0).name  = "Type"
    Args(0).value = 16
    Args(1).name  = "SubType"
    Args(1).value = 2
    Args(2).name  = "Name"
    Args(2).value = name1
    Args(3).name  = "Format"
    Args(3).value = 0
    Args(4).name  = "Separator"
    Args(4).value = " "
    
    frame = doc.CurrentController.Frame
    fnDispatch = disp.executeDispatch(frame, ".uno:InsertField", "", 0, Args())
End Sub

Function GetScript(pkg As String, func As String) As Object
    spf      = createUnoService("com.sun.star.script.provider.MasterScriptProviderFactory")
    scripts  = spf.createScriptProvider(" ")
    GetScript = scripts.getScript("vnd.sun.star.script:" & pkg & "$" & func & "?language=Python&location=share")
End Function