'@   Copyright 2014
'@
'@   Gregory Hicks
'@ Software Engineer
'@
'@ Licensed under Creative Commons
'@
'@ You may use or extend this work under the following conditions:
'@ 1. You must include this copyright in your derived works.
'@ 2. The software is made available "As-Is". 
'@    No liabilities accepted for your use or updates applied
'@    No Warranties, implicit or implied apply to use.
'@
'
'
'	AJAX Synchronous Code Base (UFT)
'
' Code consists of various techniques to manage AJAX synchronisation
'
'  Until and While wait for a timeout period for something to come into or out scope 
'

' Feature:
' Global timer object
'
' Scenario:
'

class clsGblTimer
end class

'@Description User Defined Sync operation
'@
'@ Function waits upto a global timeout value for the
'@ test_object to come into existance
'@
'@ Accepts: Reference to the test object
'@ Returns: Boolean
Public Function UntilExist(ByRef test_object)

    ' Dictionary object loaded at the start of the script
    ' compares to cDctVariables MercTimer
    ' could be improved by making a class that removes most of the if statment code
    gDctVariables("MercTimer").Start

    do
	If gDctVariables("MercTimer").elapsedtime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
	Desktop.RunAnalog "Track3"
	
    Loop until test_object.Exist(1)
	
    UntilExist = test_object.Exist(0)
	
End Function

'@Description User Defined Sync operation
'@
'@ Function waits upto a global timeout value for the
'@ test_object to leave existance
'@
'@ Accepts: Reference to the test object
'@ Returns: Boolean
Public Function WhileExist(ByRef test_object)
    ' 	Wait till global timeout for object to leave exist
    ' returns Boolean
    gDctVariables("MercTimer").Start
   	do
		Desktop.RunAnalog "Track3"
		If gDctVariables("MercTimer").elapsedtime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
	
	Loop while test_object.Exist(1)
	
	WhileExist = not test_object.Exist(1)
	
End Function

'@Description User Defined Sync operation
'@
'@ Function waits upto a global timeout value for the
'@ test_object to come into existance
'@
'@ Accepts: Reference to the test object
'@          ROProperty type, eg micClass or Visible
'@          vValue the value we expect to find when this succeed
'@ Returns: Boolean
Public Function UntilPropertyIs(ByRef test_object , sRoProperty , vValue )
    ' 	Wait till global timeout for object to leave exist
    ' returns Boolean
    UntilPropertyIs = true
    gDctVariables("MercTimer").Start
   	do
		Desktop.RunAnalog "Track3"
		If gDctVariables("MercTimer").elapsedtime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
	
	Loop until test_object.getROProperty( sRoProperty ) = vValue
	
	UntilPropertyIs = test_object.getROProperty( sRoProperty ) = vValue
	
End Function

'@Description User Defined Sync operation
'@
'@ Function waits upto a global timeout value for the
'@ test_object to come into existance
'@
'@ Accepts: Reference to the test object
'@          ROProperty type, eg micClass or Visible
'@          vValue the value we expect to find when this succeed
'@ Returns: Boolean
Public Function WhilePropertyIs(ByRef test_object , sRoProperty , vValue )
    ' 	Wait till global timeout for object to leave exist
    ' returns Boolean
    gObjTimer.Start
   	do
		Desktop.RunAnalog "Track3"
		If gDctVariables("MercTimer").elapsedtime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
	
	Loop while test_object.getROProperty( sRoProperty ) = vValue
	
	WhilePropertyIs = not test_object.getROProperty( sRoProperty ) = vValue
	
End Function

'@Description User Defined Sync/Child obj operation
'@
'@ Function wait on objects while a property is value
'@
'@ Accepts: Reference to the test object
'@          ROProperty type, eg micClass or Visible
'@          vValue the value we expect to find when this succeed
'@ Returns: Boolean
Function WhileChObjPropertyIs(byref obj , ROProperty , vValue)
	Dim ch , x
    ' 	Wait till global timeout for child object
    ' returns Nothing or object
	
	gDctVariables("MercTimer").Start
	
	do 	
		WhileChObjPropertyIs = Nothing
		
		set ch = obj.ChildObjects
		For x = 0 To ch.count - 1
			if UCase(ch(x).getROProperty(ROProperty)) = UCase(vValue) then set WhileChObjPropertyIs = ch(x) : exit for
		Next
		Desktop.RunAnalog "Track3"
		
		If WhileChObjPropertyIs is Nothing Then exit do
		
		If gDctVariables("MercTimer").ElapsedTime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
		
	Loop 
	
End Function

'@Description User Defined Sync/Child obj operation
'@
'@ Function waits for a child object property to get to an expected state or times out
'@ 
'@
'@ Accepts: Reference to the test object that contains a child
'@          ROProperty type, eg micClass or Visible
'@          vValue the value we expect to find when this succeed
'@ Returns: Boolean
Function UntilChObjPropertyIs(byref obj , ROProperty , vValue)
	Dim ch , x
    ' 	Wait till global timeout for child object
    ' returns Nothing or object
	
	gDctVariables("MercTimer").Start
	
	do 	
		set UntilChObjPropertyIs = Nothing
		set ch = obj.ChildObjects
		For x = 0 To ch.count - 1
			if UCase(ch(x).getROProperty(ROProperty)) = UCase(vValue) then set UntilChObjPropertyIs = ch(x) : exit for
		Next
		Desktop.RunAnalog "Track3"
		
		If not UntilChObjPropertyIs is Nothing Then exit do
		
		If gDctVariables("MercTimer").ElapsedTime > gDctVariables("Global TimeOut") Then gDctVariables("MercTimer").Reset : exit do
		
	Loop 
	
End Function


'@Description User Defined Sync/Child obj operation'@
'@ Function clicks on link usnig test obect name property
'@
'@ Accepts: Reference to the test object, could have been just the link name
'@ Returns: Boolean
Function WebClickLink( byref test_obj )
	Dim lCh , oDescr , lclIdx
	
	Set oDescr = Description.Create
	oDescr("micClass").Value = "Link"
	oDescr("Visible").Value = true
	
	Set lCh = Browser("creationtime:=0").Page("name:=.*").ChildObjects(oDescr)
	
	If lCh.Count > 0 Then
		For lclIdx  = 0 To lCh.Count
			If lCh(lclIdx).Name = test_Obj.Name Then
				lCh(lclIdx).Click
			End If			
		Next
	End If
	
End Function

RegisterUserFunc "Link", "UntilExist", "UntilExist"
RegisterUserFunc "Link", "WhileExist", "WhileExist"
RegisterUserFunc "Link", "WebClickLink", "WebClickLink"
RegisterUserFunc "Link", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "Link", "WhilePropertyIs", "WhilePropertyIs"

RegisterUserFunc "WebEdit", "UntilExist", "UntilExist"
RegisterUserFunc "WebEdit", "WhileExist", "WhileExist"
RegisterUserFunc "WebEdit", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "WebEdit", "WhilePropertyIs", "WhilePropertyIs"

RegisterUserFunc "WebButton", "UntilExist", "UntilExist"
RegisterUserFunc "WebButton", "WhileExist", "WhileExist"
RegisterUserFunc "WebButton", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "WebButton", "WhilePropertyIs", "WhilePropertyIs"

RegisterUserFunc "WebTable", "UntilExist", "UntilExist"
RegisterUserFunc "WebTable", "WhileExist", "WhileExist" 
RegisterUserFunc "WebTable", "UntilChObjPropertyIs", "UntilChObjPropertyIs"
RegisterUserFunc "WebTable", "WhileChObjPropertyIs", "WhileChObjPropertyIs"
RegisterUserFunc "WebTable", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "WebTable", "WhilePropertyIs", "WhilePropertyIs"  

RegisterUserFunc "WebElement", "UntilExist", "UntilExist"
RegisterUserFunc "WebElement", "WhileExist", "WhileExist"
RegisterUserFunc "WebElement", "ChildObjectWait", "ChildObjectWait"
RegisterUserFunc "WebElement", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "WebElement", "WhilePropertyIs", "WhilePropertyIs"

RegisterUserFunc "WinObject", "UntilExist", "UntilExist"
RegisterUserFunc "WinObject", "WhileExist", "WhileExist"
RegisterUserFunc "WinObject", "UntilChObjPropertyIs", "UntilChObjPropertyIs"
RegisterUserFunc "WinObject", "WhileChObjPropertyIs", "WhileChObjPropertyIs"
RegisterUserFunc "WinObject", "UntilPropertyIs", "UntilPropertyIs"
RegisterUserFunc "WinObject", "WhilePropertyIs", "WhilePropertyIs"

RegisterUserFunc "Browser", "WhileExist", "WhileExist"
RegisterUserFunc "Browser", "UntilExist", "UntilExist"

RegisterUserFunc "Page", "UntilExist", "UntilExist"
RegisterUserFunc "Page", "WhileExist", "WhileExist"