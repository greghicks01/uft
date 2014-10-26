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
'			ADODB Manager Class
'
' Feature:
' Create a new database wrapper and connect with a connection string
'
' Scenario:
' Given the class is instantiated
' And we supply a valid connection sting
' Result is we connect to the database
'
'
' Feature:
' Create a new Recordset and run a query
'
' Scenario:
' Given the first test passes
' And we supply a valid record set name and SQL statement
' The record set produces a result matching the SQL statement
'
'

'@ Start: Test Harness
' connectionString = ""
' SQLString = ""

set testADO = newADODB

testADO.DBConnect ""
set rs1 = testADO.executeSQL("Set 1" , "Select * from t")
set rs2 = testADO.executeSQL("Set 2" , "Update t set f = <value> where f = 'value'")

do until rs1.eof
	'process rs 1
loop

do until rs2.eof
	'process rs 2
loop

'@ End: Test Harness

function newADODB
	newADODB = new clsADODB
end function

class clsADODB

	private objConnection, objRecordSets

	sub class_initialize
		on error resume next
			Set objConnection = CreateObject("ADODB.Connection")
			if err.number > 0 then Err.raise vbError + 1
		on error goto 0
		
		set objRecordSets = CreateObject("Scripting.Dictionary")
		
	end sub
	
	sub class_terminate : catch
		DBClose
		set objRecordSets = nothing
		set objConnection = nothing
	end sub
	
	private sub catch
		if err.number = 0 then exit sub
	end sub
	
	sub DBconnect ( conString )
		 objConnection.open conString
	end sub
	
	sub DBClose
		on error resume next
		for each rs in objRecordSets
			rs.close
		next
		objConnection.Close
		on error goto 0
		err.clear
	end sub
	
	private sub newRecSet ( stringRecordSetID )
		if stringRecordSetID = "" then Err.Raise vbErrorNumber + 2 : exit sub
		if objRecordSets.exist(stringRecordSetID) then exit sub
		
		objRecordSets.Add stringRecordSetID , CreateObject("ADODB.Recordset")
		
	end sub
	
	function executeSQL ( RecordSetID , SQLStatement )
		newRecSet ( RecordSetID )
		set executeSQL = objRecordSets(RecordSetID).open ( SQLStatement , objConnection )
	end function
	
end class