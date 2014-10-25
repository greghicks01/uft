'@
'@
'@
'@
'@
'@
'@
'@

' Feature:
' Create a new database wrapper and connect with a connection string
'
' Test:
' Given the class is instantiated
' And we supply a valid connection sting
' Result is we connect to the database
'
' Scenario:
'
' Feature:
' Create a new Recordset and run a query
'
' Test:
' Given the first test passes
' And we supply a valid record set name and SQL statement
' The record set produces a result matching the SQL statement
'
' Scenario:
'

'@ Start: Test Harness
' connectionString = ""
' SQLString = ""

set testADO = newADODB



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