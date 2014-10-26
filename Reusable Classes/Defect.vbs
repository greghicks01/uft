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
'@'
'
'
'
'
'
'


function  RaiseDefect( strSummary )

   Exit function
    Dim QCConnection

    'AssignedTo     The name of the user to whom the defect is assigned.  
    'Attachments    The Attachment factory for the object.  
    'DetectedBy     The name of the user who detected the defect.   
    'ID             The item ID.  
    'IsLocked       Checks if object is locked for editing.  
    'Modified       Checks if the item has been modified since last refresh or post operation. If true, the field properties on the server side are not up to date.  
    'Priority       The defect priority.  
    'Status         The defect status.  
    'Summary        A short description of the defect.  
    'Virtual 


    Set QCConnection = QCUtil.QCConnection
    
    'Get the IBugFactory 
    Set BugFactory = QCConnection.BugFactory 
    
    'Add a new, empty defect 
    Set Bug = BugFactory.AddItem (Nothing) 
    
    'Enter values for required fields 
    Bug.Project = QCConnection.ProjectName
    Bug.Status = "New" 
    Bug.Summary = strSummary
    Bug.DetectedBy = QCConnection.UserName ' user that must exist in the database's user list 
    
    'Post the bug to the database ( commit ) 
    Bug.Post 

    set QCConnection = nothing

end function