' --------------------------------------------------------------------------
' Copyright(c) 2000-2006 Eurocenter DDC.
' No 65, Walukarama Road, Colombo 3, Sri Lanka
' All rights reserved.
'
' This software is the confidential and proprietary information of 
' Eurocenter DDC (Confidential Information). You shall not disclose such
' Confidential Information and shall use it only in accordance with the
' terms of the license agreement you entered into with Eurocenter.
'
' Description       : This script will add a new company into the Xis database when a new company is created in SuperOffice
'		      Also will update the name of the company, when the name is changed in SuperOffice	
' Author            : SKA
' Created Timestamp : 18/01/2007
' -------------------------------------------------------------------------- 


Dim HasContactNameChanged
HasContactNameChanged = False

'Triggers when a field of a company is changed
Sub OnCurrentContactFieldChanged(fieldname)
    	
    	'SOMessageBox "OnCurrentContactFieldChanged event : " + fieldname
    	If 	fieldname = "contact.name" Then
    		HasContactNameChanged = True
    	ElseIf fieldname = "contact.number2" Then
    		HasContactNameChanged = True
    	End If	
End Sub

'triggered when a new company is created
Sub OnCurrentContactCreated()
	 Dim strSQL 
    	 Dim rsContact 
         Dim UDef
                  
         Dim SOClient
	 Set SOClient = CreateObject("SuperOffice.Application")
	 Dim cur
	 Set cur = SOClient.CurrentContact
	 
	 'MsgBox "The current contact id is: " & cur.Identity & " The current name is: " & cur.Name, vbInformation + vbOkOnly, "SuperCOM"
         
         If cur.IsAvailable Then
         	
         	Dim strCon
         	strCon = GetConnectionString() 		
	
	 	Dim Conn 
	 	Set Conn = CreateConnection(strCon)
         	
         	'MsgBox "ExistsInXis : " & cur.UDef.ByName("ExistsInXis").Value
         	'MsgBox "cur.Category.Id : " & cur.Category.Id
         	'MsgBox "HasContactNameChanged : " & HasContactNameChanged
        	
        	If cur.UDef.ByName("ExistsInXis").Value = 0  Then 'And cur.Category.Id <> 3
        		
	            	'Check if contact exists in x|is
	                strSQL = "SELECT [FirmaID] FROM [Firma] WHERE [Socuid] = " & cur.Identity
	            	Set rsContact = GetRS(strSQL, Conn)
	            	
	            	'if contact doesn't exist create contact in xis
	            	If rsContact.EOF Then
	            		Dim companyName
	            		'replace ' character in company name
	            		companyName = Replace(cur.Name, "'", "''")
	                	strSQL = "INSERT INTO Firma(Firma, SOCuID, so_number2) VALUES('" & _
	                		companyName & "'," & _
	                		"'" & cur.Identity & "'," & _
	                		cur.Number2 & ")"
	                	                	
		                If ExecuteSQL(strSQL, Conn) = False Then
		                    rsContact.Close
		                    Set rsContact = Nothing
		                    MsgBox "Feil under oppdatering av xis!"
		                End If
	                
	            	End If
	            	'update SO with created in X|is state
	            	Set UDef = cur.UDef.ByName("ExistsInXis")
	            	UDef.Value = 1
	                'SOClient.CurrentContact.Save 'Does not work ?
	            	
	            	rsContact.Close
	            	Set rsContact = Nothing
	            	HasContactNameChanged = False 'reset state of name field
	            	
        	End If
        	CloseConnection(Conn)
        	
        End If

End Sub



'Triggers when a existing company is chnaged. 
Sub OnCurrentContactSaved()
	 
	 Dim strSQL 
    	 Dim rsContact 
         Dim UDef
                  
         Dim SOClient
	 Set SOClient = CreateObject("SuperOffice.Application")
	 Dim cur
	 Set cur = SOClient.CurrentContact
	 
	 'MsgBox "The current contact id is: " & cur.Identity & " The current name is: " & cur.Name, vbInformation + vbOkOnly, "SuperCOM"
         
         If cur.IsAvailable Then
         	
         	Dim strCon
         	strCon = GetConnectionString() 		
	
	 	Dim Conn 
	 	Set Conn = CreateConnection(strCon)
         	
         	'MsgBox "ExistsInXis : " & cur.UDef.ByName("ExistsInXis").Value
         	'MsgBox "cur.Category.Id : " & cur.Category.Id
         	'MsgBox "HasContactNameChanged : " & HasContactNameChanged
        	
        	  	
        	'One or more fields might have changed
        	'Name has changed, must be propageted to Xis
        	If cur.UDef.ByName("ExistsInXis").Value = 1 And HasContactNameChanged Then 'the name of the contact has changed
            		
            		HasContactNameChanged = False 'reset state of name field
            		Dim companyName
            		'replace ' character in company name
            		companyName = Replace(cur.Name, "'", "''")
            		strSQL = "UPDATE [Firma] SET [Firma] = '" & companyName & "', [so_number2]=" & cur.Number2 & "WHERE [Socuid] = " & cur.Identity
            
            
	            	If ExecuteSQL(strSQL, Conn) = False Then
	                	rsContact.Close
	                	Set rsContact = Nothing
	                	MsgBox "Feil under oppdatering av firmanavn i xis!"
	            	End If
        	End If
        	CloseConnection(Conn)
        	
        End If
	
End Sub

'Returns a forward readonly recordset
Function GetRS(strSQL, Conn)
		Dim oRS 
		'MsgBox "GetRS"	
		Set oRS	= CreateObject("ADODB.Recordset")
		'Prepare firehose RS for dataretrieval
		oRS.LockType =	1 	'adLockReadOnly
		oRS.CursorType = 0	'adOpenForwardOnly 
		on error resume next
		oRS.Open strSQL, Conn
		If ( Conn.Errors.Count > 0 ) Then
			on error goto 0
			MsgBox "Error in :" & strSQL
			set GetRS = nothing
			Exit Function
		End If
		on error goto 0
		Set GetRS = oRS
End Function

'Executes a crud SQL statement (Create, Update, Delete and Insert)
Function ExecuteSQL(strSQL, Conn)
		Dim oRS 
		Dim Cmd
		'MsgBox "ExecuteSQL"
		ExecuteSQL = false		
		
		set Cmd = CreateObject("ADODB.Command")
		set Cmd.ActiveConnection = Conn
		Cmd.CommandText = strSQL
		Cmd.CommandType = &H0001 'adCmdText
		on error resume next
		Cmd.Execute 

		If ( Conn.Errors.Count > 0 ) Then
			MsgBox "Error in :" & strSQL
		else
			ExecuteSQL = true
			
		End If
		on error goto 0
		set Cmd.ActiveConnection = nothing
		set Cmd = nothing
End Function

'Read the connection details of the Xis db from the registry
Function GetConnectionString()
	 
	 Dim objShell
	 Set objShell = CreateObject("WScript.Shell")
	 xtra_Provider = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Provider")
	 xtra_Username = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Username")
 	 xtra_Password = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Password")
 	 xtra_DBServer = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\DBServer")
 	 xtra_DBName = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\DBNAME")
	 xtra_Options = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Options")
	 	 
	 Dim strConnection
	 	 
         'xtra_Provider = "SQLOLEDB.1"
         'xtra_Username = "xtrawebuser"
         'xtra_Password = "xtrawebuser"
         'xtra_DBServer = "eccoldev04vm02\xtra_integration"
         'xtra_DBName = "Xis"
         'xtra_Options =  "Persist Security Info=True"
         strConnection = "Provider=" & xtra_Provider & ";" &_
				"User ID=" & xtra_Username & ";" &_
				"Password=" & xtra_Password & ";" &_
				"Data Source=" & xtra_DBServer & ";" &_
				"Initial Catalog=" & xtra_DBName  & ";" &_
				xtra_Options
				
	 'MsgBox "Connection : " & strConnection

	 GetConnectionString = trim(strConnection)
End Function

'Creates the db connection
Function CreateConnection(sConString)
	Dim Conn
	Set Conn = CreateObject("adodb.Connection")
	Conn.Open(sConString)
	Set CreateConnection = Conn
End Function

'Closes the db connection
Sub CloseConnection(byref Conn)
		'MsgBox "Close connection"
		If (IsObject(Conn)) Then
			If (not Conn is nothing) Then
				If ( Conn.State <> 0) Then
					Conn.Close()
				End If
			end if
		End If
End Sub


'Sub MakeSave(byref objSO)
'	If not (objSO is nothing) Then
'	    	result = MsgBox ("There have been no programmatically changes to your current company." & vbCrLf & "But, by pushing 'YES' this will save the changes on your current company, are you sure?", vbInformation + vbYesNo, "SuperCOM")
'	    	if result = vbYes then
'	        	MsgBox "MakeSave 1"
'	        	objSO.CurrentContact.Save
'	        	MsgBox "MakeSave 2"
'	    	else
'	        	MsgBox "The company change were not saved.", vbInformation + vbOkOnly, "SuperCOM"
'	    	end if
'	Else
'	    MsgBox "Unable to connect to database"
'	end if

'End Sub
