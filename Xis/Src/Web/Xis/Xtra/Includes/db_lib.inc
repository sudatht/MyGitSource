<%
	const XIS = "XIS"
	const EFP = "EFP"
	
	'Prefixes used:
	'Render - Prefix for HTML output functions 
	'Get  - Prefix for Get object functions

	' ********************************************
	' Name: GetConnectionstring
	' Purpose: Returns the specified connectionstring
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] strSource - Containing either "XIS" or "EFP"
	'   [in] strParams - Additional parameters
	'   [Returns] The connection string for the specified source
	' Notes: The second parameter is used to specify additional
	' options in the connection string
	' ********************************************
	function GetConnectionstring(strSource, strParams)
		dim strConnection
		if (strSource = CON_XIS) then
			strConnection = Session("xtra_ConnectionString") & _
			"uid=" & Session("xtra_RuntimeUserName") & _
			";pwd=" & Session("xtra_RuntimePassword") & ";" & strParams
		elseif (strSource = CON_EFP) then
			strConnection = application("IMPconnection") & strParams
		end if

		if (len(trim(strConnection))=0) then
			if session("debug")="true"	then
				' Print all error messages from connection
				Response.write "<p class='warning'>Kunne ikke bygge oppkoblingsstreng!" & "</p>"
			end if
			GetConnectionstring = ""
		else
			GetConnectionstring = trim(strConnection)
		end if
	End function


	' ********************************************
	' Name: RenderSqlErrors
	' Purpose: Renders all the errors as HTML
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] oCon - the connection object containing all the errors
	'   [in] sSource - A string containing additional information 
	'   regarding the context of the error: SQL, Connectionstring etc.
	' Notes: Only displays details about the error if session("debug") is set to true
	' ********************************************
	Sub RenderSqlErrors(oCon, sSource)
		dim NOFErrors
		dim iCurrentError

		if session("debug")="true"	then
			' Print all error messages from connection
			Response.write "<p class='warning'>F&oslash;lgende feil oppstod:" & "</p>"
			Response.write "<p><strong>Kilde:</strong><br>" & sSource & "</p>"
			NOFErrors = oCon.errors.count
			iCurrentError = 0
			while (iCurrentError < NOFErrors)
				Response.write "<p><strong>Description:</strong><br>" & oCon.errors(iCurrentError).description & "</p>"
				iCurrentError = iCurrentError + 1
			wend
		else
			Response.write "<p class='warning'>En uopprettelig feil oppstod. Vennligt kontakt IT-ansvarlig." & "</p>"
		end if
	End Sub

	' ********************************************
	' Name: CloseConnection
	' Purpose: Closes the specified connection
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] oCon - the connection object to use close 
	' Notes: -
	' ********************************************
	sub CloseConnection(byref oCon)
		If (IsObject(oCon)) Then
			If (not oCon is nothing) Then
				If ( oCon.State <> 0) Then
					oCon.Close()
				End If
			end if
		End If
	end sub


	' ********************************************
	' Name: GetFirehoseRS
	' Purpose: Returns a forward readonly recordset (firehose)
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sSQL - Statement to execute
	'   [in] oCon - the Open connection to use
	'   [out] An adodb.recordset object
	' Notes: Renders errors and returns nothing if Query failes
	' ********************************************
	Function GetFirehoseRS(sSQL, oCon)
		Dim oRS 
		
		On Error Resume Next
		
		If Len( sSQL ) = 0 Then
			Set getRecordset = Nothing
			Exit Function
		End If

		if session("debug")="true"	then
			call Response.Write( "<p>" & sSQL & "</p>") 
		end if

		Set oRS	= server.createobject("adodb.RecordSet")

		'Prepare firehose RS for dataretrieval
		oRS.CursorLocation =	0	'adOpenForwardOnly 
		oRS.LockType =			1	'adLockReadOnly
		oRS.ActiveConnection =	oCon

		oRS.Open sSQL, oCon

		If ( oCon.Errors.Count > 0 ) Then
			call RenderSqlErrors(oCon,sSQL)
			CloseConnection(oCon)
			On Error Goto 0
			set getFirehoseRS = nothing
			Exit Function
		End If
		Set getFirehoseRS = oRS
		On Error Goto 0
	End Function


	' ********************************************
	' Name: GetDynamicRS
	' Purpose: Returns a dynamic readonly recordset
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sSQL - Statement to execute
	'   [in] oCon - the Open connection to use
	'   [out] An adodb.recordset object
	' Notes: Renders errors and returns nothing if Query failes
	' Being a dynamic recordset, the recordcount 
	'property on the returned recordset is supported
	' ********************************************
	function GetDynamicRS(sSQL, oCon)
		Dim oRS 
		
		On Error Resume Next
		
		If Len( sSQL ) = 0 Then
			Set GetDynamicRS = Nothing
			Exit Function
		End If

		if session("debug")="true"	then
			call Response.Write( "<p>" & sSQL & "</p>") 
		end if
		
		Set oRS	= server.createobject("adodb.RecordSet")

		'Prepare dynamic cursor RS for dataretrieval
		oRS.CursorLocation =	3	'adOpenStatic 
		oRS.LockType =			1	'adLockReadOnly
		oRS.ActiveConnection =	oCon

		oRS.Open sSQL, oCon
		
		If ( oCon.Errors.Count > 0 ) Then
			call RenderSqlErrors(oCon,sSQL)
			CloseConnection(oCon)
			On Error Goto 0
			set GetDynamicRS = nothing
			Exit Function
		End If
		Set GetDynamicRS = oRS
		On Error Goto 0
	end function


	' ********************************************
	' Name: ExecuteCRUDSQL
	' Purpose: Executes a crud SQL statement (Create, Update, Delete and Insert)
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sSQL - Statement to execute
	'   [in] oCon - the Open connection to use
	'   [Returns] - TRUE/FALSE
	' Notes: If session("debug") is on, and the query fails
	' the errors will be rendered onscreen.
	' ********************************************
	function ExecuteCRUDSQL(sSQL, oCon)
		Dim oRS 
		Dim oCommand
		
		ExecuteCRUDSQL = FALSE		
		
		If Len( sSQL ) = 0 Then
			Exit Function
		End If
		if session("debug")="true"	then
			call Response.Write( "<p>" & sSQL & "</p>") 
		end if

		set oCommand = server.createobject("adodb.command")
		set oCommand.ActiveConnection =  oCon
		oCommand.commandText = sSQL
		On Error Resume Next
		call oCommand.execute (,,1 + 128)
		If ( oCon.Errors.Count > 0 ) Then
			call RenderSqlErrors(oCon,sSQL)
		else
			ExecuteCRUDSQL = TRUE
		End If
		set oCommand.ActiveConnection = nothing
		set oCommand = nothing
		On Error Goto 0
	end function



	' ********************************************
	' Name: GetConnection
	' Purpose: Opens a connection using the sConnectString as a connection string
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sConnectString - full connection string
	'   Returns an adodb.connection object if connection was sucessful, nothing else
	' Notes: Renders errors and returns nothing if connection failes
	' ********************************************
	function GetConnection(sConnectString)
		dim oCon
		
		Set oCon = server.createobject("adodb.Connection")

		On Error Resume Next
		
		'Establish connection to DB
		oCon.Open(sConnectString)
		
		If ( oCon.Errors.Count > 0 ) Then
			call RenderSqlErrors(oCon, sConnectString)
			On Error Goto 0
			set GetConnection = nothing
			Exit Function
		End If
		On Error goto 0
		
		set GetConnection = oCon
	end function

	' ********************************************
	' Name: HasRows
	' Purpose: check if recordset has records
	' Author: -
	' Parameters:
	'   [in] oRS - recordset
	'   [out] True if recordset has records
	' Notes: checks if recordset is object and if it is nothing
	' ********************************************
	Function HasRows(ByRef oRS)
		HasRows = False
		If ( IsObject( oRS ) ) Then
			If Not ( oRS Is Nothing ) Then
				If Not( oRS.State = 0 ) Then
					If Not( oRS.BOF And oRS.EOF ) Then
						HasRows = True
					End If
				End If
			End If
		End If
	end function
%>
