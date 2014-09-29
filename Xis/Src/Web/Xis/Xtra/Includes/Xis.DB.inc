<%
	const XIS = "XIS"
	const EFP = "EFP"
	
	'Prefixes used:
	'Render - Prefix for HTML output subs
	'Get  - Prefix for Get object functions


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
	function GetConnectionstring(strSource, strParams)
		dim strConnection
		if (strSource = XIS) then
			strConnection = Session("xtra_ConnectionString") & _
			"uid=" & Session("xtra_RuntimeUserName") & _
			";pwd=" & Session("xtra_RuntimePassword") & ";" & strParams
		elseif (strSource = EFP) then
			strConnection = application("IMPconnection") & strParams
		end if

		if (len(trim(strConnection)) = 0) then
			if session("debug") = true	then
				' Print all error messages from connection
				Response.write "<p class='warning'>Kunne ikke bygge oppkoblingsstreng!" & "</p>"
			end if
			GetConnectionstring = vbNullString
		else
			GetConnectionstring = trim(strConnection)
		end if
		
	End function


	' Name: RenderSqlErrors
	' Purpose: Renders all the errors as HTML
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] Conn - the connection object containing all the errors
	'   [in] Conn - A string containing additional information 
	'   regarding the context of the error: SQL, Connectionstring etc.
	' Notes: Only displays details about the error if session("debug") is set to true
	Sub RenderSqlErrors(Conn, Source)
		dim NOFErrors
		dim iCurrentError

		if session("debug") = true then
			' Print all error messages from connection
			Response.write "<p class='warning'>F&oslash;lgende feil oppstod:" & "</p>"
			Response.write "<p><strong>Kilde:</strong><br>" & Source & "</p>"
			NOFErrors = Conn.Errors.Count

			for iCurrentError = 0 to NOFErrors - 1
				Response.write "<p><strong>Description:</strong><br>" & Conn.Errors(iCurrentError).Description & "</p>"
			next
		else
			Response.write "<p class='warning'>En uopprettelig feil oppstod. Vennligt kontakt IT-ansvarlig." & "</p>"
		end if
	End Sub


	' Name: CloseConnection
	' Purpose: Closes the specified connection
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] Conn - the connection object to use close 
	' Notes: -
	sub CloseConnection(byref Conn)
		If (IsObject(Conn)) Then
			If (not Conn is nothing) Then
				If ( Conn.State <> 0) Then
					Conn.Close()
				End If
			end if
		End If
	end sub


	' Name: GetFirehoseRS
	' Purpose: Returns a forward readonly recordset (firehose)
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sSQL - Statement to execute
	'   [in] Conn - the Open connection to use
	'   [out] An adodb.recordset object
	' Notes: Renders errors and returns nothing if Query failes
	Function GetFirehoseRS(strSQL, Conn)
		Dim oRS 
			
		If LenB( strSQL ) = 0 Then
			Set getRecordset = Nothing
			Exit Function
		End If

		if session("debug") = true	then
			call Response.Write( "<p>" & strSQL & "</p>") 
		end if

		Set oRS	= Server.CreateObject("ADODB.Recordset")
		'Prepare firehose RS for dataretrieval
		oRS.LockType =	1 	'adLockReadOnly
		oRS.CursorType = 0	'adOpenForwardOnly 
		on error resume next
		oRS.Open strSQL, Conn
		If ( Conn.Errors.Count > 0 ) Then
			on error goto 0
			call RenderSqlErrors(Conn, strSQL)
			set getFirehoseRS = nothing
			Exit Function
		End If
		on error goto 0
		Set getFirehoseRS = oRS
	End Function
	
	
	' Name: GetFirehoseRecSet
	' Purpose: Returns a recordset for paging(firehose)
	' Author: TPH
	' Date: 12.05.2008
	' Parameters:
	'   [in] sSQL - Statement to execute
	'   [in] Conn - the Open connection to use
	'   [out] An adodb.recordset object
	' Notes: Renders errors and returns nothing if Query failes
	Function GetFirehoseRecSet(strSQL, Conn)
		Dim oRS 
			
		If LenB( strSQL ) = 0 Then
			Set getFirehoseRecSet = Nothing
			Exit Function
		End If

		if session("debug") = true	then
			call Response.Write( "<p>" & strSQL & "</p>") 
		end if

		Set oRS	= Server.CreateObject("ADODB.Recordset")
		'Prepare firehose RS for dataretrieval
		oRS.LockType =	1 	'adLockReadOnly
		oRS.CursorType = 0	'adOpenForwardOnly
		oRs.CursorLocation = 3  ' adUseClient
		on error resume next
		oRS.Open strSQL, Conn
		If ( Conn.Errors.Count > 0 ) Then
			on error goto 0
			call RenderSqlErrors(Conn, strSQL)
			set getFirehoseRecSet = nothing
			Exit Function
		End If
		on error goto 0
		Set getFirehoseRecSet = oRS
	End Function


	' Name: GetDynamicRS
	' Purpose: Returns a dynamic readonly recordset
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] strSQL - Statement to execute
	'   [in] Conn - the Open connection to use
	'   [out] An adodb.recordset object
	' Notes: Renders errors and returns nothing if Query failes
	' Being a dynamic recordset, the recordcount 
	'property on the returned recordset is supported
	function GetDynamicRS(strSQL, Conn)
		Dim oRS 
		
		
		If LenB(strSQL) = 0 Then
			Set GetDynamicRS = Nothing
			Exit Function
		End If

		if session("debug") = true	then
			call Response.Write( "<p>" & strSQL & "</p>") 
		end if
		
		Set oRS	= Server.CreateObject("ADODB.Recordset")

		'Prepare dynamic cursor RS for dataretrieval
		oRS.CursorType = 3	'adOpenStatic 
		oRS.LockType = 1	'adLockReadOnly
		on error resume next
		oRS.Open strSQL, Conn
		If ( Conn.Errors.Count > 0 ) Then
			on error goto 0
			call RenderSqlErrors(Conn, strSQL)
			set GetDynamicRS = nothing
			Exit Function
		End If
		on error goto 0
		Set GetDynamicRS = oRS
	end function


	' Name: ExecuteCRUDSQL
	' Purpose: Executes a crud SQL statement (Create, Update, Delete and Insert)
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] strSQL - Statement to execute
	'   [in] Conn - the Open connection to use
	'   [Returns] - TRUE/FALSE
	' Notes: If session("debug") is on, and the query fails
	' the errors will be rendered onscreen.
	function ExecuteCRUDSQL(strSQL, Conn)
		Dim oRS 
		Dim Cmd
		
		ExecuteCRUDSQL = false		
		
		If LenB( strSQL ) = 0 Then
			Exit Function
		End If
		if session("debug") = true then
			call Response.Write( "<p>" & strSQL & "</p>") 
		end if

		set Cmd = server.createobject("ADODB.Command")
		set Cmd.ActiveConnection = Conn
		Cmd.CommandText = strSQL
		Cmd.CommandType = &H0001 'adCmdText
		on error resume next
		Cmd.Execute 

		If ( Conn.Errors.Count > 0 ) Then
			call RenderSqlErrors(Conn, strSQL)
		else
			ExecuteCRUDSQL = true
		End If
		on error goto 0
		set Cmd.ActiveConnection = nothing
		set Cmd = nothing
	end function

	' Name: GetConnection
	' Purpose: Opens a connection using the sConnectString as a connection string
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sConnectString - full connection string
	'   Returns an adodb.connection object if connection was sucessful, nothing else
	' Notes: Renders errors and returns nothing if connection failes
	function GetConnection(sConnectString)
		dim Conn
		
		Set Conn = server.createobject("adodb.Connection")

		'Establish connection to DB
		Conn.Open(sConnectString)
		
		If ( Conn.Errors.Count > 0 ) Then
			call RenderSqlErrors(Conn, sConnectString)
			set GetConnection = nothing
			Exit Function
		End If
		
		set GetConnection = Conn
	end function

	' Name: GetClientConnection
	' Purpose: Opens a client connection using the sConnectString as a connection string
	' Author: FJM
	' Date: 14.03.2003
	' Parameters:
	'   [in] sConnectString - full connection string
	'   Returns an adodb.connection object if connection was sucessful, nothing else
	' Notes: Renders errors and returns nothing if connection failes
	function GetClientConnection(sConnectString)
		dim Conn
		
		Set Conn = server.createobject("ADODB.Connection")
		
		'Establish connection to DB
		Conn.CursorLocation = 3 'adUseClient
		Conn.Open(sConnectString)
		
		If ( Conn.Errors.Count > 0 ) Then
			call RenderSqlErrors(Conn, sConnectString)
			set GetConnection = nothing
			Exit Function
		End If
		
		set GetClientConnection = Conn
	end function

	' Name: HasRows
	' Purpose: check if recordset has records
	' Author: -
	' Parameters:
	'   [in] oRS - recordset
	'   [out] True if recordset has records
	' Notes: checks if recordset is object and if it is nothing
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
