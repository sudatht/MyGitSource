<%option explicit%>
<%
dim strServer				'Name of SQL server or path to ACCESS database
dim strDatabase				'Initial catalog / database on SQL Server
dim strUsername				'Username, if required
dim strPassword				'password, if required
dim strTimeout				'Number of seconds before server gives up on connection attempt
dim strProvider				'What kind of provider to use to connect to database
dim strNetlib				'SQL server specific connection string parameter
dim strConnection			'Connection string used to connect to database, either concatinated from all
							'all the parameters or user defined.
dim ConConnector			'instance of ADODB.connection used to connect to database
dim rsDBTables				'recordset containing definitions of all the tables on the server
dim rsDBCols				'recordset containing definitions of all the columns in the selected table
dim strConnectionStatus		'Result of an ettempt to connect to a database or a helpful message
dim strColValue				'The value of a field in a table 
dim objfld					'Field object, used to iterate through table/column definitions
dim objError				'Error object, used to iterate through all errors after a connection attempt
dim blnConnected			'True if a successfull connection was made, false otherwise
dim strURLConnection		'Connection string stored as a url
dim intTableID				'id of a selected table (SQL server only)

blnConnected = false		'Initialize to false, no connection has been attempted

'SQL stored in constants for easy maintance:
'SQL for table definition retrieval
const cSQLGetTables = "Select id, name, crdate from sysobjects where xtype='u' order by name"
'SQL for column definition retrieval
const cSQLGetCols = "Select name, xtype, length from syscolumns where id = "
'Color scheme for page stored in constants, for easy maintaince:
const cColorLight = "lightblue"
const cColorNormal = "blue"
'Name of page to postback to
const C_PAGE_NAME = "dbConnection.asp"

'Get params if page is posted, posted if hidden field contains the value "1"
if (cint(request.querystring("hdnPosted")) = 1)  then
	 strServer		= request.querystring("txtServer")
	 strDatabase	= request.querystring("txtDatabase")
	 strUsername	= request.querystring("txtUsername")
	 strPassword	= request.querystring("txtPassword")
	 strConnection	= request.querystring("txtConnection")
	 strTimeout		= request.querystring("txtTimeout")
	 strNetLib		= request.querystring("txtNetLib")
	 strProvider	= request.querystring("txtProvider")

	 'If a querystring is posted, it overrides individual parameters
	 if len(trim(strConnection)) = 0 then
		'If no timeout value is specified, default to 15 seconds
		if len(trim(strTimeout)) = 0 then
			strTimeout = "15"
		end if

		'Different connection strings depending on which provider is used (Access or SQL server specific):
		'SQL server:
		if strProvider = "SQLOLEDB.1" then
	 		strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strServer & _
	 		";Initial Catalog=" & strDatabase & ";" & _
			"User ID=" & strUsername & ";" &_
			"Password=" & strPassword & ";" & _
			"timeout=" & strTimeout & ";" & _
			"Network Library=" & strNetLib & ";" 		
		'Access:
		elseif strProvider = "Microsoft.Jet.OLEDB.4.0" then
			strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strServer 
			if len(trim(strUsername))>0 then
				strConnection = strConnection & "User ID=" & strUsername & ";"
				if len(trim(strPassword))>0 then
					strConnection = strConnection & "Password=" & strPassword & ";"
				end if
			end if
		end if
	 end if

	'Try to connect and determine connection status..
	Set ConConnector = Server.CreateObject("ADODB.Connection")
	on error resume next
	ConConnector.Open strConnection
	'If any errors occured, this collection contains a description and number
	if ConConnector.errors.count > 0 then
		strConnectionStatus = "Unable to connect!" & vbcrlf
		for each objError in ConConnector.errors
			strConnectionStatus = strConnectionStatus & objError & vbcrlf
		next
	else
		strConnectionStatus = "Connection successful."
		blnConnected = true
		strURLConnection = server.urlencode(strConnection)
	end if
	on error goto 0
else
		'If page has not been posted, show helpful message..
		strConnectionStatus = "Please either fill out indiviual server parameters OR connection string." & vbcrlf & _
		"hit 'connect' to test and verify connection."
end if

%>
<HTML>
<HEAD>
	<TITLE>SQL Server/Access Connector</TITLE>
	<style type="text/css">
		SPAN.MenuCategorySelected {margin:0px;padding-left:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:none }
		SPAN.MenuCategory		 {margin:0px;padding-right:3px; border-top: 1px solid black; border-left:1px solid black; border-right:2px solid black; border-bottom:1px solid black }
	</style>
</HEAD>
	<BODY>
		<!--#INCLUDE FILE="top_menu.asp"-->	
		<FORM ACTION="<%=C_PAGE_NAME%>" METHOD="GET">
			<input type="hidden" name="hdnPosted" id="hdnPosted" value="1">
			<TABLE border="0" cellpadding="0" cellspacing="0" width="80%" align="center">
				<TR>
					<TD colspan="3" align="center"><H3 style="color:<%=cColorNormal%>">SQL Server/Access Connector<H3></TD>
				<TR>
				<TR>
					<TD colspan="3" align="left"><i>Connection parameters</i></TD>
				<TR>
				<TR>
					<TD  colspan="3"><HR style="color:<%=cColorLight%>" width="100%"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Server:&nbsp;</TD>
					<TD  align="left"><input name="txtServer" id="txtServer" value="<%=strServer%>">&nbsp;Name of SQL server or path to Access database.</TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Username:&nbsp;</TD>
					<TD  align="left"><input name="txtUsername" id="txtUsername" value="<%=strUsername%>"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Password:&nbsp;</TD>
					<TD  align="left"><input name="txtPassword" id="txtPassword" value="<%=strPassword%>"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Provider:&nbsp;</TD>
					<TD  align="left">
					<SELECT name="txtProvider" id="txtProvider">
					<%
					if strProvider = "SQLOLEDB.1" then
					%>
						<OPTION value="SQLOLEDB.1" selected>SQL Server (SQLOLEDB.1)</OPTION>
					<%
					else
					%>
						<OPTION value="SQLOLEDB.1" >SQL Server (SQLOLEDB.1)</OPTION>
					<%
					end if

					if strProvider = "Microsoft.Jet.OLEDB.4.0" then
					%>
						<OPTION value="Microsoft.Jet.OLEDB.4.0" selected>Access (Microsoft.Jet.OLEDB.4.0)</OPTION>
					<%
					else
					%>
						<OPTION value="Microsoft.Jet.OLEDB.4.0">Access (Microsoft.Jet.OLEDB.4.0)</OPTION>
					<%
					end if
					%>
					</SELECT>	
					</TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Database:&nbsp;</TD>
					<TD  align="left"><input name="txtDatabase" id="txtDatabase" value="<%=strDatabase%>" style="background-color:<%=cColorLight%>">&nbsp;Which database to connect to.</TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Timeout:&nbsp;</TD>
					<TD  align="left"><input name="txttimeout" id="txttimeout" maxlength="3" value="<%=strTimeout%>" style="background-color:<%=cColorLight%>"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Network Library:&nbsp;</TD>
					<TD  align="left">
					<SELECT name="txtNetLib" id="txtNetLib" style="background-color:<%=cColorLight%>">
					<%
					if strNetLib = "dbmssocn" then
					%>
						<OPTION value="dbmssocn" selected>TCP/IP (dbmssocn)</OPTION>
					<%
					else
					%>
						<OPTION value="dbmssocn" >TCP/IP (dbmssocn)</OPTION>
					<%
					end if

					if strNetLib = "dbnmpntw" then
					%>
						<OPTION value="dbnmpntw" selected>Named Pipes (dbnmpntw)</OPTION>
					<%
					else
					%>
						<OPTION value="dbnmpntw">Named Pipes (dbnmpntw)</OPTION>
					<%
					end if

					if strNetLib = "dbmsrpcn" then
					%>
						<OPTION value="dbmsrpcn" selected>Multiprotocol ((RPC) dbmsrpcn)</OPTION>
					<%
					else
					%>
						<OPTION value="dbmsrpcn">Multiprotocol ((RPC) dbmsrpcn)</OPTION>
					<%			
					end if

					if strNetLib = "dbmsspxn" then
					%>				
						<OPTION value="dbmsspxn" selected>NWLink IPX/SPX (dbmsspxn)</OPTION>
					<%
					else
					%>				
						<OPTION value="dbmsspxn">NWLink IPX/SPX (dbmsspxn)</OPTION>
					<%			
					end if

					if strNetLib = "dbmsadsn" then
					%>				
						<OPTION value="dbmsadsn" selected>AppleTalk (dbmsadsn)</OPTION>
					<%
					else
					%>				
						<OPTION value="dbmsadsn">AppleTalk (dbmsadsn)</OPTION>
					<%			
					end if

					if strNetLib = "bmsvinn" then
					%>					
						<OPTION value="bmsvinn" selected>Banyan VINES (bmsvinn)</OPTION>
					<%
					else
					%>				
						<OPTION value="bmsvinn">Banyan VINES (bmsvinn)</OPTION>
					<%				
					end if
					%>
					</SELECT>	
					</TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD align="right"><div style="width:15px;background-color:<%=cColorLight%>;border-style:inset;border-width:thin;">&nbsp;</div></TD>
					<TD align="left">&nbsp;denotes sql server specific parameters.</TD>
				<TR>		
				<TR>
					<TD  colspan="3">&nbsp;</TD>
				<TR>
				<TR>
					<TD colspan="3" align="left"><i>Connection string</i></TD>
				<TR>
				<TR>
					<TD  colspan="3"><HR style="color:<%=cColorLight%>" width="100%"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">Connection:&nbsp;</TD>
					<TD  align="left"><Textarea rows="8" cols="50" name="txtConnection" id="txtConnection"><%=strConnection%></textarea></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD>&nbsp;</TD>		
					<TD align="left"><input type="button" onclick="javascript:document.all.txtConnection.value=''" value="reset connection string"></TD>
				<TR>
				<TR>
					<TD  colspan="3">&nbsp;</TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD>&nbsp;</TD>			
					<TD align="left"><input type="submit" name="btnConnect" id="btnConnect" value="connect"></TD>
				<TR>
				<TR>
					<TD colspan="3" align="left"><i>Connection status</i></TD>
				<TR>
				<TR>
					<TD  colspan="3"><HR style="color:<%=cColorLight%>" width="100%"></TD>
				<TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD  align="right">&nbsp;</TD>
					<TD  align="left"><Textarea rows="8" cols="50"><%=strConnectionStatus%></textarea></TD>
				<TR>			
			</TABLE>
		</FORM>
		<%
		'If an successfull connection has been made, get tables definitions..
		if blnConnected then
			'If Access, open schema for database and list all tables..
			if strProvider = "Microsoft.Jet.OLEDB.4.0" then
				set rsDBTables = ConConnector.OpenSchema(20) 'adSchemaTables
				if NOT rsDBTables.EOF then
					response.write "<TR>"
					for each objfld in rsDBTables.fields
						Response.Write "<TD>" & objfld.name & "</TD>"
					next
					response.write "</TR>"
					%>
					<TABLE border="1" cellpadding="0" cellspacing="2" width="80%" align="center">
					<%		
					while not rsDBTables.EOF
						response.write "<TR>"
						for each objfld in rsDBTables.fields
							Response.Write "<TD>" & objfld.value & "</TD>"
						next
						response.write "</TR>"
						rsDBTables.movenext
					wend
				end if
				%>
				</TABLE>
				<%				
			elseif strProvider = "SQLOLEDB.1" then
				set rsDBTables = ConConnector.execute(cSQLGetTables)
				if NOT rsDBTables.EOF then
					%>
					<TABLE border="0" cellpadding="0" cellspacing="2" width="80%" align="center">
						<TR>
							<TD colspan="3" align="left"><i>Database tables</i></TD>
						<TR>
						<TR>
							<TD  colspan="3"><HR style="color:<%=cColorLight%>" width="100%"></TD>
						<TR>
						<%
						response.write "<TR>"
						response.write "<TD><STRONG>Expand</STRONG></TD>"
						response.write "<TD><STRONG>Name</STRONG></TD>"
						response.write "<TD><STRONG>Created date</STRONG></TD>"
						response.write "</TR>"

						while not rsDBTables.EOF
							intTableID =  rsDBTables("id").value
							response.write "<TR>"
							response.write "<TD><STRONG><a id='#" & intTableID & "' name='#"& intTableID & "' href='" & C_PAGE_NAME & "?txtProvider=" & server.URLEncode(strProvider) &"&hdnPosted=1&id=" & intTableID & "&txtConnection=" & strURLConnection & "#" & intTableID & "'>Expand</a></STRONG></TD>"

							response.write "<TD>" &  rsDBTables("name").value & "</TD>"
							response.write "<TD>" &  rsDBTables("crdate").value & "</TD>"


							response.write "</TR>"
							if (rsDBTables("id")=clng(request.querystring("id"))) then
								set rsDBCols = ConConnector.execute(cSQLGetCols & rsDBTables("id") & " order by name")
								%>
								<TR>
									<TD colspan="4">
										<TABLE border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
								<%
								dim strbgStyle
								dim lngRowcount
								
								strbgStyle = "white"
								lngRowcount = 1
								response.write "<TR>"
								for each objfld in rsDBCols.fields
									response.write "<TD><STRONG>" &  objfld.name & "</STRONG></TD>"
								next
								response.write "</TR>"
								while not rsDBCols.EOF
									if lngRowcount mod 2 = 0 then
										strbgStyle = "white"
									else
										strbgStyle = cColorLight
									end if
									response.write "<TR style='background-color:" & strbgStyle & "'>"
									for each objfld in rsDBCols.fields
										if isnull(objfld.value) then
											strColValue = "&nbsp;"
										else
											strColValue = objfld.value
										end if
										response.write "<TD>" &  strColValue & "</TD>"
									next
									response.write "</TR>"
									lngRowcount = lngRowcount + 1 
									rsDBCols.movenext
								wend
								%>
										</TABLE>
									</TD>
								</TR>
								<%
							end if
							rsDBTables.movenext
						wend
						%>
					</TABLE>
					<%
				end if
			end if
			set rsDBTables = nothing
			set rsDBCols = nothing
			set ConConnector = nothing
		end if
		%>
	</BODY>
</HTML>