<%option explicit%>
<!--#INCLUDE FILE="includes/db_lib.inc"-->
<%
	dim oRs
	dim oCon
	dim objFSO
	dim fld
	dim sSQL
	dim nCount
	dim sConnectionString	
	dim strHTTPAdress : strHTTPAdress	= Application("HTTPadress")
	dim FilePath
	dim strVikarID
	dim FileSource
	
	sConnectionString = GetConnectionstring("XIS","")
	
	set oCon = GetConnection(sConnectionString)
	sSQL = "SELECT [Vikar].[fornavn], [Vikar].[etternavn], [Vikar].[VikarID], [CV].[Filename] FROM [CV] INNER JOIN [VIKAR] ON [Vikar].[VikarID] = [CV].[ConsultantID] WHERE  [Filename] IS NOT NULL AND [TYPE] = 'C' "
	set oRs = GetFirehoseRS(sSQL, oCon) 
	
%>
<HTML>
	<HEAD>
		<TITLE>X|is - Settings</TITLE>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">

	<HEAD>
	<BODY>
		<div class="pageContainer" id="pageContainer">
			<div class="content">
				<h2>overføre Suspect CV</h2>
				<p class="listing">
				<TABLE>
				<tr>
					<th>Navn</th>
					<th>Opplasted CV-fil</th>
				</tr>
				<%				

				if HasRows(oRs) then
					Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
							
					while not (oRs.EOF)
						strVikarID = oRs.fields("vikarID").value
						FilePath = Application("ConsultantFileRoot") & strVikarID & "\"
						FileSource = Application("CVFileRoot") & "\" & oRs.fields("Filename").value
						If not objFSO.FolderExists(FilePath) then
							objFSO.createFolder(FilePath)	
							response.write "Lager:" & FilePath & "<br>"
						end if
						'Flytte CV fil
						
						If objFSO.FileExists(FileSource) then
							If not objFSO.FileExists(FilePath & "\" & oRs.fields("Filename").value) then
								objFSO.Movefile FileSource, FilePath 
								'objFSO.copyFile FileSource, FilePath
								response.write "Flytter:" & FileSource & " til " & FilePath & "...<br>"
							else
								response.write "Fil finnes fra før:" & FileSource & "!<br>"									
							end if
						else
							response.write "Fil finnes ikke:" & FileSource & "!<br>"							
						end if
						'Oppdatere CV referanse
						sSQL = "Update [CV]  Set [FileName]= NULL where [ConsultantID]='" & strVikarID & "' AND [Type] = 'C'"
						'response.write "sSQL:" & sSQL & "<br>"
						if not ExecuteCRUDSQL(sSQL, oCon) then
							response.end
						end if
						response.write "<tr><td>" & oRs.fields("Fornavn").value & " " & oRs.fields("etternavn").value & "</td><td><a href='http://"  & strHTTPAdress & "\Xtra\CVUpload\CVdok\" & oRs.fields("fileName").value & "' target='_blank'>" & oRs.fields("fileName").value &  "</a></td></tr>"
						oRs.movenext
						response.write "<br>"
					wend
					set objFSO = nothing
					oRs.close

				end if
				CloseConnection(oCon)
				set oCon = nothing
				set oRs = nothing
				%>
				</TABLE>
				</p>
			</div>
		</DIV>
	</BODY>
</HTML>