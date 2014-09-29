<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	brukerID	= Session("BrukerID")

	dim OSuspect
	dim lVikarID
	dim strFornavn
	dim strEtternavn
	dim strFodtselsdato
	dim strAdresse
	dim strPostAdresse
	dim oConn
	dim oCommand
	dim rsDuplicates
	dim strHeading
	dim rsVikar
	dim strOpprettet
	dim strSQL

	strHeading = "Mulig duplikat oppf&oslash;ring"


	'Parameter retrieval and validation
	' Move parameters to local variables
	lVikarID = Request.Querystring("suspectID")

	' Check VikarID
	If lVikarID = "" Then
		AddErrorMessage("Suspectid  mangler!")
	   call RenderErrorMessage()	
	End If

	set OSuspect					= Server.CreateObject("XtraWeb.Suspect")
	OSuspect.XtraConString			= Application("Xtra_intern_ConnectionString")
	OSuspect.XtraDataShapeConString = Application("ConXtraShape")

	if not OSuspect.GetSuspect(lVikarID) then
		OSuspect.JobApplication.CleanUp()
		OSuspect.CV.CleanUp()
		OSuspect.CleanUp
		set OSuspect = nothing
		AddErrorMessage("Suspect ikke funnet!")
		call RenderErrorMessage()	
	end if

	if OSuspect.DataValues("overfort") then
		OSuspect.JobApplication.CleanUp()
		OSuspect.CV.CleanUp()
		OSuspect.CleanUp
		set OSuspect = nothing
		AddErrorMessage("Suspect allerede overført!")
		call RenderErrorMessage()	   
	end if

	'Transfer name and birthdate to variables, for future use
	strFornavn			= OSuspect.DataValues("forNavn")
	strEtternavn		= OSuspect.DataValues("etterNavn")

	if isnull(OSuspect.DataValues("foedselsDato")) then
		strFodtselsdato = "null"
	else
		strFodtselsdato = OSuspect.DataValues("foedselsDato")
	end if

	' Open database connection
	Set oConn = GetConnection(GetConnectionstring(XIS, ""))	

	strSQL = "select A.Adresse, A.Postnr, A.PostSted " & _
			"from V_SUSPECT V, V_SUSPECT_ADRESSE A" &_
			" where V.suspectID = " & lVikarID  &_
			" and V.suspectID = A.adresseRelID" & _
			" and A.AdresseType = 2"

	set rsVikar = GetFirehoseRS(strSQL, oConn)

	if (HasRows(rsVikar)) then
		strAdresse			= rsVikar.fields("Adresse").value
		strPostAdresse		= rsVikar.fields("Postnr").value & "&nbsp;" & rsVikar.fields("PostSted").value
	end if
	rsVikar.close
	set rsVikar = nothing

	'We have no use for this object clean it up, clean it up
	OSuspect.JobApplication.CleanUp()
	OSuspect.CV.CleanUp()
	OSuspect.cleanup()
	set OSuspect = nothing

	'Retrieve all consultants that are possible duplicates
	strSQL = "EXEC sp_intra_getPossibleDuplicates '"  & strFornavn & "', '" & strEtternavn & "', " & dbdate(strFodtselsdato)

	set rsDuplicates = GetFirehoseRS(strSQL, oConn)

	if HasRows(rsDuplicates) = false then
		rsDuplicates.close
		set rsDuplicates = nothing
		CloseConnection(objCon)
		set objCon = nothing
		'There are no possible duplicates, redirect to tranfer suspect page
		Response.Redirect "SuspectNy.asp?suspectID=" & lVikarID
	end if
%>
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading%></title>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>		
		<script language="javaScript" type="text/javascript">		
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				//linker i submeny
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>		
					if (modKey && modKey2 && keyChar=="S")
					{
						parent.frames[funcFrameIndex].location=("/xtra/vikarSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>
					if (modKey && modKey2 && keyChar=="Y")
					{
						parent.frames[funcFrameIndex].location=("/xtra/VikarDuplikatSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="W")
					{
						parent.frames[funcFrameIndex].location=("/xtra/jobb/SuspectList.asp");
					}
					<% 
				End If 
				%>
			}
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;			
		</script>
	</head>		
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
			</div>
			<div class="content">		
				<form Name="suspekt" action="SuspectNy.asp?suspectID=<%=lVikarID%>" method="post">
					<div class="listing">
						<table>
							<tr>
								<th colspan="6" >Du pr&oslash;ver &aring; overf&oslash;re f&oslash;lgende suspect</th>
							</tr>
							<tr>
								<th>&nbsp;</th>
								<th>f&oslash;dselsdato</th>
								<th>Vikar</th>
								<th>Adresse</th>
								<th colspan="2">Postadresse</th>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td><%=strFodtselsdato%>&nbsp;</td>
								<td><%=strFornavn%>&nbsp;<%=strEtternavn%></td>
								<td><%=strAdresse%>&nbsp;</th>
								<td colspan="2"><%=strPostAdresse%></td>

							</tr>
							<tr>
								<th colspan="6" >F&oslash;lgende mulige duplikater finnes</th>
							</tr>
							<tr>
								<th>Ansattnr</th>
								<th>f&oslash;dselsdato</th>
								<th>Vikar</th>
								<th>Adresse</th>
								<th>Postadresse</th>
								<th>Opprettet</th>
							</tr>
							<%
							while (not rsDuplicates.EOF)
								%>
								<tr>
									<td><%=rsDuplicates("AnsattNummer").value%>&nbsp;</td>
									<td><%=rsDuplicates("foedselsDato").value%>&nbsp;</td>
									<td><a href='../vikarvis.asp?vikarid=<%=rsDuplicates("vikarid").value%>'><%=rsDuplicates("etternavn").value & ", " & rsDuplicates("fornavn").value%></a></td>
									<td><%=rsDuplicates("Adresse").value%>&nbsp;</td>
									<td><%=rsDuplicates("Postnr").value & " " & rsDuplicates("poststed").value%>&nbsp;</td>
									<td><%=rsDuplicates("regdato").value%>&nbsp;</td>
								</tr>
								<%
								rsDuplicates.movenext
							wend
							rsDuplicates.close
							set rsDuplicates = nothing
							%>
						</table>
					</div>						
					<table>
						<tr>
							<td>
								<input type='submit' value="Opprett som ny" id="submit">
							</td>
						</tr>
					</table>
				</form>
			</div>
	    </div>
	</body>
</html>
<%
CloseConnection(objCon)
set objCon = nothing
%>