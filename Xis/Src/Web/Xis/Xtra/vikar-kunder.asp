<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim RecordsFound : RecordsFound = false	
	dim strSQL
	dim Conn
	dim VikarID
	dim selection
	dim rsVikar
	dim VikarNavn
	dim strKunde
	
	const DEFAULT_OPPDRAGVIKAR_STATUS = 4	

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") = "" Then

		If Request( "VikarID" ) = "" Then
			AddErrorMessage("Feil:VikarID mangler!")
			call RenderErrorMessage()		
		Else
			VikarID = Request( "VikarID" )
		End If
		selection = DEFAULT_OPPDRAGVIKAR_STATUS
	Else
		VikarID = Request( "tbxVikarID" )
		selection = Request("rbnselection")
	End If

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' hent navn på vikar
	strSQL = "SELECT Navn=Fornavn+' '+Etternavn FROM VIKAR WHERE VikarID = " & VikarID 
	Set rsVikar = GetFirehoseRS(strSQL, Conn )
	Vikarnavn = rsVikar( "Navn" )
	rsVikar.Close: set rsVikar = Nothing

	' SQL for search
   	strSQL = "SELECT DISTINCT F.FirmaID, F.SOCuID,F.CRMAccountGuid, F.Firma, OV.OppdragID" &_
		", OV.Fradato, OV.Tildato" &_
		", O.Beskrivelse, S.Status , MB.Etternavn,MB.fornavn " &_
		" FROM OPPDRAG_VIKAR OV, OPPDRAG O, FIRMA F, H_OPPDRAG_VIKAR_STATUS S, MEDARBEIDER  MB " &_
		" WHERE OV.VikarID = " & VikarID &_
		" AND O.OppdragID = OV.OppdragID" &_
		" AND F.FirmaID = OV.FirmaID" &_
		" AND S.OppdragVikarStatusID = OV.StatusID" &_
		" AND O.AnsmedID=MB.MedID"

    ' Append selection of status if other than all SELECTed
	If selection > 0 Then
		strSQL = strSQL & " AND OV.statusID = " & selection
	End If

	strSQL = strSQL &_
		" UNION " &_
		"SELECT FirmaID=o.Kundenr, F.SOCuID, isnull(F.CRMAccountGuid,'') AS CRMAccountGuid, F.Firma, OppdragID = o.Oppdragsnr" &_
		", Fradato = o.StartDato, Tildato=o.Sluttdato" &_
		", o.Beskrivelse, Status = S.OppdragsStatus ,Etternavn=null, fornavn= null" &_
		" FROM gamle_oppdrag o, Firma F, H_OPPDRAG_STATUS S" &_
		" WHERE Vikarnr = " & VikarID &_
		" AND S.OppdragsStatusID = o.StatusID" &_
		" AND o.kundenr = F.FirmaID"
		

	strSQL = strSQL & " ORDER BY OV.Fradato DESC"
	
	set rsRapport = GetFirehoseRS(strSQL, Conn)

	If (HasRows(rsRapport) = true) Then
		RecordsFound = true
	End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Tidligere kontakter for <%=Vikarnavn%></title>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
		<script language='javascript' src='js/contentMenu.js'></script>
		<script language="javascript">
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
				<h1>Tidligere kontakter for <%=Vikarnavn%></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" id="Table1">
						<tr>
							<td>
								<table cellpadding="0" cellspacing="2" id="Table3">
									<tr>
										<td class="menu" id="menu1" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);">
											<a href="/xtra/vikarvis.asp?vikarid=<%=VikarID%>" title="Vis vikar">
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis</a>
										</td>
									</tr>
								</table>
							</td>
							<td class="right">
							<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>
						</tr>
					</table>
				</div>			
			</div>
			<div class="content">
				<form name="formEn" ACTION="vikar-kunder.asp" METHOD="POST" ID="Form1">
					<input type="hidden" NAME="tbxPageNo" VALUE="1" ID="tbxPageNo">
					<input type="hidden" NAME="tbxVikarID" VALUE="<%=VikarID%>" ID="tbxVikarID">
					<table>
						<tr>
							<%
							strSQL = "SELECT [OppdragVikarStatusID], [Status] FROM [H_OPPDRAG_VIKAR_STATUS] ORDER BY [Status]"
							set rsStatus = GetFirehoseRS(strSQL, Conn)
							if(HasRows(rsStatus) = true) then
								%>
								<td>
									<select id="rbnSelection" name="rbnSelection">
										<%
										if(selection = DEFAULT_OPPDRAGVIKAR_STATUS) then
											selected = "selected"
										end if
										%>			
										<option <%=selected%> value="0">(alle)</option>
										<%
										while not rsStatus.EOF
											if(clng(selection) = rsStatus("OppdragVikarStatusID").value) then
												selected = "selected"
											else
												selected = ""
											end if
											%>
											<option <%=selected%> value="<%=rsStatus("OppdragVikarStatusID")%>"><%=rsStatus("Status")%></option>
											<%
											rsStatus.MoveNext
										wend
										%>
									</select>
								</td>
								<%
								rsStatus.close
							end if
							set rsStatus = nothing
							%>
							<td><input type="submit" name="pbnDataAction" value="Søk" id="Submit1"></td>
						</tr>
					</table>
				</form>
				<%
				' Display
				If (RecordsFound) Then 
					%>
					<div class="listing">
						<table width="98%" cellpadding="0" cellspacing="1" ID="Table2">
							<tr>
								<th>Kontakt</th>
								<th>Oppdrag</th>
								<th>Beskrivelse</th>
								<th>Status</th>
								<th>Ansvarlig </th>
								<th>Fradato</th>
								<th>Tildato</th>
								<th>Ant Dager</th>
								
							</tr>
							<%
							STAT = ""						'for å ferske statusendringer (for ikke å blande gamle med nye)
							OPPD = ""					 	'for å ferske forskjellige oppdrag i loopen
							tilDato = rsRapport("TilDato")	'høyeste tildato i OPPDRAG_VIKAR
							fraDato = rsRapport("FraDato")	'laveste fradato i OPPDRAG_VIKAR
							antDager = 0					'antall dager samlet i OPPDRAG_VIKAR

							Do Until rsRapport.EOF
								If rsRapport("OppdragID") <> OPPD Or rsRapport("Status") <> STAT Then
									If OPPD <> "" Then  'ikke første gangen 
										%>
										<td><% =fradato %></td>
										<td><% =tilDato %></td>
										<td class="right"><% =antDager %></td>
										<% 
										antDager = 0 : fradato = "" : tildato = "" 
									End If 
									%>
									</tr>
									<tr>
									
										<%
											linkurl = Application("CRMAccountLink") & rsRapport("CRMAccountGuid") & "%7d&pagetype=entityrecord"
											strKunde = "<a href=" & linkurl & " target='_blank'>" & rsRapport("Firma").Value & " </a>"							
										%>
								
										<td><%=strKunde%></td>
																				
										<% 
										If rsRapport("Status") = "Ferdig" Then 	'ikke link til oppdrag før 01.01.98 
											%>
											<td><% =rsRapport("OppdragID") %></td>
											<% 
										Else 
											%>
											<td><%=CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & rsRapport("oppdragID").Value, rsRapport( "OppdragID"), "Vis Oppdrag '" & rsRapport( "beskrivelse") & "'" ) %></td>										
											<% 
										End If 
										%>
											  
									
										<td><%=rsRapport( "Beskrivelse" ) %></td>
										<td><% =rsRapport( "Status" ) %></td>
											<td><% =rsRapport("Etternavn")& " " & rsRapport("fornavn") %></td>
										<%
										OPPD = rsRapport("OppdragID")    'setter nytt oppdrag
										STAT = rsRapport("Status")
									End If 		'nytt oppdrag i loopen
								antDager = antDager + rsRapport("Tildato") - rsRapport("Fradato") + 1
								If tilDato < rsRapport("TilDato") Or fraDato = "" Then
									tilDato = rsRapport("Tildato")
								end if
								If fraDato > rsRapport("fraDato") Or fraDato = "" Then
									fraDato = rsRapport("fradato")
								end if
								
								rsRapport.MoveNext
							Loop 
							%>
							<td><% =fradato %></td>
							<td><% =tilDato %></td>
							<td class="right"><% =antDager %></td>
						
						</tr>
						<%
						rsRapport.close
					End If 'Records found
					set rsRapport = nothing
					%>
				</table>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>