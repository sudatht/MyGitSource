<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.SuperOffice.Integration.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	dim brukerID
	dim valg
	dim SlettAction
	dim strID
	dim rsOppdrag
	dim Conn
	dim strKunde

	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")					
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	
	
	brukerID	= Session("BrukerID")
	SlettAction	= trim(Request.QueryString("Slett"))
	strID		= Request.QueryString("ID")
	
	if (Session("medarbID") = 0) then
	'user must be logged on, abort!
		AddErrorMessage("Du er ikke logget på!")
		call RenderErrorMessage()	
	end if


	' slette oppdrag fra hotlist
	if (len(SlettAction) > 0) then
		If SlettAction = "oppdrag" Then
			strSQL = "Delete from HOTLIST where ID = " & strID
			Conn.Execute(strSQL)
		End If
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Hotlist oppdrag</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='../js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='../js/navigation.js' id='navigationScripts'></script>		
		<script language="javaScript" type="text/javascript">
			//lager felles variabler
			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				<%
				If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
					%>				
					if (modKey && modKey2 && keyChar=="P")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistOppdrag.asp");
					}
					<%
				end if
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="V")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistVikar.asp");
					}
					<%
				end if
				%>
			}
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	<body onLoad="fokus()">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<a id="Top"></a>
				<h1>Mine oppdrag</h1>
			</div>
			<div class="content">
				<div class="listing">				
					<%
					strOppdrag = "select H.*, O.Beskrivelse, O.OppdragId, O.Oppdragskode, F.Firma, isnull(F.CRMAccountGuid,'') AS CRMAccountGuid, F.SOCuID, O.Fradato, O.TilDato, S.OppdragsStatus "&_
					" FROM HOTLIST H, Oppdrag O, Firma F, h_oppdrag_status S " &_
					" WHERE O.FirmaID = F.FirmaID " &_
					" AND H.navnID = F.FirmaID " &_
					" AND H.status = 1 " &_
					" AND H.brukerID = " & brukerID &_
					" AND H.oppdragID = O.oppdragID " &_
					" AND O.StatusID *= S.Oppdragsstatusid " &_
					" ORDER BY H.navn "

					set rsOppdrag = GetFirehoseRS(strOppdrag, Conn)		
					%>
					<table width="96%" ID="Table1">
						<tr>
							<th>Oppdragnr</th>
							<th>Oppdragsbeskrivelse</th>
							<th>Kunde</th>
							<th>Fra Dato</th>
							<th>Til Dato</th>
							<th>Status</th>
							<th>Slette</th>
						</tr>
						<%
						' Show search result
						while (not rsOppdrag.EOF )
							%>
							<tr>
								<td><%=rsOppdrag("oppdragID")%></td>
								<td><%=CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?oppdragID=" & rsOppdrag("oppdragID").Value, rsOppdrag( "beskrivelse"), "Vis Oppdrag '" & rsOppdrag( "beskrivelse") & "'" )%></td>
								
								<%
									linkurl = Application("CRMAccountLink") & rsOppdrag("CRMAccountGuid") & "%7d&pagetype=entityrecord"										
									strKunde = "<a href=" & linkurl & " target='_blank'>" & rsOppdrag("Firma").Value & " </a>"										
								%>											
											
								<td><%=strKunde%></td>
								<td><%=rsOppdrag("fraDato")%></td>
								<td><%=rsOppdrag("tilDato")%></td>	
								<td><%=rsOppdrag("oppdragsStatus")%></td>
								<td><a href="hotlistOppdrag.asp?Slett=oppdrag&ID=<%=rsOppdrag("ID")%>" ><img src="../Images/icon_delete.gif" alt="Slett"></a></td>
							</tr>
							<%
							rsOppdrag.MoveNext 
						wend 
						rsOppdrag.close
						set rsOppdrag = nothing 
						%>
					</table>
				</div>
				<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="Til toppen">Til toppen</a><br>&nbsp;
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing		
%>