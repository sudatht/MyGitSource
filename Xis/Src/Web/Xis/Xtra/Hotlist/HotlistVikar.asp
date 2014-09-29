<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	dim brukerID
	dim SlettAction
	dim strID
	dim objResource

	set objResource = Server.CreateObject("Localizer.ResourceManager")

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
	
	brukerID	= Session("BrukerID")
	SlettAction	= trim(Request.QueryString("Slett"))
	strID		= Request.QueryString("ID")
	
	if (Session("medarbID") = 0) then
	'user must be logged on, abort!
		AddErrorMessage(objResource.GetText("MsgNotLogged"))
		call RenderErrorMessage()
	end if


	' Slette hotlist vikar fra hotlist
	if (len(SlettAction) > 0) then
	'Only check type of delete action if it is a delete action
		If SlettAction = "vikar" Then
			strSQL = "Delete from HOTLIST where ID = " & strID
			conn.Execute(strSQL)
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
		<title><%objResource.WriteText("Heading")%></title>
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
				<h1><%objResource.WriteText("Heading")%></h1>
			</div>
			<div class="content">
				<div class="listing">
					<%
					strVikar = "SELECT " & _
					"HOTLIST.ID, " & _
					"HOTLIST.BrukerID, " & _
					"HOTLIST.navnID, " & _
					"HOTLIST.navn, " & _
					"HOTLIST.status, " & _
					"VIKAR.VikarID, " & _
					"VIKAR.Etternavn, " & _
					"VIKAR.Telefon, " & _
					"VIKAR.MobilTlf, " & _
					"VIKAR.Fax, " & _
					"VIKAR.EPost, " & _
					"ADRESSE.Adresse, " & _
					"ADRESSE.Postnr, " & _
					"ADRESSE.PostSted, " & _
					"VIKAR_ANSATTNUMMER.ansattnummer " & _
					"FROM HOTLIST " & _
					"LEFT OUTER JOIN VIKAR ON HOTLIST.NavnID = VIKAR.Vikarid " & _
					"LEFT OUTER JOIN ADRESSE ON VIKAR.Vikarid = ADRESSE.adresseRelID " & _
					"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
					"WHERE VIKAR.vikarID = HOTLIST.navnID " & _
					"AND HOTLIST.status = '3' " & _
					"AND HOTLIST.brukerID = " & brukerID & _
					"AND ADRESSE.AdresseRelasjon = '2' " & _
					"AND ADRESSE.AdresseType = '1' " & _
					"ORDER BY VIKAR.Etternavn "

					set rsVikar = GetFirehoseRS(strVikar, Conn)
					%>
					<table width="96%" ID="Table1">
						<tr>
							<th><%objResource.WriteText("AnsattNo")%></th>
							<th><%objResource.WriteText("Vikar")%></th>
							<th><%objResource.WriteText("Address")%></th>
							<th><%objResource.WriteText("Telephone")%></th>
							<th><%objResource.WriteText("Mobile")%></th>
							<th><%objResource.WriteText("Epost")%></th>
							<th><%objResource.WriteText("Delete")%></th>
						</tr>
						<%
						' Show search result				
						While (not rsVikar.EOF)
							%>
							<tr>
								<td>
									<%
									if (rsVikar("ansattnummer").Value <> "" ) then
										Response.Write rsVikar("ansattnummer").Value
									else
										Response.Write "---"
									end if
									%>
								</td>
								<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsVikar( "VikarID" ), rsVikar( "Navn"), "Vis vikar " & rsVikar( "Navn") )%></td>
								<td><%=rsVikar("Adresse") & ", " & rsVikar("postnr") & " " & rsVikar("PostSted")%></td>
								<td class="nowrap"><% =rsVikar("Telefon")%>&nbsp;</td>
								<td class="nowrap"><% =rsVikar("MobilTlf")%>&nbsp;</td>
								<td><a href="mailto:<% =rsVikar("Epost")%>"><font ID=fnt2<%=nummer%> class="groenn"><% =rsVikar("Epost")%>&nbsp;</a></td>
								<td class="center"><a href="hotlistVikar.asp?Slett=vikar&ID=<%=rsVikar("ID")%>"><img src="../Images/icon_delete.gif" alt="Slett"></a></td>
							</tr>
							<%
							rsVikar.MoveNext
						Wend 
						rsVikar.close
						set rsVikar = nothing 
						%>
					</table>
					<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="<%objResource.WriteText("Top")%>"><%objResource.WriteText("Top")%></a><br>&nbsp;
				</div>
			</div>
		</div>
	</body>
</html>
<%
set objResource = Nothing
CloseConnection(Conn)
set Conn = nothing		
%>