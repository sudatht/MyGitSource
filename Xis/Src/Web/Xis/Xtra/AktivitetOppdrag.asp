<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%

	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if	
	
	dim strSQL
	dim StrSQLWhere
	dim rsAktivitet
	dim VikarID
	dim lngOppdragID
	dim strSQLType
	dim lngActivityType
	dim strRegistereredBy
	dim strSQLRegisteredBy
	dim strAktivitetsFraDato
	dim strAktivitetsTilDato
	dim strFraDato
	dim strTilDato
	dim strKunde

	dim lngBrukerID
	dim rsHotlist

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList
	dim strShowAuto
	
	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = "oppdrag"

	lngBrukerID = Session("BrukerID")

	strShowAuto = Request.Querystring( "chkShowAutoRegister" )
	
	' Do we have OppdragID ?
	If Request("OppdragID") = "" Then
		AddErrorMessage("Feil:Parameter for oppdragid mangler!")
		call RenderErrorMessage()		
	Else
		lngOppdragID  = Request("OppdragID")
	End If

	' Do we have VikarID ?
	If Request("VikarID") = "" Then
		VikarID = 0
	Else
		VikarID = Request("VikarID")
	End If


	' Get the people that has registered activities on this oppdrag
	If (Request.Querystring( "dbxRegisteredBy" ) = "" OR Request.Querystring( "dbxRegisteredBy" ) = "Alle") Then
	   strRegistereredBy = ""
	   strSQLRegisteredBy = ""
	Else
	   strRegistereredBy = cstr(Request.Querystring( "dbxRegisteredBy" ))
	   strSQLRegisteredBy = " AND [A].[registrertavID]  = " & strRegistereredBy & " "
	End If

	' is there an activity selected
	If (Request.Querystring( "dbxActivityType" ) = "") OR (Request.Querystring( "dbxActivityType" ) = 0) Then
	   lngActivityType = 0
	   strSQLType = ""
	Else
	   lngActivityType = clng(Request.Querystring( "dbxActivityType" ))
	   strSQLType = " AND [A].[AktivitetTypeID] =" & lngActivityType & " "
	End If

	' Add selection on Fradato ?
	If Request.querystring("tbxFradato") <> "" Then
	   StrSQLWhere = strSql & " AND [A].[AktivitetDato] >= " & DbDate( Request.querystring("tbxFradato") )
	   strAktivitetsFraDato = Request.querystring("tbxFradato")
	End If

	'on/off auto registered activity list
	If strShowAuto <> "showAuto" Then
		StrSQLWhere = StrSQLWhere & " AND [A].[AutoRegistered] = 0"
	End If
	
	' Add selection on Tildato ?
	If Request.querystring("tbxTildato") <> "" Then
	   ' Add selection on tildato
	   StrSQLWhere = StrSQLWhere & " AND [A].[AktivitetDato]  <= " & DbDate( Request.querystring("tbxTildato") )
	   strAktivitetsTilDato = Request.querystring("tbxTildato")
	End If

	
	
	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	strSQL = "SELECT [Beskrivelse], [FirmaID], [FraDato], [TilDato] FROM [OPPDRAG] WHERE [OppdragID] = " & lngOppdragID
	set rsOppdrag = GetFirehoseRS(StrSQL, Conn)

	Oppdrag = rsOppdrag( "Beskrivelse" )
	lngFirmaID = rsOppdrag("FirmaID" )
	strFraDato = rsOppdrag( "fraDato" )
	strTilDato = rsOppdrag( "tilDato" )

	rsOppdrag.Close
	set rsOppdrag = Nothing

	strSQL = "SELECT * FROM [HOTLIST] WHERE [Status] = 1 AND [BrukerID] =" & lngBrukerID & " AND [oppdragID] =" & lngOppdragID
	set rsHotlist = GetFirehoseRS(StrSQL, Conn)
	if (HasRows(rsHotlist) = true) then
		blnShowHotList = true
		strAddToHotlistLink = "AddHotlist.asp?kode=1&amp;oppdragID=" & lngOppdragID & "&amp;kundeNavn=" & server.URLEncode(strFirma) & "&amp;kundeNr=" & lngFirmaID
		strHotlistType = "oppdrag"
	end if
	rsHotlist.close
	set rsHotlist = nothing

	'Top Menu init
	dim strALinkStart
	dim strALinkEnd
	dim AToolbarAtrib(6,3) '0 = Enable/Disable, 1 = form/Link to activate, 2 = close link form string

	'Lagre Oppdrag
	AToolbarAtrib(0,0) = "0"
	AToolbarAtrib(0,1) = ""
	AToolbarAtrib(0,2) = ""
	'Vise oppdrag
	AToolbarAtrib(1,0) = "1"
	AToolbarAtrib(1,1) = "<a href='WebUI/OppdragView.aspx?OppdragID=" & lngOppdragID & "' title='Vis oppdrag'>"
	AToolbarAtrib(1,2) = "</a>"
	'Til Endre oppdrag
	AToolbarAtrib(2,0) = "1"
	AToolbarAtrib(2,1) = "<a href='oppdragNy.asp?OppdragID=" & lngOppdragID & "' title='Endre oppdrag'>"
	AToolbarAtrib(2,2) = "</a>"
	'Tilknytte konsulent
	AToolbarAtrib(3,0) = "1"
	AToolbarAtrib(3,1) =  "<form action='VikarSoek.asp?kurskode=" & strKurskode & "' METHOD='POST'  name='frmAddConsultant'>" & _
	"<input name='hdnPosted' TYPE='hidden' VALUE='1'>" & _
	"<input name='tbxOppdragID' TYPE='hidden' VALUE='" & lngOppdragID & "'>" & _
	"<input name='tbxFirmaID' TYPE='hidden' VALUE='" & lngFirmaID & "'>" & _
	"<a href='javascript:document.all.frmAddConsultant.submit();' title='Tilknytt vikar'>"
	AToolbarAtrib(3,2) = "</a></form>"
	'Aktiviteter for oppdrag
	AToolbarAtrib(4,0) = "0"
	AToolbarAtrib(4,1) = ""
	AToolbarAtrib(4,2) = ""
	'Kalender
	AToolbarAtrib(5,0) = "1"
	AToolbarAtrib(5,1) = "<form action='Kalender.asp?OppdragID=" & lngOppdragID & "' METHOD='POST' name='frmJobCalendar'></form>" & _
	"<a href='javascript:document.all.frmJobCalendar.submit();' title='Vis kalender for oppdrag'>"
	AToolbarAtrib(5,2) = "</a></form>"
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Aktiviteter tilknyttet oppdrag</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" src="Js/javascript.js"></script>
		<script type="text/javascript" src="js/contentMenu.js"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>		
		<script language="javaScript" type="text/javascript">
			function shortKey(e) 
			{	
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
								
				if (modKey && modKey2 && keyChar=="S")
				{	
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
				<% 
				If HasUserRight(ACCESS_TASK, RIGHT_WRITE) Then
					%>				
					if (modKey && modKey2 && keyChar=="Y")
					{	
						parent.frames[funcFrameIndex].location=("/xtra/OppdragNy.asp");
					}
					<% 
				End If 
				%>				
			}

			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;		
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Aktiviteter tilknyttet oppdrag</h1>
				<div class="contentMenu">
					<!--#include file="Includes/Top_Menu_job2.asp"-->
				</div>
			</div>
			<div class="content">
				
				<table>
				<tr>
				<td style="width:150px;">
			
				<%
				If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
					%>				
					<form name='formTo' ACTION='AktivitetNy.asp?OppdragID=<%=lngOppdragID %>&FirmaID=<%=lngFirmaID%>&VikarID=<%=VikarID%>' METHOD="POST">
						<input name="source" type="HIDDEN" value="oppdrag">
						<input NAME="tbxOppdragID" Type="HIDDEN" Value="<%=lngOppdragID %>">
						<input NAME="pbnDataAction" TYPE="SUBMIT" VALUE='Ny aktivitet'>
					</form>
					<%
				end if
				%>
				</td>
				
				<form name="formEn" ACTION="aktivitetOppdrag.asp">
				
				<td style="width:150px;"></td>
				<td style="width:100px;">Vis auto aktiviteter:</td>
				<td>
				
					<%if (strShowAuto="showAuto") then
						%>
						<input class="checkbox" type="checkbox" id="chkShowAutoRegister" name="chkShowAutoRegister" value="showAuto" onclick="submit()" checked/>	
						<%
					else
						%>
						<input class="checkbox" type="checkbox" id="chkShowAutoRegister" name="chkShowAutoRegister" value="showAuto" onclick="submit()" />
						<%
					end if
					%>
				</td>
				</tr>
				<tr>				
					<input TYPE="HIDDEN" NAME="OppdragID" VALUE="<%=lngOppdragID %>" ID="Hidden1">
					<input TYPE="HIDDEN" NAME="VikarID" VALUE="<%=VikarID%>" ID="Hidden2">
				<td style="width:150px;">
					Start dato: <input ID="lnk1" onFocus="focused(this.id)" NAME="tbxFradato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strAktivitetsFraDato%>">&nbsp;&nbsp;
				</td>
				<td style="width:150px;">	
					Slutt dato: <input ID="lnk2" onFocus="focused(this.id)" NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strAktivitetsTilDato%>">&nbsp;&nbsp;
				</td>
				<td style="width:100px;">	
					Aktivitetstype: 
				</td>
				<td>	
					<select NAME="dbxActivityType" ID="Select1">
						<option VALUE="0">Alle typer</option>
						<%
						strSQL = "SELECT [AktivitetTypeID], [AktivitetType] FROM [H_AKTIVITET_TYPE] ORDER BY AutoRegister,2"
						set rsAktivitet = GetFirehoseRS(StrSQL, Conn)
						While not rsAktivitet.EOF
							If clng(rsAktivitet("AktivitetTypeID")) = clng(lngActivityType) Then
								strSelected = " SELECTED"
							Else
								strSelected = ""
							End If

							response.write "<option VALUE='" & rsAktivitet("AktivitetTypeID") & "' " & strSelected & ">" & rsAktivitet("AktivitetType") & "</option>"
							rsAktivitet.movenext
						wend
						rsAktivitet.close
						set rsAktivitet = nothing
						%>
					</select>
				</td>
				<td>
					&nbsp;Registrert av:
					<select NAME="dbxRegisteredBy">
						<option VALUE="Alle">Alle </option>
						<%
						strSql ="SELECT DISTINCT [Bruker].[Navn], [Bruker].[MedarbID] FROM [Bruker] INNER JOIN [Aktivitet] ON [Aktivitet].[RegistrertAvID] = [Bruker].[MedarbID] WHERE [Aktivitet].[oppdragid] = " & lngOppdragID  & " ORDER BY [Bruker].[Navn]"
						set rsRegistrertBy = GetFirehoseRS(StrSQL, Conn)
						While not rsRegistrertBy.EOF
							If cstr(rsRegistrertBy("MedarbID")) = cstr(strRegistereredBy) Then
								strSelected = " SELECTED"
							Else
								strSelected = ""
							End If

							response.write "<option VALUE='" & rsRegistrertBy("MedarbID") & "' " & strSelected & ">" & rsRegistrertBy("Navn") & "</option>"
							rsRegistrertBy.movenext
						wend
						rsRegistrertBy.close
						set rsRegistrertBy = nothing
						%>
					</select>
					<input TYPE="SUBMIT" VALUE="Søk">
				</td>
				</tr>
				</table>	
					
					
					
				</form>
				<%
				' Build sql-statement
				strSql = "SELECT [A].[AktivitetID], [A].[Aktivitetdato], [A].[AutoRegistered], [AT].[AktivitetType], [A].[FirmaID], [F].[Firma], isnull([F].CRMAccountGuid,'') AS CRMAccountGuid, [F].[SOCuID], [A].[VikarID], " & _
					" [A].[OppdragID], [V].[Fornavn] + ' ' + [V].[Etternavn] AS [Vikarnavn], CASE WHEN ([M].[fornavn]) IS NULL THEN [A].RegistrertAv ELSE ([M].[fornavn] + ' ' + [M].[etternavn]) END AS Navn, [A].[Notat] " & _
					" FROM [AKTIVITET] AS [A] " & _
					" INNER JOIN [H_AKTIVITET_TYPE] AS [AT] ON [A].[AktivitetTypeID] = [AT].[AktivitetTypeID] " & _
					" LEFT OUTER JOIN [medarbeider] AS [M] ON [A].[RegistrertAvID] = [M].[medid] " & _
					" LEFT OUTER JOIN [Vikar] AS [V] ON [A].[VikarID] = [V].[VikarID] " & _
					" LEFT OUTER JOIN [Firma] AS [F] ON [A].[FirmaID] = [F].[FirmaID] " & _
					" WHERE [A].[OppdragID] = " & lngOppdragID  & " " & _
					strSQLType & _
					strSQLRegisteredBy & _
					StrSQLWhere & _
					" ORDER BY [A].[AktivitetDato] DESC "
					
					
					'response.write StrSQL
					'mmmm
					
				' Get all records
				set rsAktivitet = GetFirehoseRS(StrSQL, Conn)
				if (HasRows(rsAktivitet) = true) then
					%>
					<div class="listing">
						<table cellpadding='0' cellspacing='1'>
							<tr>
								<th>Dato</th>
								<th>Reg. av</th>
								<th>Type</th>
								<th>Kontakt</th>
								<th>Vikar</th>
								<th>Oppdrag</th>
								<th>Kommentar</th>
								<%
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									%>
									<th>Slette</th>
									<%
								end if
								%>
							</tr>
							<%
							Do Until rsAktivitet.EOF
								' Create row
								Response.Write "<tr>"
								Response.Write "<td>" & rsAktivitet( "AktivitetDato") & "</td>"
								Response.Write "<td>" & rsAktivitet( "Navn") & "</td>"
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									If (rsAktivitet("AutoRegistered") = 0) Then
										Response.Write "<td><A href='AktivitetNy.asp?AktivitetID=" & rsAktivitet( "AktivitetID" ) & "&source=oppdrag&FirmaID=" &rsAktivitet( "FirmaID") & "&VikarID=" & rsAktivitet( "VikarID") & "&OppdragID=" &rsAktivitet( "OppdragID") & "'>" & rsAktivitet( "AktivitetType") & "</a></td>"
									Else
										Response.Write "<td>" & rsAktivitet( "AktivitetType") & "</td>"
									End If
								Else
									Response.Write "<td>" & rsAktivitet( "AktivitetType") & "</td>"
								end if								
								' Show link only if value
								linkurl = Application("CRMAccountLink") & rsAktivitet("CRMAccountGuid") & "%7d&pagetype=entityrecord"
								
								If  rsAktivitet("FirmaID") <> 0 Then
									strKunde = "<a href=" & linkurl & " target='_blank'>" & rsAktivitet("Firma").Value & " </a>"									
								Else
									strKunde = rsAktivitet("Firma").Value
								End If
								
								Response.Write "<td>" & strKunde & "</td>"
								
								

								' Show link only if value
								If  rsAktivitet( "VikarID" ) <> 0 Then
									Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsAktivitet( "VikarID" ), rsAktivitet( "Vikarnavn"), "Vis vikar " & rsAktivitet( "Vikarnavn") ) & "</td>"
								Else
									Response.Write "<td>&nbsp;</td>"
								End If

								' Show link only if value
								If  rsAktivitet( "OppdragID" ) <> 0 Then
									Response.Write "<td><A href='WebUI/OppdragView.aspx?OppdragID=" & rsAktivitet( "OppdragID" ) & "'>" & rsAktivitet( "OppdragID") & "</a></td>"
								Else
									Response.Write "<td>&nbsp;</td>"
								End If
								Response.Write "<td>" & rsAktivitet( "Notat" ) & "</td>"
								If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
									If (rsAktivitet("AutoRegistered") = 0) Then
										Response.Write "<td><A href='AktivitetDB.asp?source=oppdrag&pbnDataAction=slette&tbxFirmaID=" & rsAktivitet( "FirmaID" ) & "&tbxVikarID=" & rsAktivitet( "VikarID" ) & "&tbxOppdragID=" & rsAktivitet( "OppdragID") & "&tbxAktivitetID=" & rsAktivitet( "AktivitetID" ) & "&chkShowAutoRegister=" & strShowAuto & "'><img src='images/icon_delete.gif' alt='Slette'></a></td>"
									Else
										Response.Write "<td class='center'>&nbsp;</td>"
									End If
								end if
								Response.Write "</tr>"
								rsAktivitet.MoveNext
							Loop
							%>
						</table>
					</div>
					<%
				end if
				' Close recordset
				rsAktivitet.Close
				set rsAktivitet = Nothing
				%>
				
				<br>&nbsp;
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>