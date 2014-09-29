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

	dim lngBrukerID
	dim rsHotlist

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList
	dim strShowAuto	
	
	dim weekno
	dim split
	
	
	
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
	
	
	' Do we have Week No ?
	If Request("weekno") = "" Then
		weekno = 0
	Else
		weekno = Request("weekno")
	End If
	
	' Do we have Split value ?
	If Request("split") = "" Then
		split = 0
	Else
		split = Request("split")
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

	StrSQLWhere = StrSQLWhere & " AND [A].[AutoRegistered] = 1"
	
	
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
		<title>Activity Timesheet Log</title>
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
				<h1>Timeliste logg</h1>
			</div>
			<div class="content">
				
				<form name="formEn" ACTION="aktivitetOppdrag.asp">
				
				<table>				
				<tr>				
					<input TYPE="HIDDEN" NAME="OppdragID" VALUE="<%=lngOppdragID %>" ID="Hidden1">
					<input TYPE="HIDDEN" NAME="VikarID" VALUE="<%=VikarID%>" ID="Hidden2">
				
				</tr>
				</table>	
					
					
					
				</form>
				<%
				' Build sql-statement
				strSql = "SELECT [A].[AktivitetID], [A].[Aktivitetdato], [AT].[AktivitetType], [A].[FirmaID], [F].[Firma], [F].[SOCuID], [A].[VikarID], " & _
					" [A].[OppdragID], [V].[Fornavn] + ' ' + [V].[Etternavn] AS [Vikarnavn], CASE WHEN ([M].[fornavn]) IS NULL THEN [A].RegistrertAv ELSE ([M].[fornavn] + ' ' + [M].[etternavn]) END AS Navn, [A].[Notat] " & _
					" FROM [AKTIVITET] AS [A] " & _
					" INNER JOIN [H_AKTIVITET_TYPE] AS [AT] ON [A].[AktivitetTypeID] = [AT].[AktivitetTypeID] " & _
					" LEFT OUTER JOIN [medarbeider] AS [M] ON [A].[RegistrertAvID] = [M].[medid] " & _
					" LEFT OUTER JOIN [Vikar] AS [V] ON [A].[VikarID] = [V].[VikarID] " & _
					" LEFT OUTER JOIN [Firma] AS [F] ON [A].[FirmaID] = [F].[FirmaID] " & _
					" WHERE [A].[OppdragID] = " & lngOppdragID  & " AND [A].WeekNo = " & weekno & " AND [A].Split = " & split & " " & _
					strSQLType & _
					strSQLRegisteredBy & _
					StrSQLWhere & _
					" ORDER BY [A].[AktivitetDato] DESC "
					
					
				
				'response.write strSql
				'nnn
				
					
				' Get all records
				set rsAktivitet = GetFirehoseRS(StrSQL, Conn)
				'response.write StrSQL
				'cccc
				
				if (HasRows(rsAktivitet) = true) then
					%>
					<div class="listing">
						<table cellpadding='0' cellspacing='1'>
							<tr>
								<th>Dato</th>
								<th>Reg. av</th>								
								<th>Kommentar</th>								
							</tr>
							<%
							Do Until rsAktivitet.EOF
								' Create row
								Response.Write "<tr>"
								Response.Write "<td>" & rsAktivitet( "AktivitetDato") & "</td>"
								Response.Write "<td>" & rsAktivitet( "Navn") & "</td>"								
								Response.Write "<td>" & rsAktivitet( "Notat" ) & "</td>"								
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