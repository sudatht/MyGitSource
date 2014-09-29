<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<%

	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if' Create Sql search string

	dim strSQL
	dim description
	dim strKunde

	' Check datatype for vikarid
	If Request.form("tbxVikarID") <> "" Then
		If Not IsNumeric(Request.form("tbxVikarID") ) Then
			AddErrorMessage("Ulovlig verdi. vikarid må være numerisk!")
		End If
	End If

	' Check datatype for ansattnummer
	If Request.form("tbxAnsattnr") <> "" Then
		If Not IsNumeric(Request.form("tbxAnsattnr") ) Then
			AddErrorMessage("Ulovlig verdi. Ansattnummer må være numerisk!")
		End If
	End If

	' Check datatype for Oppdragnr
	If Request.form("tbxOppdragID") <> "" Then
		If Not IsNumeric(Request.form("tbxOppdragID") ) Then
			AddErrorMessage("Ulovlig verdi. Oppdragnr må være numerisk!")
		End If
	End If

	' Check datatype for Oppdragnr
	If Request.form("tbxFirmaID") <> "" Then
		If Not IsNumeric(Request.form("tbxFirmaID") ) Then
			AddErrorMessage("Ulovlig verdi. Kontaktnr må være numerisk!")
		End If
	End If

	if(HasError() = true) then
		call RenderErrorMessage()
	end if

	' Add selection on status
	If Request.form("dbxStatus") > 0 Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & " " & "[O].[StatusID] = " & Request.form("dbxStatus")
	End If

	' Add selection on Ansvarlig
	If Request.form("dbxMedarbeider") > 0 Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & " " & "[O].[AnsMedID] = " & Request.form("dbxMedarbeider")
	End If

	' Add selection on Avdelingskontor
	If (Request.form("dbxAvdeling")> 0) Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & " [O].[Avdelingskontorid] = " & Request.form("dbxAvdeling")
	End If

	' Add selection on tjenesteområde
	If Request.form("dbxtjenesteomrade") > 0 Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & " [O].[tomID] = " & Request.form("dbxtjenesteomrade")
	End If

	' Add vikartype
	If Request.Form("dbxType") <> 0 Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & " " & "[O].[TypeID] = " & Request.form("dbxType")
	End If

	' Add selection on Fradato ?
	If Request.Form("tbxFradato") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If

		' Add selection on tildato
		strSQL = strSQL & "[Fradato] >= " & DbDate( Request.form("tbxFradato") )
	End If

	' Add selection on Tildato ?
	If Request.Form("tbxTildato") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		' Add selection on tildato
		strSQL = strSQL & "[Tildato]  <= " & DbDate( Request.form("tbxTildato") )
	End If

	' Add selection on Etternavn ?
	If Request.Form("tbxEtternavn") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & "[OppdragId] IN (SELECT [oppdragid] FROM [oppdrag_vikar], [vikar] WHERE [oppdrag_vikar].[vikarid] = [vikar].[vikarid] AND [Etternavn] LIKE '" & Request.form("tbxEtternavn") & "%'" & " ) "
	End If

	' Add selection on ID ?
	If Request.Form("tbxVikarID") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		' Add selection on ID
		strSQL = strSQL & "[OppdragId] IN (SELECT [oppdragid] FROM [oppdrag_vikar] WHERE [vikarid] = " & Request.form("tbxVikarID") & " ) "
	End If

	If Request.Form("tbxAnsattnr") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		' Add selection on ID
		strSQL = strSQL & "[OppdragId] IN ( SELECT [OPPDRAG_VIKAR].[Oppdragid] FROM [OPPDRAG_VIKAR] LEFT OUTER JOIN [VIKAR_ANSATTNUMMER] ON [OPPDRAG_VIKAR].[Vikarid] = [VIKAR_ANSATTNUMMER].[Vikarid] WHERE [VIKAR_ANSATTNUMMER].[ansattnummer] = '" & Request.form("tbxAnsattnr") & "' ) "
	End If

	' Add selection on Firma ?
	If Request.Form("tbxFirma") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		strSQL = strSQL & "[F].[Firma] LIKE '" & Request.form("tbxFirma") & "%'"
	End If

	' Add selection on FirmaID ?
	If Request.Form("tbxFirmaID") <> "" Then
		If strSQL <> "" Then
			strSQL = strSQL & " AND "
		End If
		' Add selection on ID
		strSQL =  strSQL & "[O].[FirmaID] = " & Request.form("tbxFirmaID")
	End If
	
	' Add selection on Description ?
	description = Trim(Request.Form("tbxDescription"))
	If description <> "" Then
		description = Replace(description,"%","")
		description = "%" & description & "%"
		
		If strSQL <> "" Then
			strSQL = strSQL & " AND " 
		End If
		' Add selection on ID
		strSQL =  strSQL & "[O].[Beskrivelse] like '" & description & "'"
	End If

	' Add selection on OppdragID ?
	If Request.Form("tbxOppdragID") <> "" Then
		if not isnull(Request.form("tbxOppdragID")) then
   			' ID is unique number. Remove all other selections
   			strSQL = ""
   			' Add selection on ID
   			strSQL = "[OppdragID] = " & Request.form("tbxOppdragID")
   		end if
	End If

	If strSQL <> "" Then
		strSQL = strSQL & " AND "
	End If

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Do search
	strSQL = "SELECT TOP 100 [O].[OppdragId], [O].[Oppdragskode], [O].[Beskrivelse], [F].[FirmaID], [F].[Firma], isnull([F].CRMAccountGuid,'') AS CRMAccountGuid ,[F].[SOCuID],  [O].[Fradato], [O].[TilDato], [S].[OppdragsStatus] " &_
			" FROM [Oppdrag] AS [O], [Firma] AS [F],  [h_oppdrag_status] AS [S] " &_
			" WHERE " & strSQL &_
			" [O].[FirmaID] = [F].[FirmaID] " &_
			" AND [O].[StatusID] = [S].[Oppdragsstatusid] " &_
			" ORDER BY [F].[Firma] DESC"

	Set rsOppdrag = GetFirehoseRS(strSQL, Conn)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
		<title>Oppdrag Resultat</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>			
		<script language="javaScript" type="text/javascript">
			function shortKey(e) 
			{	
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
				
				if (event.keyCode == 13)
				{
					document.all.frmJobSearch.submit();
				}				
				if (modKey && modKey2 && keyChar=="S")
				{	
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
			}
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;			
		</script>
	</head>
	<base TARGET="_self">
	<body onLoad="fokus()">
		<div class="pageContainer" id="pageContainer">
			<a id="Top"></a>
			<div class="contentHead1">
				<h1>Søkeresultat oppdrag</h1>
			</div>
			<div class="content">
				<div class="listing">
					<table cellspacing="1" cellpadding="0" ID="Table1">
						<tr>
							<th>Opp.nr.</th>
							<th>Beskrivelse</th>
							<th>Kontakt</th>
							<th>Startdato</th>
							<th>Sluttdato</th>
							<th>Status</th>
							<th>Ansattnummer</th>
							<th>Vikar</th>
						</tr>
						<%
						Do Until (rsOppdrag.EOF)
							%>
							<tr>
								<td class="right"><% =rsOppdrag("OppdragID") %></td>
								<td><a title="Vis oppdrag" href="WebUI/OppdragView.aspx?OppdragID=<%=rsOppdrag("OppdragID") %>"><% =rsOppdrag("Beskrivelse") %></a></td>
								
								<%
									linkurl = Application("CRMAccountLink") & rsOppdrag("CRMAccountGuid") & "%7d&pagetype=entityrecord"
									strKunde = "<a href=" & linkurl & " target='_blank'>" & rsOppdrag("Firma").Value & " </a>"							
								%>
								
								<td><%=strKunde%></td>							
								<td><% =rsOppdrag("Fradato")%></td>
								<td><% =rsOppdrag("TilDato")%></td>
								<td><% =rsOppdrag("OppdragsStatus")%></td>
								<%
								strSQL = "SELECT OPPDRAG_VIKAR.VikarID, VIKAR.Etternavn, VIKAR_ANSATTNUMMER.ansattnummer " &_
								"FROM OPPDRAG_VIKAR " & _
								"LEFT OUTER JOIN VIKAR ON VIKAR.Vikarid = OPPDRAG_VIKAR.Vikarid " & _
								"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
								"WHERE OPPDRAG_VIKAR.StatusID = 4 " &_
								"AND OPPDRAG_VIKAR.OppdragID = " & rsOppdrag("OppdragID")

								set rsNavn = GetFirehoseRS(strSQL, conn)
								If (HasRows(rsNavn) = true) Then
									%>
									<td class="right"><%=rsNavn("ansattnummer")%>&nbsp;</td>
									<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsNavn( "VikarID" ), rsNavn( "Etternavn"), "Vis vikar " & rsNavn( "Etternavn") )%></td>
									<% 
								else
									%>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<% 		
								End If
								rsNavn.Close
								set rsNavn = Nothing
								%>
							</tr>
							<%
							rsOppdrag.MoveNext
						Loop
						rsOppdrag.Close
						Set rsOppdrag = Nothing
						%>
					</table>
				</div>
				<a href="#top"><img src="./Images/icon_GoToTop.gif" alt="Til toppen">Til toppen</a><br>&nbsp;	
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>