<%@ Language="VbScript" %>
<%
option explicit
Response.Expires = 0
%>
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.SuperOffice.Integration.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If HasUserRight(ACCESS_CUSTOMER, RIGHT_READ) = false Then
		Response.Redirect("\xtra\IngenTilgang.asp")
	end if

	dim strSQL
	dim Conn
	dim AkseptSelected
	dim gammleMed
	dim rsContact
	dim StrSQLWhere
	dim rsAktivitet
	dim FirmaID
	dim strRegistereredBy
	dim rsRegistrertBy
	dim strSelected	
	dim strSQLRegisteredBy
	dim strFraDato
	dim strTilDato
	dim lngBrukerID

	lngbrukerID = Session("BrukerID")

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	' Check input values
	'Check is this is first time we display this page..
	If LenB(Request( "IsPostback")) = 0 Then
		'First time check for mandatory values
		' Do we have FirmaID ?
		If Request( "cuid" ) = "" Then
			AddErrorMessage("Systemfeil: FirmaID mangler!")
			call RenderErrorMessage()
		Else
			strSQL = "SELECT FirmaId FROM Firma WHERE SOCUID = " & request("cuid")
			set rsContact = GetFirehoseRS(strSQL, conn)
			if(HasRows(rsContact) = true) then
				FirmaID = rsContact("FirmaId")			
			end if
			rsContact.close
			set rsContact = nothing
		End If

		' Default value is 'Aksept'
		AkseptSelected = "CHECKED"
		gammleMed = "Ja"
	Else
		' Add values from current page
		FirmaID = Request( "firmaId" )
	End If	
	
	' Get the people that has registered activities on this oppdrag
	If (Request.Querystring( "dbxRegisteredBy" ) = "" OR Request.Querystring( "dbxRegisteredBy" ) = "Alle") Then
		strRegistereredBy = ""
		strSQLRegisteredBy = ""
	Else
		strRegistereredBy = cstr(Request.Querystring( "dbxRegisteredBy" ))
		strSQLRegisteredBy = " AND [A].[registrertavID]  = " & strRegistereredBy & " "
	End If

	' Add selection on Fradato ?
	If Request.querystring("tbxFradato") <> "" Then
		StrSQLWhere = strSql & " AND [A].[AktivitetDato] >= " & DbDate( Request.querystring("tbxFradato") )
		strFraDato = Request.querystring("tbxFradato")
	End If

	' Add selection on Tildato ?
	If Request.querystring("tbxTildato") <> "" Then
		' Add selection on tildato
		StrSQLWhere = StrSQLWhere & " AND [A].[AktivitetDato]  <= " & DbDate( Request.querystring("tbxTildato") )
		strTilDato = Request.querystring("tbxTildato")
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
		<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
		<script language="javaScript" type="text/javascript">
			var i=0;
			function focused(f)
			{
				i=f.substring(3,6);
				i=parseInt(i);
			}
		</script>
		<title>Aktiviteter</title>
	</head>
	<body>
	<div class="pageContainer" id="pageContainer">
		<div class="content">
			<form name="formEn" ID="Form1">
				<input type="hidden" name="firmaId" value="<%=FirmaID%>" ID="Hidden1">
				<input type="hidden" name="IsPostback" value="1" ID="Hidden2">
				Start dato: <input ID="lnk1" onFocus="focused(this.id)" NAME="tbxFradato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strFraDato%>">&nbsp;&nbsp;
				Slutt dato: <input ID="lnk2" onFocus="focused(this.id)" NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" ONBLUR="dateCheck(this.form, this.name)" value="<%=strTilDato%>">&nbsp;&nbsp;
				Registrert av:
				<select NAME="dbxRegisteredBy" ID="Select1">
					<option VALUE="Alle">Alle</option>
					<%
					strSQL ="SELECT DISTINCT [Bruker].[Navn], [Bruker].[MedarbID] " &_ 
						"FROM [Bruker] INNER JOIN [Aktivitet] ON [Aktivitet].[RegistrertAvID] = [Bruker].[MedarbID] " &_ 
						"WHERE [Aktivitet].[FirmaID] = " & FirmaID &_
						" AND ([Aktivitet].[VikarId] > 0 OR [Aktivitet].[OppdragID] > 0) " &_
						" ORDER BY [Bruker].[Navn]"
					set rsRegistrertBy = GetFirehoseRS(strSQL, Conn)
					While not rsRegistrertBy.EOF
						If cstr(rsRegistrertBy("MedarbID")) = cstr(strRegistereredBy) Then
							strSelected = " SELECTED"
						Else
							strSelected = ""
						End If
						%><option value='<%=rsRegistrertBy("MedarbID")%>'<%=strSelected%>><%=rsRegistrertBy("Navn")%></option>
						<%
						rsRegistrertBy.movenext
					wend
					rsRegistrertBy.close
					set rsRegistrertBy = nothing
					%>
					</select>
					<input type="submit" value="Søk" ID="Submit1" NAME="Submit1">
			</form>
			<%
			' Build sql-statement
			strSql = "SELECT [A].[AktivitetID], [A].[Aktivitetdato], [A].[AutoRegistered], [AT].[AktivitetType], [A].[FirmaID], [A].[VikarID], " & _
					" [A].[OppdragID], [V].[Fornavn] + ' ' + [V].[Etternavn] AS [Vikarnavn], CASE WHEN ([M].[fornavn]) IS NULL THEN [A].RegistrertAv ELSE ([M].[fornavn] + ' ' + [M].[etternavn]) END AS Navn, [A].[Notat] " & _
					" FROM [AKTIVITET] AS [A] " & _
					" INNER JOIN [H_AKTIVITET_TYPE] AS [AT] ON [A].[AktivitetTypeID] = [AT].[AktivitetTypeID] " & _
					" LEFT OUTER JOIN [medarbeider] AS [M] ON [A].[RegistrertAvID] = [M].[medid] " & _
					" LEFT OUTER JOIN [Vikar] AS [V] ON [A].[VikarID] = [V].[VikarID] " & _
 					" WHERE [A].[FirmaID] = " & FirmaID & _
					strSQLRegisteredBy & _
					StrSQLWhere & _
					" AND ([A].[VikarId] > 0 OR [A].[OppdragID] > 0) " &_
					" ORDER BY [A].[AktivitetDato] DESC "

			' Get all records
			set rsAktivitet = GetFirehoseRS(strSQL, Conn)

			' No records found ?
			If (hasRows(rsAktivitet)) Then
				' Create table only when records found
				%>
				<div class="listing">
				<table cellpadding='0' cellspacing='1' ID="Table1">
					<tr>
						<th>Dato</th>
						<th class="nowrap" nowrap>Reg. av</th>
						<th>Vikar</th>
						<th>Oppdrag</th>
						<th>Kommentar</th>
						<th>Slette</th>
					</tr>
				<%
				Do Until (rsAktivitet.EOF)
					' Create row
					Response.Write "<tr>"
					Response.Write "<td>" & rsAktivitet( "AktivitetDato") & "</td>"
					Response.Write "<td>" & rsAktivitet( "navn") & "</td>"
					' Show link only if value
					If  rsAktivitet( "VikarID" ) <> 0 Then
						
						Response.Write "<td class='nowrap'>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsAktivitet( "VikarID" ), rsAktivitet( "Vikarnavn"), "Vis vikar " & rsAktivitet( "Vikarnavn") ) & "</td>"
					Else
						Response.Write "<td>&nbsp;</td>"
					End If
					' Show link only if value
					If  rsAktivitet( "OppdragID" ) <> 0 Then
						Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & rsAktivitet( "OppdragID" ), rsAktivitet( "OppdragID"), "Vis oppdrag med oppdragnr " & rsAktivitet( "OppdragID") ) & "</td>"						
						'Response.Write "<td><A href='WebUI/OppdragView.aspx?OppdragID=" & rsAktivitet( "OppdragID" ) & "'>" & rsAktivitet( "OppdragID") & "</a></td>"
					Else
						Response.Write "<td>&nbsp;</td>"
					End If
					Response.Write "<td>" & rsAktivitet( "Notat" ) & "&nbsp;</td>"
					
					If (rsAktivitet("AutoRegistered") = 0) Then
						Response.Write "<td class='center'><a href='aktivitetDB.asp?source=kontaktaktiviteter&pbnDataAction=slette&tbxFirmaID=" & rsAktivitet( "FirmaID" ) & "&tbxVikarID=" & rsAktivitet( "VikarID" ) & "&tbxOppdragID=" & rsAktivitet( "OppdragID") & "&tbxAktivitetID=" & rsAktivitet( "AktivitetID" ) & "'><img src='/xtra/images/icon_delete.gif' alt='Slett' width='12' height='12' border='0'></a></td>"
					Else
						Response.Write "<td class='center'>&nbsp;</td>"
					End If
					Response.Write "</tr>"
				   rsAktivitet.MoveNext
				Loop
				response.write "</table>"
				rsAktivitet.close
			end if
			%>
			</div>
		</div>
	</div>
</body>
</html>
<%
CloseConnection(Conn)
set rsAktivitet = Nothing
set rsRegistrertBy = Nothing
set Conn = nothing		
%>