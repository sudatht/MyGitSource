<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim Conn
	dim FoundRecords : FoundRecords = false
	dim rsRapport
	
	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") <> "" Then
		' Add values from current page
		Fradato  = Request.Form( "tbxFradato" )
		Tildato  = Request.Form( "tbxTildato" )

		' Search value exists ?
		If Fradato <> "" And Tildato <> ""  Then

			if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
				AddErrorMessage("Fradato kan ikke være senere enn tildato!")
				call RenderErrorMessage()			
			end if

			' Get all Avdeling
			strSQL = "SELECT AvdelingID, Avdeling FROM Avdeling"
			set rsAvdeling = GetFirehoseRS(strSQL, Conn)

			' Move recordset to array
			ArrAvdeling = rsAvdeling.GetRows

			' Close recordset and release rs
			rsAvdeling.Close
			set rsAvdeling = Nothing

			' Get all
			strSQL = "SELECT V.VikarID, V.Fornavn, V.Etternavn, Ansvarlig = M.Fornavn + ' ' + M.Etternavn, V.Godkjentdato, V.GodkjentAv, VA.AvdelingID " &_
					"FROM VIKAR V, VIKAR_AVDELING VA, Medarbeider M " &_
					"WHERE V.Godkjentdato >= " & DbDate( fradato) &_
					" AND V.Godkjentdato <= " & DbDate( Tildato) &_
					" AND V.StatusID = 3 " &_
					" AND V.VikarID *=  VA.VikarID " &_
					" AND V.AnsMedID =  M.MedID " &_
					" ORDER BY  V.Etternavn, V.Fornavn"

			'Response.write strSQL
			Set rsRapport = GetFireHoseRS( strSQL, Conn )

			'Records found ?
			If (HasRows(rsRapport) = true) Then
				FoundRecords = true
			End If
		End If
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
		<script language="javascript" src="../Js/javascript.js" type="text/javascript"></script>
		<title>Nye vikarer</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Nye vikarer</h1>
			</div>
			<div class="content">
			<form name="formEn" ACTION="nyeVikarer.asp" METHOD="POST">
				<input type="hidden" NAME="tbxPageNo" VALUE="1">
				<table>
					<tr>
						<td>Fra dato:</td>
						<td><input NAME="tbxFraDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)" ID="Text1"> </td>
						<td>Til dato:</td>
						<td><input NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)" ID="Text2"> </td>
						<td><input type="submit" name="pbnDataAction" value="     Søk    "></td>
					</tr>
				</table>
			</form>
			<%
			' Create table only when records found
			If  (foundRecords = true)  Then
				' Create table
				Response.Write "<div class='listing'><table cellpadding='0' cellspacing='1'>"

				' Create table heading
				Response.Write "<tr>"
				Response.Write "<th>Navn</th>"
				Response.Write "<th>Godkjent dato</th>"
				Response.Write "<th>Godkjent av</th>"
				Response.Write "<th colspan='10'>Avdeling</th>"

				Response.Write "</tr>"
				Do Until rsRapport.EOF
					If rsRapport("VikarID") <> VikarID Then
						' Create row
						Response.Write "<tr>"						
						Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsRapport( "VikarID" ), rsRapport( "Etternavn") & " " & rsRapport( "Fornavn"), "Vis vikar " & rsRapport( "Etternavn") & " " & rsRapport( "Fornavn") ) & "</td>"
						Response.Write "<td>" & rsRapport("Godkjentdato") & "</td>"
						Response.Write "<td>" & rsRapport("Ansvarlig") & "</td>"
						VikarID = rsRapport("VikarID")
					End If
					Response.Write "<td>"
					For Counter = LBound( ArrAvdeling, 2 ) To UBound(  ArrAvdeling , 2 )
						If rsRapport("AvdelingID")  = ArrAvdeling( 0, Counter )  Then
							Response.Write ArrAvdeling( 1, Counter )
							Exit For
						End If
					Next
					Response.Write "&#160;</td>"

					' Get next record
					rsRapport.MoveNext
				Loop
				' Close recordset
				rsRapport.Close
				set rsRapport = Nothing
				' End table
				Response.Write "</table></div>"
			End If
			%>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>