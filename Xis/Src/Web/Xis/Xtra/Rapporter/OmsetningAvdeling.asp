<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim strSQL
	dim rsRapport
	dim Conn
	dim foundRecords : foundRecords = false

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") <> "" Then
		' Add values from current page
		Fradato = Request.Form( "tbxFradato" )
		Tildato = Request.Form( "tbxTildato" )
		Frabelop = Trim(Request.Form("tbxFrabelop"))

		If Request.Form( "tbxFrabelop" ) = "" Then
			Frabelop = 0
		Else
			If IsNumeric(Frabelop) then
			  	Frabelop = Frabelop
			else
			  	AddErrorMessage("Skriv inn et gyldig nummer!")
				call RenderErrorMessage()
			end if	
		End If
	   
		' Search value exist ?
		If Fradato <> "" AND Tildato <> ""  Then

			if (Fradato <> "" AND Tildato <> "") then
				if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
					AddErrorMessage("Fradato kan ikke være senere enn tildato!")
					call RenderErrorMessage()
				end if
			end if

			' Open database connection
			Set Conn = GetConnection(GetConnectionstring(XIS, ""))

			' Get all
			strSQL = "SELECT A.AvdelingID, Omsetning=Sum(FA.Linjesum), A.Avdeling " &_
				"FROM FAKTURAGRUNNLAG FA, OPPDRAG O, AVDELING A " &_
				"WHERE FA.Fakturadato >= " & DbDate( fradato) &_
				" AND FA.Fakturadato <= " & DbDate( Tildato) &_
				" AND FA.OppdragID = O.OppdragID " &_
				" AND O.AvdelingID = A.AvdelingID " &_
				" GROUP BY A.AvdelingID, A.Avdeling " &_
				" HAVING SUM( FA.Linjesum ) >= " & Replace(Frabelop,",",".") &_
				" ORDER BY Omsetning Desc"

			set rsRapport = GetFirehoseRS(strSQL, Conn)

			'Records found ?
			If (HasRows(rsRapport) = true)  Then
				FoundRecords = true
			End If
		end if
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
		<title>Omsetning pr. avdeling ( fakturert )</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Omsetning pr avdeling ( fakturert )</h1>
			</div>
			<div class="content">
				<form name="formEn" ACTION="omsetningAvdeling.asp" METHOD="POST">
					<input type="hidden" NAME="tbxPageNo" VALUE="1">
					<table>
						<tr>
							<td>Fra dato:</td>
							<td><input NAME="tbxFraDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)"> </td>
							<td>Til dato:</td>
							<td><input NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)"> </td>
							<td>Omsetning over kr:</td>
							<td><input NAME="tbxFrabelop" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Frabelop%>"> </td>
							<td><input type="submit" name="pbnDataAction" value="     Søk    "></td>
						</tr>
					</table>
				</form>
				<%
				' Create table only when records found
				If  (foundRecords)  Then
					' Create table
					Response.Write "<div class='listing'><table cellpadding='0' cellspacing='0'>"

					' Create table heading
					Response.Write "<tr>"
					Response.Write "<th>Avdeling</th>"
					Response.Write "<th>Omsetning</th>"
					Response.Write "</tr>"

					Do Until rsRapport.EOF
						' Create row
						Response.Write "<tr>"
						Response.Write "<td>" & rsRapport( "Avdeling") & "</td>"
						Response.Write "<TD class=right>" & FormatNumber( rsRapport("Omsetning"), 0 ) & "</td>"
						Response.Write "</tr>"
						' Get next record
						rsRapport.MoveNext
					Loop
					' Close recordset
					rsRapport.Close
					' Clear recordset
					set rsRapport = Nothing
					Response.Write "</table>"
				End If
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>