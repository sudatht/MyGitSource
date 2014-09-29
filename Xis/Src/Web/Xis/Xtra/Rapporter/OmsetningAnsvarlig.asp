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
	dim Frabelop
	dim RecordsFound : RecordsFound = false

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") <> "" Then
		' Add values from current page
		
		if Request.Form( "tbxFradato" ) <> "" then
			Fradato = Request.Form( "tbxFradato" )
		else
			AddErrorMessage("Fradato må fylles ut!")
			call RenderErrorMessage()
		end if

		if Request.Form( "tbxTildato" ) <> "" then
			Tildato = Request.Form( "tbxTildato" )
		else
			AddErrorMessage("Tildato må fylles ut!")
			call RenderErrorMessage()
		end if
		
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

		' First time page called and search value exist ?
		If Fradato <> "" And Tildato <> ""  Then
			' Get database connection
			Set Conn = GetConnection(GetConnectionstring(XIS, ""))

			' Get all
			strSQL = "SELECT M.MedID, Omsetning=Sum(FA.Linjesum), M.Etternavn, M.Fornavn " &_
				"FROM FAKTURAGRUNNLAG FA, OPPDRAG O, MEDARBEIDER M " &_
				"WHERE FA.Fakturadato >= " & DbDate( fradato) &_
				" AND FA.Fakturadato <= " & DbDate( Tildato) &_
				" AND FA.OppdragID = O.OppdragID " &_
				" AND O.AnsMedID = M.MedID " &_
				" GROUP BY M.MedID, M.Etternavn, M.Fornavn " &_
				" HAVING Sum( FA.Linjesum ) >= " & Replace(Frabelop,",",".") &_
				" ORDER BY Omsetning Desc"

			set rsRapport = GetFirehoseRS(strSQL, Conn)
			
			'Records found ?
			If (HasRows(rsRapport) = True) Then
				RecordsFound = true
			else
				set rsRapport = nothing
				CloseConnection(Conn)
				set Conn = nothing			
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
		<script LANGUAGE="javaScript" src="/xtra/js/javaScript.js"></script>
		<title>Omsetning pr. ansvarlig medarbeider ( fakturert )</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Omsetnings pr. ansvarlig medarbeider ( fakturert )</h1>
			</div>
			<div class="content">			
				<form name="formEn" ACTION="omsetningAnsvarlig.asp" METHOD="POST">
					<input type="hidden" NAME="tbxPageNo" VALUE="1">
					<table>
						<tr>
							<td>Fra dato:</td>
							<td><input NAME="tbxFraDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"> </td>
							<td>Til dato:</td>
							<td><input NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text2"> </td>
							<td>Omsetning over kr:</td>
							<td><input NAME="tbxFrabelop" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Frabelop%>"></td>
							<td><input type="submit" name="pbnDataAction" value="     Søk    " onclick="dateInterval(this.form, this.name)"></td>
						</tr>
					</table>
				</form>
				<div class="listing">
					<%
					' Create table only when records found
					If  (RecordsFound = true) Then
						' Create table
						Response.Write "<table>"

						' Create table heading
						Response.Write "<tr>"
						Response.Write "<th>Ansvarlig</th>"
						Response.Write "<th>Omsetning</th>"
						Response.Write "</tr>"

						Do Until rsRapport.EOF
							' Create row
							Response.Write "<tr>"
							Response.Write "<td>" & rsRapport( "Etternavn")  & " " & rsRapport( "Fornavn") & "</td>"
							Response.Write "<TD class=right>" & FormatNumber( rsRapport("Omsetning"), 0 ) & "</td>"
							Response.Write "</tr>"
							' Get next record
							rsRapport.MoveNext
						Loop

						' Close recordset
						rsRapport.Close
						' Clear recordset
						set rsRapport = Nothing

						' End table
						Response.Write "</table>"
						CloseConnection(Conn)
						set Conn = nothing
					End If
					%>
				</div>
			</div>
		</div>
	</body>
</html>