<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim foundRecords : foundRecords = false
	dim strSQL
	dim Conn
	dim teller
	dim fraMaaned
	dim tilMaaned
	dim selected
	dim periodeTilSQL
	dim egneSQL
	dim EgneAvkrysset
	dim AMonths(12)

	AMonths(0) = "(Velg måned)"
	AMonths(1) = "Januar"
	AMonths(2) = "Februar"
	AMonths(3) = "Mars"
	AMonths(4) = "April"
	AMonths(5) = "Mai"
	AMonths(6) = "Juni"
	AMonths(7) = "Juli"
	AMonths(8) = "August"
	AMonths(9) = "September"
	AMonths(10) = "Oktober"
	AMonths(11) = "November"
	AMonths(12) = "desember"
	
	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") <> "" Then
		' Add values FROM current page
		fraMaaned  = Request.Form( "tbxFraManed" )
		tilMaaned  = Request.Form( "tbxTilManed" )
	   
		if (fraMaaned = 0 and tilMaaned = 0) then
			AddErrorMessage("Du må velge en periode!")
			call RenderErrorMessage()
		end if

		if (fraMaaned = 0 and tilMaaned > 0) then				
			fraMaaned = tilMaaned
			tilMaaned = 0
		end if

		if (tilMaaned > 0) then				
			periodeFraSQL = " AND DATEPART (month, foedselsdato) >= " & fraMaaned 
			periodeTilSQL = " AND DATEPART (month, foedselsdato) <= "  & tilMaaned
		else
			periodeFraSQL = " AND DATEPART (month, foedselsdato) = " & fraMaaned 
			periodeTilSQL = ""
		end if

		if(Request("chkKunEgne") = "1") then
			EgneAvkrysset = "checked"
			egneSQL = " AND AnsMedID = " & Session("medarbID")
		else
			EgneAvkrysset = ""
			egneSQL = ""
		end if
		
		' Get all
		strSQL = "SELECT VikarID, Fornavn, Etternavn, Foedselsdato " &_
			" FROM VIKAR " &_
			" WHERE 1 = 1 " &_
			periodeFraSQL &_
			periodeTilSQL &_
			egneSQL &_
			" AND StatusID = 3 " &_
			" ORDER BY DATEPART (month, foedselsdato), DATEPART (day, foedselsdato)"

		Set rsRapport = getFirehoseRS( strSQL, Conn )

		' No records found ?
		foundRecords = hasRows(rsRapport)
	else
		fraMaaned = Month(now())
		tilMaaned = 0
		EgneAvkrysset = "checked"
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language="javaScript" src="../Js/javaScript.js"></script>
		<title>Fødselsdager</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Fødselsdager</h1>
			</div>
			<div class="content">
			<form name="formEn" ACTION="foedselsdager.asp" METHOD="POST">
				<input type="hidden" NAME="tbxPageNo" VALUE="1">
				<table>
					<tr>
						<td>Fra måned:</td>
						<td>
							<select name="tbxFraManed" id="tbxFraManed">
								<%
								for teller = 0 to 12
									if cint(fraMaaned) = teller then
										selected = "selected"
									else
										selected = ""
									end if
									%>
									<option <%=selected%> value="<%=teller%>"><%=AMonths(teller)%></option>
									<%
								next
								%>
							</select>
						</td>
						<td>Til måned:</td>
						<td>
							<select name="tbxTilManed" id="tbxTilManed">
								<%
								for teller = 0 to 12
									if cint(tilMaaned) = teller then
										selected = "selected"
									else
										selected = ""
									end if
									%>
									<option <%=selected%> value="<%=teller%>"><%=AMonths(teller)%></option>
									<%
								next
								%>
							</select>
						</td>
						<td><label for="chkKunEgne">Kun egne</label>&nbsp;<input type="checkbox" <%=EgneAvkrysset%> class="checkbox" id="chkKunEgne" name="chkKunEgne" value="1"></td>
						<td><input type="submit" name="pbnDataAction" value="Søk"></td>
					</tr>
				</table>
			</form>
			<%
			' Create table only when records found
			If (foundRecords = true)  Then
				' Create table
				Response.Write "<div class='listing'><table cellpadding='0' cellspacing='1'>"

				' Create table heading
				Response.Write "<tr>"
				Response.Write "<th>Navn</th>"
				Response.Write "<th>Fødselsdato</th>"
				Response.Write "</tr>"

				' Create new fromdate
				FromMonth = Month( Fradato )
				FromDay = Day( Fradato )

				Do Until rsRapport.EOF
						' Create row
						Response.Write "<tr>"
						
						Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsRapport("VikarID"), rsRapport( "Fornavn") & " " & rsRapport( "Etternavn"), "Vis vikar " & rsRapport( "Fornavn") & " " & rsRapport( "Etternavn") ) & "</td>"
						Response.Write "<td class='right'>" & rsRapport("Foedselsdato") & "</td>"
						Response.Write "</tr>"

					' Get next record
					rsRapport.MoveNext
				Loop

				' Close recordset
				rsRapport.Close

				' Clear recordset
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