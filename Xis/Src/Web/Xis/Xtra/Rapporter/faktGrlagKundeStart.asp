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
	dim heading
	dim rsFirma
	dim vikar : vikar = false

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	'Henter parametere for valgt måned og vikar

	periode =  request("periode")
	avdeling = request("dbxAvdeling")
	avdelinger = request("dbxAvdeling")

	if avdeling = "0" Then
		avdeling = ""
		strSQL = "SELECT AvdelingID FROM Avdeling"
		Set rsAvdeling = GetFireHoseRS(strSQL, Conn)
    	Do Until rsAvdeling.EOF
    		avdeling = avdeling & rsAvdeling("AvdelingID") & ","
			rsAvdeling.MoveNext
		loop
		rsAvdeling.close
		set rsAvdeling = Nothing
		strLengde = len(avdeling)
		avdeling = left(avdeling,(strLengde-1))
	end if

	mnd =  mid(periode, 5)
	aar = left(periode, 4)

	if NOT mnd="" Then
		periode = aar & mnd
		
		if mnd < 10 Then periode = aar & "0" & mnd
		
		vikar = true

		SELECT	CASE mnd
			CASE 1
				mnd2 = "Januar"
			CASE 2
				mnd2 = "Februar"
			CASE 3
				mnd2 = "Mars"
			CASE 4
				mnd2 = "April"
			CASE 5
				mnd2 = "Mai"
			CASE 6
				mnd2 = "Juni"
			CASE 7
				mnd2 = "Juli"
			CASE 8
				mnd2 = "August"
			CASE 9
				mnd2 = "September"
			CASE 10
				mnd2 = "Oktober"
			CASE 11
				mnd2 = "November"
			CASE 12
				mnd2 = "Desember"
		END SELECT
		'Henter opplysninger om vikar og oppdrag i perioden

		strSQL ="SELECT DISTINCT v.firmaID, f.SOCuiD, f.firma "&_
			" FROM firma f, vikar_ukeliste v, oppdrag o "&_
			" WHERE v.faktperiode = " & periode &_
			" AND v.firmaid = f.firmaid "&_
			" AND v.oppdragID = o.oppdragID "&_
			" AND o.avdelingID IN("& avdeling &") "&_
			" ORDER BY f.firma "
		Set rsFirma = GetFireHoseRS(strSQL, Conn)
		
		heading = "Fakturagrunnlag for " & mnd2 & "  " & aar 
	else
		heading = "Fakturagrunnlag"
	End if
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
		<title><%=heading%></title>
		<script language="javaScript" type="text/javascript">
			function sjekk()
			{
				var countBx = document.forms[1].elements.length;

				for (var i=0; i < countBx; i++)
				{
					if (document.forms[1].elements[i].type == "checkbox")
					{
					if (document.forms[1].elements[i].checked == true)
						document.forms[1].elements[i].checked = false;
					else
						document.forms[1].elements[i].checked = true;
					}
				}
			}
		</script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=heading%></h1>
			</div>
			<div class="content">
			<form action="faktgrlagkundeStart.asp" method="post">
				<%
				YYYY = year(date)
				MM = month(date) - 9

				if MM < 1 THEN
					MM = MM + 12
					YYYY = YYYY-1
				End If

				i = 1
				%>
					<SELECT NAME="periode" ID="Select1">
						<% 
						do until i = 13
							SELECT CASE MM

								CASE 1
									mnd = "januar"
								CASE 2
									mnd = "februar"
								CASE 3
									mnd = "mars"
								CASE 4
									mnd = "april"
								CASE 5
									mnd = "mai"
								CASE 6
									mnd = "juni"
								CASE 7
									mnd = "juli"
								CASE 8
									mnd = "august"
								CASE 9
									mnd = "september"
								CASE 10
									mnd = "oktober"
								CASE 11
									mnd = "november"
								CASE 12
									mnd = "desember"
							END SELECT

							if (trim(YYYY&MM)= trim(request("periode"))) THEN
								sel = " SELECTED"
							else
								sel = ""
							end if
							%>
							<option value="<%=YYYY & MM%>" <% =sel %>> <%=mnd & " " & YYYY %></option>
							<%
							MM = MM + 1
							If MM > 12 Then
								MM = 1
								YYYY = YYYY + 1
							End If
							i = i+1
						loop
						%>
					</select>
					Avdeling:
					<select NAME="dbxAvdeling" ID="Select2">
						<option value="0">Alle</option>
						<%
						' Get avdeling
						strSQL = "SELECT AvdelingID, Avdeling FROM Avdeling WHERE Show_Hide = 0 ORDER BY avdeling"
						Set rsAvdeling = GetFireHoseRS(strSQL, Conn)
						Do Until rsAvdeling.EOF
							if rsAvdeling("AvdelingID")=cint(request("dbxAvdeling")) THEN
      							strSelected =  " SELECTED"
							else
      							strSelected = ""
							end if
							%>
							<option value="<% =rsAvdeling("AvdelingID") %>" <%=strSelected%>><% =rsAvdeling("Avdeling") %></option>
							<%   
							rsAvdeling.MoveNext
						Loop
						' Close and release recordset
						rsAvdeling.Close
						Set rsAvdeling = Nothing
						%>
					</select>
					<INPUT TYPE=submit value=" Velg periode og avdeling  " ID="Submit1" NAME="Submit1">
				</form>
				<form action="faktgrlagKundeKontakt.asp" target="_new" ID="Form2">
					<input type="hidden" name="periode" value="<%=periode%>" ID="Hidden1">
					<input type="hidden" name="avdelinger" value="<%=avdeling%>" ID="Hidden2">
					<% 
					if vikar = true Then 
						if (HasRows(rsFirma))  Then 
							%>
							<div class="listing">
								<input type=submit value="Hent utvalg" ID="Submit2" NAME="Submit2">
								<table cellpadding='0' cellspacing='0' ID="Table1">
									<tr>
										<th>FirmaID</th>
										<th>Firma</th>
										<th>&nbsp;</th>
									</tr>
									<%
									do until rsFirma.EOF
										%>
										<tr>
											<td class="right">
												<%=rsFirma("firmaid")%>
											</td>
											<td>
												<%=rsFirma("firma")%>
											</td>
											<td>
												<INPUT TYPE="radio" class="radio" NAME="valg" VALUE="<%=rsFirma("firmaid")%>"
										ID="Radio1"	</td>
										</tr>
										<%
										rsFirma.movenext
									loop
									%>
								</table>
								<input type="submit" value="Hent utvalg" ID="Submit3" NAME="Submit3">
							</div>
							<%
						end if
						rsFirma.close
						set rsFirma = nothing
					end if
					%>
				</form>					
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>