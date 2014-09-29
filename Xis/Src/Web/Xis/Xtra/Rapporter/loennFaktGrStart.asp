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
		avdeling = left(avdeling,(strLengde - 1))
	end if

	mnd =  mid(periode, 5)
	aar = left(periode, 4)

	if NOT mnd="" Then
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

		if avdeling <> "" then
			strSQL = "SELECT DISTINCT d.vikarid, Navn=(v.Etternavn+ ', ' + v.Fornavn), VIKAR_ANSATTNUMMER.ansattnummer "&_
				" FROM vikar v LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON v.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, dagsliste_vikar D, oppdrag O "&_
				" WHERE datepart(month, D.dato)="& mnd &_
				" AND datepart(year, D.dato)= "& aar &_
				" AND d.vikarid=v.vikarid "&_
				" AND O.oppdragID = d.oppdragid "&_
				" AND O.avdelingID IN(" & avdeling & ") Order by 2"
				
			Set rsVikar = GetFireHoseRS(strSQL, Conn)				
		end if
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
	</head>
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
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Lønn-/ fakturagrunnlag <%if mnd2 <> "" then %> for <% = mnd2 & " " & aar%> <%end if%></h1>
			</div>
			<div class="content">
				<form action="loennFaktGrStart.asp" method="post">
					<%
					YYYY = year(date)
					MM   = month(date) - 9


					if MM < 1 THEN
						MM = MM + 12
						YYYY = YYYY-1
					End If

					i=1
					%>
					<select NAME="periode">
						<% 
						do until i = 13
						
							SELECT	CASE MM
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
							Else
								sel = ""
							end if
							%>
							<option value="<%=YYYY&MM%>" <%=sel%>> <%=mnd&" "&YYYY%></option>
							<%
							MM = MM+1
							If MM > 12 Then
								MM = 1
								YYYY = YYYY + 1
							End If
							i = i+1
						loop
						%>
					</select>
					Avdeling:
					<SELECT NAME="dbxAvdeling">
						<option VALUE="0">Alle</option>
							<%
							' Get avdeling
							strSQL = "Select AvdelingID, Avdeling from Avdeling WHERE Show_Hide = 0 order by avdeling"
							Set rsAvdeling = GetFireHoseRS(strSQL, Conn)

							Do Until rsAvdeling.EOF
								if rsAvdeling("AvdelingID")=cint(request("dbxAvdeling")) THEN
	      							strSelected =  " SELECTED"
								else
	      							strSelected = ""
								end if
								%>
								<option VALUE="<% =rsAvdeling("AvdelingID") %>" <%=strSelected%>><% =rsAvdeling("Avdeling") %></option>
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
				<%
				if avdeling <> "" then
					%>
					<form action="loennFaktGrunnlag.asp" ID="Form2">
						<input type="hidden" name="periode" value="<%=periode%>" ID="Hidden1">
						<input type="hidden" name="avdelinger" value="<%=avdeling%>" ID="Hidden2">
						<% 
						if vikar = true Then 
							if HasRows(rsVikar) = true  Then 
								%>
								<A HREF=#A>A</a>
								<A HREF=#B>B</a>
								<A HREF=#C>C</a>
								<A HREF=#D>D</a>
								<A HREF=#E>E</a>
								<A HREF=#F>F</a>
								<A HREF=#G>G</a>
								<A HREF=#H>H</a>
								<A HREF=#I>I</a>
								<A HREF=#J>J</a>
								<A HREF=#K>K</a>
								<A HREF=#L>L</a>
								<A HREF=#M>M</a>
								<A HREF=#N>N</a>
								<A HREF=#O>O</a>
								<A HREF=#P>P</a>
								<A HREF=#R>R</a>
								<A HREF=#S>S</a>
								<A HREF=#T>T</a>
								<A HREF=#U>U</a>
								<A HREF=#V>V</a>
								<A HREF=#Ø>Ø</a>
								<A HREF=#Å>Å</a>
				
								<input type=submit value="Hent rapport"><br>
								<input type=button onClick="sjekk();" value="Merk alle">
								<div class="listing">
									<table>
										<tr>
											<th>Ansattnummer</th>
											<th>Navn</th>
											<th>&nbsp;</th>
											<%	
											dim strBokmerke
											dim strBokstav
											While not rsvikar.EOF
												If strBokstav <> Left(rsVikar("navn"),1 ) Then
													strBokstav = Left(rsVikar("navn"),1 )
													strBokmerke = strBokstav
												else
													strBokmerke = ""
												End If
												%>
												</tr>
												<tr ID="<%=strBokmerke%>">
													<td class="right"><%=rsvikar("ansattnummer").Value%></td>
													<td><%=rsvikar("navn")%></td>
													<td><INPUT class="checkbox" TYPE="CHECKBOX" NAME="valg" VALUE="<%=rsVikar("vikarid")%>"</td>
												</tr>
												<%			
												rsvikar.movenext
											Wend
											%>
										<tr>
											<th colspan="3"><input type=submit value="Hent rapport"></th>
										</tr>
									</table>
								</div>
								<%
							end if
							rsVikar.close
							set rsVikar = nothing
							CloseConnection(Conn)
							set Conn = nothing
						end if
						%>
					</form>
					<%
				end if
				%>
			</div>
		</div>
	</body>
</html>

