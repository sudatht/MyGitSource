<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
	dim rsAvd
	dim viskode
	dim mm
	dim yy
	dim dato1
	dim strOppdragID
	dim avd
	dim valgt_avd
	dim teller
	dim k
	dim kk
	dim BES
	dim bok
	dim rsFakt
	dim mark
	dim strKontaktperson
	Dim cts 
	dim personRs 		

	' Connect to database
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	'Parameters
	Session("tilgang") = 3

	If Request("viskode") = "" Then
		viskode = session("viskode")
	Else
		viskode = Request("viskode")
		session("viskode") = viskode
	End IF

	mm = Datepart("m", Date): mm = mm + 1
	yy = Datepart("yyyy", DateValue( Date))
	If mm = 13 Then yy = yy + 1: mm = 1
	If mm < 10 Then dd = "01.0" Else dd="01."
	yy = Right(CStr(yy), 2)
	dato1 = dd & mm & "." & yy

	If Request("Dato1") <> "" Then
		dato1 = Request("Dato1")
	End If

	session("limitDato") = dato1
	session("oldStartDate") = dato2

	If Request("OppdragID") = "" Then
		strOppdragID = Request.Form("OppdragID")
	Else
		strOppdragID = Request("OppdragID")
	End If

	if len(Request("avd")) = 0 then
		avd = "0"
	else
		avd = Request("avd")
	end if
	

	If Request("velgAvdeling") <> "" Then
		valgt_avd = CInt(Request("velgAvdeling"))
	else
		valgt_avd = 0
	End If

	' BUTTONS
	If viskode =  1 Then k = "-> " Else k = "     "
	If viskode =  2 Then kk = "-> " Else kk = "     "
%>
<html>
	<head>
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
		<title>Fakturagrunnlag</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<a name="top"></a>
			<div class="contentHead1">
				<h1>Fakturabehandling</h1>
				<h2>Nye fakturer</h2>
			</div>
			<div class="content">
				<br>
				<p>
					<form name="EN" action="Faktura_list.asp?viskode=1" method="post" ID="Form1">
						<SELECT id="avd" name="avd">
							<option value="">Alle</option>
							<%
							dim rsAvdelinger
							set rsAvdelinger = HentAlleAktiveAvdelinger()
							If(HasRows(rsAvdelinger)) then
								while not rsAvdelinger.EOF
									if cint(avd) = clng(rsAvdelinger("AvdelingID")) then
										selected = "selected"
									else
										selected = ""
									end if
									%>
									<option <%=selected%> value="<%=rsAvdelinger("AvdelingID")%>"><% = rsAvdelinger("Avdeling") %></option>
									<%
									rsAvdelinger.movenext
								wend
								rsAvdelinger.close
							end if
							set rsAvdelinger = nothing
							%>							
						</select>
						&nbsp;<input type="text" size="7" maxlength="8" name="dato1" id="dato1" VALUE="<% =dato1 %>" ONBLUR="dateCheck(this.form, this.name)" >
						&nbsp;<input type="SUBMIT" value="<% =k %>Nye fakturaer" ID="Submit1" NAME="Submit1">
					</form>
				</p>
			</div>
			<div class="contentHead"><h2>Klargjorte fakturaer</h2></div>
			<div class="content">
				</p>
					<form name="syv" action="Faktura_list.asp?viskode=2&dato1=<% =dato1 %>" method="post" ID="Form2">
						<%
						set rsAvd = HentAlleAktiveAvdelinger()
						If(HasRows(rsAvd)) then
							%>
							<SELECT name="velgAvdeling" ID="Select1">
								<option value="0" <%if valgt_avd = 0 then%> selected <% end if %>>
									Alle
								</option>
								<%
								while not rsAvd.EOF
									%>
									<option value="<%=rsAvd("AvdelingID").value%>" <%if valgt_avd = rsAvd("AvdelingID").value then%> selected <%end if%>>
										<%=rsAvd("Avdeling").value%>
									</option>
									<%
									rsAvd.MoveNext
								wend
								%>
							</select>
							<%
							rsAvd.close
						end if
						set rsAvd = nothing
						%>
						&nbsp;<input TYPE="submit" VALUE="<% =kk %>     Klargjorte -  overføring        " ID="Submit2" NAME="Submit2">
					</form>
				</p>
			</div>
			<div class="contentHead"><h2>Grunnlag</h2></div>
			<div class="content">
				<%
				' Display table headdings
				If viskode = 1 Then 
					%>
					<p>
						<A HREF=#A>A</A>
						<A HREF=#B>B</A>
						<A HREF=#C>C</A>
						<A HREF=#D>D</A>
						<A HREF=#E>E</A>
						<A HREF=#F>F</A>
						<A HREF=#G>G</A>
						<A HREF=#H>H</A>
						<A HREF=#I>I</A>
						<A HREF=#J>J</A>
						<A HREF=#K>K</A>
						<A HREF=#L>L</A>
						<A HREF=#M>M</A>
						<A HREF=#N>N</A>
						<A HREF=#O>O</A>
						<A HREF=#P>P</A>
						<A HREF=#R>R</A>
						<A HREF=#S>S</A>
						<A HREF=#T>T</A>
						<A HREF=#U>U</A>
						<A HREF=#V>V</A>
						<A HREF=#Ø>Ø</A>
						<A HREF=#Å>Å</A>
					</p>
					<div class="listing">
					<% 
				End If
				strSQL = ""
				If viskode = 1 Then
					If avd > 0 Then
						strSQL = "SELECT DISTINCT D.OppdragID, D.VikarID, D.Bestilltav, " &_
						" F.FirmaID, F.Firma, F.SOCuID, kontakt = (K.Etternavn + ' ' + K.Fornavn), D.SOBestilltAv," &_
						" Vnavn = (V.Etternavn + ' ' + V.Fornavn)" &_
						" FROM  DAGSLISTE_VIKAR D, OPPDRAG O, FIRMA F, KONTAKT K, VIKAR V" &_
						" WHERE  D.OppdragID = O.OppdragID " &_
						" AND D.FirmaID = F.FirmaID" &_
						" AND D.Bestilltav *= K.KontaktID " &_
						" AND D.VikarID = V.VikarID" &_
						" AND O.AvdelingID = " & avd &_
						" AND D.FakturaStatus < 3" &_
						" AND D.Dato < " & dbDate(session("limitDato")) &_
						" ORDER BY F.Firma, kontakt, Vnavn"
					Else
						strSQL = "SELECT DISTINCT D.OppdragID, D.VikarID, D.Bestilltav, " &_
						" F.FirmaID, F.SOCuID, F.Firma, kontakt = (K.Etternavn + ' ' + K.Fornavn), D.SOBestilltAv, " &_
						" Vnavn = (V.Etternavn + ' ' + V.Fornavn)" &_
						" FROM  DAGSLISTE_VIKAR D, OPPDRAG O, FIRMA F, KONTAKT K, VIKAR V" &_
						" WHERE  D.OppdragID = O.OppdragID " &_
						" AND D.FirmaID = F.FirmaID" &_
						" AND D.Bestilltav *= K.KontaktID " &_
						" AND D.VikarID = V.VikarID" &_
						" AND D.FakturaStatus < 3" &_
						" AND D.Dato < " & dbDate(session("limitDato")) &_
						" ORDER BY F.Firma, kontakt, Vnavn"
					End If
					Set rsFakt = GetFirehoseRS(strSQL, Conn)
				ElseIf viskode = 2 Then

					if valgt_avd > 0 then
						strSQL = "SELECT DISTINCT D.FirmaID, Firma.SOCuID, Firma, Navn = (Etternavn + ' ' + Fornavn), " &_
							" D.Kontakt, D.SOKontakt, stat1 = D.Status, FakturaDato, FakturaNr, " &_
							" AvdelingID = Avdeling " &_
							" FROM  FAKTURAGRUNNLAG D, KONTAKT, FIRMA " &_
							" WHERE  D.Status = 2 " &_
							" AND D.FirmaID = FIRMA.FirmaID " &_
							" AND D.Kontakt *= KONTAKT.KontaktID" &_
							" AND Avdeling = " & valgt_avd &_
							" ORDER BY FakturaNr"
						
						Set rsFakt = GetFirehoseRS(strSQL, Conn)								
					else
						strSQL = "SELECT DISTINCT D.FirmaID, Firma, Firma.SOCuID, Navn=(Etternavn + ' ' + Fornavn), " &_
							" D.Kontakt, D.SOKontakt, stat1 = D.Status, FakturaDato, FakturaNr," &_
							" AvdelingID = Avdeling " &_
							" FROM  FAKTURAGRUNNLAG D, KONTAKT, FIRMA " &_
							" WHERE  D.Status = 2 " &_
							" AND D.FirmaID = FIRMA.FirmaID" &_
							" AND D.Kontakt *= KONTAKT.KontaktID" &_
							" ORDER BY FakturaNr"
						
						Set rsFakt = GetFirehoseRS(strSQL, Conn)								
					end if
				End If  'viskode
				
				'	Nye fakturarer
				If viskode = 1 and hasRows(rsFakt) Then
					%>
					<table ID="Table1">
						<tr>
							<th>kontaktid</th>
							<th>Kontakt</th>
							<th>Bestilt av</th>
							<th>Oppdrag</th>
							<th colspan="3">Vikar</th>
						</tr>
						<%
						'Sjekk om det er noen nye fakturaer
						BES = ""
						Bok = Left(rsFakt("Firma"), 1)
						
						Set cts = server.CreateObject("Integration.SuperOffice")						
						
						do while not rsFakt.EOF
							If Bok <> Left(rsFakt("Firma"), 1) Then
								Bok = Left(rsFakt("Firma"), 1)
								mark = "<a name=" & Bok & "></a>"
							else
								mark = ""
							End If 
							%>
							<tr>
								<%
								strKontaktperson = rsFakt("kontakt")
								if(isnull(strKontaktperson) and not isnull(rsFakt("SOBestilltAv"))) then
									set personRs = cts.GetPersonSnapshotById(clng(rsFakt("SOBestilltAv")))
									if (not personRs.EOF) then
										if (isnull(personRs("middlename"))) then
											strKontaktperson = personRs("firstname") & " " & personRs("lastname")
										else
											strKontaktperson = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
										end if
									end if
									set personRs = nothing
								end if								
								
								If BES <> strKontaktperson Then
									BES = strKontaktperson
									%>				
										<TD width="10%"><%=mark%><% =rsFakt("FirmaID") %></TD>
										<TD width="20%"><%=CreateSONavigationLink(SUPEROFFICE_PANEL_CONTACT_URL, SUPEROFFICE_PANEL_CONTACT_URL, rsFakt("SOCuID").Value, rsFakt("Firma").Value, "Vis kontakt '" & rsFakt("Firma").Value & "'") %></TD>
										<TD width="20%"><%=strKontaktperson %></TD>
										<TD colspan="3"></TD>
									</tr>
									<tr>
										<TD width="50%" colspan="3"></td>
										<% 
								Else 
									%>
									<TD width="50%" colspan="3"><%=mark%></td>
									<% 
								End If 
								%>
								<td><% =CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "oppdragVis.asp?oppdragID=" & rsFakt("oppdragID").Value, rsFakt( "oppdragID"), "Vis Oppdrag")%></TD>
								<td><% =rsFakt("VikarID") %></td>
								<td><% =CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsFakt( "VikarID" ), rsFakt( "Vnavn"), "Vis vikar " & rsFakt( "Vnavn") ) %></td>
								<td><a HREF="Vikar_timeliste_vis3.asp?VikarID=<% =rsFakt("VikarID") %>&OppdragID=<% =rsFakt("OppdragID") %>&dato1=<%=session("limitDato") %>&frakode=3" TARGET=_new >Vis timeliste</a></TD>
							</tr>
							<%
							rsFakt.MoveNext
						loop
						set cts = nothing	
						rsFakt.close
						Set rsFakt = Nothing
						%>
					</table>
					<%
				ElseIf viskode = 2 and hasRows(rsFakt) Then
					%>
					<table ID="Table2">
						<tr>
							<th>Fakt.nr</th>
							<th>Kontakt</th>
							<th>Kontaktperson</th>
							<th>Avdeling</th>
							<th>Status</th>
							<th>Nedgrader</th>
						</tr>
						<% 
						Set cts = server.CreateObject("Integration.SuperOffice")						

						do while not rsFakt.EOF 
							strKontaktperson = rsFakt("Navn")
							if(isnull(strKontaktperson)) then
								set personRs = cts.GetPersonSnapshotById(clng(rsFakt("SOKontakt")))
								if (not personRs.EOF) then
									if (isnull(personRs("middlename"))) then
										strKontaktperson = personRs("firstname") & " " & personRs("lastname")
									else
										strKontaktperson = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
									end if
								end if
								set personRs = nothing	
							end if						
							%>
							<tr>
								<td><% =rsFakt("Fakturanr")%></td>
								<td><a href="Faktura_vis.asp?Kontakt=<% =rsFakt("Kontakt")%>&SOKontakt=<% =rsFakt("SOKontakt")%>&FirmaID=<% =rsFakt("FirmaID") %>&Avdeling=<% =rsFakt("AvdelingID") %>"><% =rsFakt("Firma")%></a></td>
								<td><%=strKontaktperson%>
								<% avd = HentAvdNavn(rsFakt("AvdelingID"))%></td>
								<td><% = avd %></td>
								<td>(<% =rsFakt("stat1") %>)</td>
								<td><A href="Faktura_lagre.asp?graderingskode=nedgrad&Kontakt=<%=rsFakt("Kontakt")%>&SOKontakt=<%=rsFakt("SOKontakt")%>&FirmaID=<%=rsFakt("FirmaID")%>&Fakturadato=<%=rsFakt("Fakturadato")%>&Fakturanr=<%=rsFakt("Fakturanr")%>&velgAvdeling=<%=valgt_avd%>&viskode=<%=viskode%>&dato1=<% =dato1 %>">Nedgrad</A></TD>
							</tr>
							<% 
							rsFakt.MoveNext
						loop
						set cts = nothing
						rsFakt.Close
						Set rsFakt = Nothing 
						%>
						<tr>
							<td colspan="6">
								<table ID="Table3">
									<tr>
										<form name="aatte" action="Faktura_overf_ordrehode.asp?avd=<%=valgt_avd%>" method="post" ID="Form3">
										<th><input name=btnOverfoer TYPE=SUBMIT  VALUE="           Lag fil til Visma              " ID="Submit3"></th>
										<%
										'<th>Siste ordrenr:
										'<th><input name=ONR TYPE=TEXT SIZE=5 ID="Text1"></th>
										%>
										</form>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					<% 
				elseIf hasRows(rsFakt) = false Then
					set rsFakt = nothing
					'Feilmelding hvis det ikke finnes poster
					call Response.Write("<p>Ingen fakturaer funnet!</p>")					
				End If 'viskode
				%>
				<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="til toppen"> Til toppen</a><br>&nbsp;
			</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>