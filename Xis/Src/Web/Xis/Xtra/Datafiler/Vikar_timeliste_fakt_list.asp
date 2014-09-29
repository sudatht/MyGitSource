<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Buffer = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Settings.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim rsVikar
	dim visKode
	dim personRs
	dim kontaktID
	dim strKontaktperson	
	dim Conn
	Dim strSQL
%>
<html>
	<head>
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
		<title>Faktura</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Faktura</h1>
			</div>		
			<div class="content">
				<%
				' PARAMETERS
				Session("tilgang") = 3

				If Request.QueryString("viskode") = "" Then
					visKode = session("viskode")
				Else
					visKode = Request.QueryString("viskode")
					session("viskode") = viskode
				End if

				mm = Datepart("m", Date): mm = mm + 1
				yy = Datepart("yyyy", DateValue( Date))
				If mm = 13 Then yy = yy + 1: mm= 1
				If mm < 10 Then dd = "01.0" Else dd="01."
				yy = Right(CStr(yy),2)
				dato1 = dd & mm & "." & yy
				dato2 = (aDate - 60)

				If Request.QueryString("Dato1") <> "" Then
					dato1 = Request.QueryString("Dato1")
				End If
				If Request.Form("Dato1") <> "" Then
					dato1 = Request.Form("Dato1")
				End If
				session("limitDato") = dato1

				If Request.QueryString("Dato2") <> "" Then
 					dato2 = Request.QueryString("Dato2")
				End If
				If Request.Form("Dato2") <> "" Then
 					dato2 = Request.Form("Dato2")
				End If
				session("oldStartDate") = dato2

				If Request.QueryString("OppdragID") = "" Then
					strOppdragID = Request.Form("OppdragID")
				Else
					strOppdragID = Request.QueryString("OppdragID")
				End If

				' BUTTONS
				If viskode =  1 Then k = "-> " Else k = "     "
				If viskode =  2 Then kk = "-> " Else kk = "     "
				If viskode =  3 Then kkk = "-> " Else kkk = "     "
				%>
				<FORM name="en" ACTION="Vikar_timeliste_fakt_list.asp?viskode=1&dato2=<% =dato2 %>" METHOD=POST ID="Form1">
					<INPUT TYPE=SUBMIT VALUE="<% =k %>Nye fakturaer" ID="Submit1" NAME="Submit1"><br>
					<input type="text" size="8" maxlength="8" NAME="DATO1" VALUE="<% =dato1 %>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1">
				</form>

				<FORM name="to" ACTION="Vikar_timeliste_fakt_list.asp?viskode=3&dato1=<% =dato1 %>" METHOD=POST ID="Form2">
					<INPUT TYPE=SUBMIT VALUE="<% =kkk %>Gamle" ID="Submit2" NAME="Submit2"><br>
					<input type="text" size="8" maxlength="8" NAME="DATO2" VALUE="<% =dato2 %>" ONBLUR="dateCheck(this.form, this.name)" ID="Text2">
				</form>

				<% 
				If viskode < 2 Or viskode = 3 Then 
					%>
					<FORM name="tre" ACTION="Vikar_timeliste_ny3.asp?frakode=2" METHOD=POST ID="Form3">
						<p>Vikarnr: <input type="text" name=VIKARID SIZE=4 ID="Text3"></p>
						<p>Oppdragnr: <input type="text" name=OPPDRAGID SIZE=4 ID="Text4"></p>
						<INPUT TYPE=SUBMIT VALUE="Søk" ID="Submit3" NAME="Submit3">
					</form>
					<% 
				End If 
				%>
				<FORM name="tre" ACTION="Vikar_timeliste_fakt_list.asp?viskode=2&dato1=<% =dato1 %>&dato2=<% =dato2 %>" METHOD=POST ID="Form4">
					<INPUT TYPE=SUBMIT VALUE="<% =kk %> Klargjorte - overføring" ID="Submit4" NAME="Submit4">
				</form> 
				<%

				Set Conn = GetConnection(GetConnectionstring(XIS, ""))
			
				' SQL
				If viskode = 1 Then

				strSQL = "SELECT DISTINCT [D].[FirmaID], [F].[Firma], [Navn]=([K].[Etternavn] + ' ' + [K].[Fornavn]), " &_
						" [D].[BestilltAv], [D].[SoBestilltAv], [stat1] = [D].[Fakturastatus], [V].[VikarID], [VNavn] = ([V].[Etternavn] + ' ' + [V].[Fornavn]), " &_
						" [AvdelingID] " &_
						" FROM [DAGSLISTE_VIKAR] AS [D], [KONTAKT] AS [K], [FIRMA] AS [F], [OPPDRAG] AS [O], [VIKAR] AS [V]" &_
						" WHERE [D].[FakturaStatus] < 3" &_
						" AND [D].[FirmaID] = [F].[FirmaID]" &_
						" AND [D].[BestilltAv] = [K].[KontaktID]" &_
						" AND [D].[OppdragID] = [O].[OppdragID]" &_
						" AND [D].[VikarID] = [V].[VikarID]" &_
						" AND [D].[Dato] < " & dbDate(session("limitDato")) &_
						" ORDER BY [Firma]"

					Set rsVikar = GetFirehoseRS(strSQL, conn)

				ElseIF viskode = 2 Then

					strSQL = "SELECT DISTINCT [D].[FirmaID], [Firma], [Navn] = ([Etternavn] + ' ' + [Fornavn]), " &_
						"[BestilltAV]=[D].[Kontakt], [SoBestilltAv] = [D].[SOKontakt], [stat1] = [D].[Status], [FakturaDato], [FakturaNr], " &_
						"[AvdelingID] = [Avdeling] " &_
						"FROM [FAKTURAGRUNNLAG] AS [D], [KONTAKT], [FIRMA] " &_
						"WHERE [D].[Status] = 2 " &_
						"AND [D].[FirmaID] = [FIRMA].[FirmaID] " &_
						"AND [D].[Kontakt] = [KONTAKT].[KontaktID] " &_
						" ORDER BY [FakturaNr]"

					Set rsVikar = GetFirehoseRS(strSQL, conn)

				ElseIf viskode = 3 Then

					strSQL = "SELECT DISTINCT [D].[FirmaID], [F].[Firma], [Navn] = ([K].[Etternavn] + ' ' + [K].[Fornavn]), " &_  
						"BestilltAV=[D].[Kontakt], [SoBestilltAv] =[D].[SOKontakt], [AvdelingID]=[D].[Avdeling] " &_  
						"FROM [FAKTURAGRUNNLAG] AS [D], [KONTAKT] AS [K], [FIRMA] AS [F] " &_
						"WHERE [D].[Status] = 3 " &_
						"AND [D].[FirmaID] = [F].[FirmaID] " &_
						"AND [D].[Kontakt] = [K].[KontaktID] " &_
						"AND [D].[FakturaDato] >= "  & dbDate(session("oldStartDate")) &_
						" ORDER BY [F].[Firma], [D].[Kontakt]"

					Set rsVikar = GetFirehoseRS(strSQL, conn)
				End If  'viskode

				If viskode > 0 Then

					If HasRows(rsVikar) = false Then 
						Response.write "<p>Ingen fakturaer med status " & viskode & "!</p>"
					Else
						' Display table headdings
						If viskode = 1 Or viskode = 3 Then 
							%>
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
							<a HREF=#Ø>Ø</A>
							<a HREF=#Å>Å</A>
							<% 
						End If 
					%>
					<div class="listing">
						<table cellpadding='0' cellspacing='0' border="1" ID="Table1">
							<tr>
							<% 
							If viskode = 2 Then 
								%>
								<th>Fakt.nr</th>
								<% 
							End If 
							If viskode = 3 Then 
								%>
								<th>kontaktid</th>
								<% 
							End If 
							%>
							<th>Kontakt</th>
							<th>Kontaktperson</th>
							<th>Avd</th>
							<% 
							If viskode < 3 Then 
								%>
								<th>FStat</th>
								<% 
								If viskode = 1 Then 
									%>
									<th>Vikar</th>
									<th>TStat</th>
									<th>LStat</th>
									<% 
								End If 
								%>
								<th>Fakt</th>
								<% 
								If viskode = 2 Then 
									%>
									<th>Nedgrad</th>
									<% 
								End If 
							End If 
								' Show data
								Bok = Left(rsVikar("Firma"), 1)
								Do Until rsVikar.EOF
								%>
									<tr>
										<% 
										If viskode = 2 Then 
											%>
											<td><% =rsVikar("Fakturanr") %></td>
											<% 
										End If 
										If viskode = 3 Then 
											%>
											<td><% =rsVikar("FirmaID") %></td>
											<% 
										End If 
										If viskode = 1 Then 
											%>
											<td> <% =rsVikar("Firma") %></td>
											<% 
										End If 
										If viskode = 2 Then 
											%>
											<td><A HREF="Faktura_vis.asp?Kontakt=<% =rsVikar("BestilltAv") %>&SOKontakt=<% =rsVikar("SoBestilltAv") %>&OppdragID=<% '=rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>&Avdeling=<% =rsVikar("AvdelingID") %>"   ><% =rsVikar("Firma") %></A></td>
											<% 
										End If 
										If Viskode = 3 Then 
											If Bok <> Left(rsVikar("Firma"),1) Then
												Bok = Left(rsVikar("Firma"),1)	
												Response.write "<A NAME=" & Bok & ">"
											End If 
											%>
											<td><A HREF="Faktura_vis_gml.asp?Kontakt=<% =rsVikar("BestilltAv") %>&SOKontakt=<% =rsVikar("SoBestilltAv") %>&FirmaID=<% =rsVikar("FirmaID") %>&AvdelingID=<% =rsVikar("AvdelingID") %>" TARGET="_new"  ><% =rsVikar("Firma") %></A></td>
											<% 
										End If 

										if(IsNull(rsVikar("BestilltAV"))) then
											kontaktID = rsVikar("SOBestilltAv")
											set personRs = cts.GetPersonSnapshotById(clng(kontaktID))
											if (not personRs.EOF) then
												if (isnull(personRs("middlename"))) then
													strKontaktperson = personRs("firstname") & " " & personRs("lastname")
												else
													strKontaktperson = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
												end if
											end if
											set personRs = nothing
										else
											kontaktID = rsVikar("BestilltAv")
											strKontaktperson = rsVikar("navn")
										end if
										%>
										<td><% =kontaktID %>-<% =strKontaktperson %></td>
										<td><% =rsVikar("AvdelingID") %></td>
										<% 
										If Viskode < 3 Then 
											%>
											<td>
											<% 
											If rsVikar("stat1") = 3 Then 
												%>
												<font COLOR="BLACK" >
												<% 
											ElseIf rsVikar("stat1") = 2 Then 
												%>
												<font COLOR="GREEN" >
												<% 
											Else 
												%>
												<font COLOR="#CB1700">
												<% 
											End If 
											%>
											<strong>(<% =rsVikar("stat1") %>)</strong>
											</font>
											</td>
											<% 
										End If 'viskode 
										If viskode = 1 Then 
											%>
											<td></td>
											<td></td>
											<td></td>
											<% 
										End If 
										If viskode < 3 Then 
											%>
											<td></td>
											<% 
										End If 
										If viskode = 1 Then 
											%>
											<td><A HREF="Faktura_vis.asp?Kontakt=<% =rsVikar("BestilltAv") %>&SOKontakt=<% =rsVikar("SoBestilltAv") %>&FirmaID=<% =rsVikar("FirmaID") %>&Avdeling=<% =rsVikar("AvdelingID") %>" TARGET=_new  >Fakt</A></td>
											<% 
										End If 
										If viskode = 2 Then 
											%>
											<td><A HREF="Faktura_lagre.asp?graderingskode=nedgrad&Kontakt=<% =rsVikar("BestilltAv") %>&SOKontakt=<% =rsVikar("SoBestilltAv") %>&OppdragID=<% '=rsVikar("OppdragID") %>&FirmaID=<% =rsVikar("FirmaID") %>&Fakturadato=<% =rsVikar("Fakturadato") %>"   >Nedgrad</A></td>
											<% 
										End If 
										rsVikar.MoveNext
										%>
									</tr>
									<% 
								Loop
								rsVikar.Close
								set rsVikar = nothing
								%>
							</table>
						</div>
						<%
						If viskode = 2 Then 
							%>
							<form name="fire" ACTION="Faktura_overf_ordrehode.asp" METHOD=POST ID="Form5">
								<input name="btnOverfoer" TYPE="SUBMIT" VALUE="Lag fil til Rubicon" ID="Submit5">
								<p>Siste ordrenr: <input name="ONR" TYPE="TEXT" SIZE="5" ID="Text5"></p>
							</form>
							<% 
						End If
						If viskode = 1 Or viskode = 3 Then 
							%>
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
							<% 
						End If 
					End If 'ingen treff
				End If 'viskode > 0 
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>