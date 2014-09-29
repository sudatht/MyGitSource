<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Settings.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'Funksjonen godkjenner skattekort for innværende år, samt også for foregående år hvis det er Januar måned.
	'Eller frem til 12 Februar.
	Function godkjennSkattekort(strSkattekort)
		Dim bGodkjentSkattekort
		Dim dtIDag
		Dim dtAar
		Dim dtMaaned
		Dim dtDag

		bGodkjentSkattekort = False
		dtIDag = Date
		dtAar = Year(dtIDag)
		dtMaaned = Month(dtIDag)
		dtDag = Day(dtIDag)

		If Not (IsNull(strSkattekort) Or strSkattekort = "" Or strSkattekort = "-1") Then
			' Sjekker om skattekort er levert for i år eller for i fjor og det er Januar
			If (dtAar = CInt(strSkattekort) Or (dtAar - 1 = CInt(strSkattekort) AND dtMaaned = 1)) Then
				bGodkjentSkattekort = True
			' Sjekker om skattekort er levert for ifjor og det Februar(før den 12.)
			ElseIf (dtAar - 1 = CInt(strSkattekort) AND dtMaaned = 2 AND dtDag <= 12) Then
				bGodkjentSkattekort = True
			End If
		End If
		godkjennSkattekort = bGodkjentSkattekort
	End Function

	
	Dim strAnsattnummer		'As String
	Dim bGodkjennLoenn		'As Boolean
	Dim Conn
	dim rsAvdeling	
	dim rsVikar
	dim viskode
	dim frakode2
	dim OppdragID
	
	viskode = session("viskode")
	frakode2 =  session("frakode2")
%>
<html>
	<head>
		<title>Lønn</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<%

			Set Conn = GetConnection(GetConnectionstring(XIS, ""))
			
			' prosessing parameters
			strVikarID = Request("VikarID")
			viskode = Request("viskode")
			strOppdragID = Request("OppdragID")
			strFirmaID = Request("FirmaID")
			status = Request("status")
			Loenndato = Request("LoennDato")
			If Request("frakode2") <> "" Then
				session("frakode2") = Request("frakode2")
			End If
			If Request("OppdragID") <> "" Then
				session("OppdragID") = Request("OppdragID")
			End If

			' SQL for finding substitute information
			strSQL = "SELECT " & _
				"Navn = (VIKAR.Fornavn + ' ' + VIKAR.Etternavn), " & _
				"VIKAR.MottattSkattekort, " & _
				"VIKAR_ANSATTNUMMER.ansattnummer " & _
				"FROM VIKAR " & _
				"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
				"WHERE VIKAR.Vikarid = " & strVikarID

			Set rsNavn = GetFirehoseRS(strSQL, Conn)
			strNavn = rsNavn("Navn").Value
			strAnsattnummer = rsNavn("ansattnummer").Value
			bGodkjennLoenn = godkjennSkattekort(rsNavn("MottattSkattekort").Value)
			rsNavn.Close
			set rsNavn = nothing

			session("FirmaID") = Request("FirmaID")

			' SQL for displaying loennsart
			strSQL = "SELECT Loennsartnr, Loennsart FROM H_LOENNSART ORDER BY Loennsartnr"
			Set rsLoennsart = GetFirehoseRS(strSQL, Conn)

			' SQL for displaying avdeling
			strSQL = "SELECT a.AvdelingID, a.Avdeling FROM Avdeling a ORDER BY a.avdnr"
			Set rsAvdeling = GetFirehoseRS(strSQL, Conn)

			' If edit find the right row to display
			strID = Request("ID")
			strLoennsartnr = ""
			strAvdeling = ""
			strEndre = Request("Endre")

			If strID <> "" Then
				strSQL = "SELECT Dato, Prosjektnr, Loennsartnr, Antall, Sats, Beloep, avdeling " &_
					"FROM VIKAR_LOEN_VARIABLE " &_
	 				"WHERE ID = " & Request("ID")
				
				Set rsID = GetFirehoseRS(strSQL, Conn)
				
				strLoennsartnr = rsID("Loennsartnr")
				strAvdeling = rsID("avdeling")
			End If

			' Form to register faste lønnsdata
			If strEndre = "Ja" or strEndre = "Ny" Then
			%>	
				<div class="content">
					<form ACTION="Vikar_varl_DB3.asp?Ny=Ja" METHOD="POST">
						<input name="VikarID" TYPE=HIDDEN VALUE=<% =strVikarID %> ID="Hidden1">
						<input name="OppdragID" TYPE=HIDDEN VALUE=<% =strOppdragID %> ID="Hidden2">
						<input name="Navn" TYPE=HIDDEN VALUE="<% =strNavn %>" ID="Hidden3">
						<input name="Endre" TYPE=HIDDEN VALUE="<% =strEndre %>" ID="Hidden4">
						<input name="ID" TYPE=HIDDEN VALUE="<% =strID %>" ID="Hidden5">
						<input name="status" TYPE=HIDDEN VALUE="<% =status %>" ID="Hidden6">
						<input name="Loenndato" TYPE=HIDDEN VALUE="<% =Loenndato %>" ID="Hidden7">
				
						<div class="listing">
							<table cellspacing="1" cellpadding="0" ID="Table1">
								<tr>
									<th>Art</th>
									<th>Antall</th>
									<th>Sats</th>
									<th>Regnskapsavd.</th>
								</tr>
								<tr>
									<th>
										<SELECT NAME="Loennsartnr" ID="Select1">
											<option VALUE= ""></option>
											<%
											Do Until rsLoennsart.EOF
												If strLoennsartnr = rsLoennsart("Loennsartnr") Then 
													sel = " SELECTED" 
												Else 
													sel = "" 
												end if
												%>
												<option VALUE=<% =rsLoennsart("Loennsartnr") %><% =sel %> ><% =rsLoennsart("Loennsartnr")& " " & rsLoennsart("Loennsart")%></option>
												<%
												rsLoennsart.MoveNext
											loop
											rsLoennsart.Close
											set rsLoennsart = nothing
											%>
   										</SELECT>
									</th>
									<th><input type=text size=4 NAME=Antall <% If strID <> "" Then response.Write "VALUE=" & rsID("Antall")%> ID="Text1"></th>
									<th><input type=text size=6 NAME=Sats <% If strID <> "" Then response.Write "VALUE=" & rsID("Sats")%> ID="Text2"></th>
									<th>
										<SELECT NAME="Avdeling" ID="Select2">
										<OPTION VALUE= "">
										<%
										Do Until rsAvdeling.EOF
											If strAvdeling = rsAvdeling("AvdelingID") Then sel = " SELECTED" Else sel = "" %>
											<OPTION VALUE=<% =rsAvdeling("AvdelingID") %><% =sel %> ><% =rsAvdeling("Avdeling")%>
											<%
											rsAvdeling.MoveNext
										loop
										rsAvdeling.Close
										set rsAvdeling = nothing
										%>
										</SELECT>
									</th>
								<tr>
								<th colspan=4 ><INPUT TYPE=SUBMIT  VALUE="                Registrer                " ID="Submit1" NAME="Submit1"><INPUT TYPE=RESET ID="Reset1" NAME="Reset1"></th>
							</table>
						</div>
					</form>
					<%
					If strID <> "" Then rsID.Close
				End If  'endre =ja or ny

				' SQL for displaying data
				If Request("gmlloenn") <> "" Then
					grense = 4
				Else
					grense = 3
				End If

				strSQL = "SELECT Id, Dato, Prosjektnr, VIKAR_LOEN_VARIABLE.Loennsartnr, Loennsart, Antall" &_
					", Sats, Beloep, stat = Overfor_loenn_status, NyLinje, LoennDato, Avdeling.Avdeling" &_
					" FROM VIKAR_LOEN_VARIABLE, AVDELING, H_loennsart " &_
					" WHERE VikarID = " & strVikarID &_
					" AND Overfor_loenn_status < " & grense &_
					" AND VIKAR_LOEN_VARIABLE.Avdeling *= Avdeling.AvdelingID" &_
					" AND VIKAR_LOEN_VARIABLE.Loennsartnr *= H_loennsart.Loennsartnr"
				Set rsVikar = GetFirehoseRS(strSQL, Conn)

				' If no record exsists
				If rsVikar.BOF = True AND rsVikar.EOF = True Then
					Response.Write "<p class='warning'>Ingen variable lønnsopplysninger.</p>"
				Else 
					' Display data
					stat = rsVikar("stat")
					Loenndato = rsVikar("Loenndato")
					%>
					<div class="contenthead1">
						<h1>Variabel lønn: <%=strAnsattnummer%> - <%=strNavn%></h1>
					</div>
					<div class="content">
						<div class="listing">
							<table cellpadding='0' cellspacing='1' ID="Table2">
								<tr>
									<th>Endre</th>
									<th>Artnr</th>
									<th>Artnavn</th>
									<th>Antall</th>
									<th>Sats</th>
									<th>Beløp</th>
									<th>Slett</th>
									<th>Avd</th>
								</tr>
								<%
								do while not rsVikar.EOF
									%>
									<tr>
										<td>
											<% 
											If rsVikar("Nylinje") = 1 Then 
												%>
												<A HREF=Vikar_varl_vis3.asp?Endre=Ja&VikarID=<%=strVikarID %>&OppdragID=<% =strOppdragID %>&viskode=<% =viskode %>&ID=<%=rsVikar("ID")%>&status=<% =stat %>&Loenndato=<% =Loenndato %>>Endre</A>
												<% 
											End If 
											%>
											&nbsp;
										</td>
										<td class="right"><%=rsVikar("Loennsartnr")%></td>
										<td><%=rsVikar("Loennsart")%></td>
										<td class="right"><%=rsVikar("Antall")%></td>
										<td class="right"><%=rsVikar("Sats")%></td>
										<td class="right"><%=rsVikar("Beloep")%></td>
										<% 
										If rsVikar("Nylinje") = 1 Then 
											%>
											<td><A HREF=Vikar_varl_db3.asp?Ny=Ja&ID=<%=rsVikar("ID")%>&OppdragID=<% =strOppdragID %>&Slett=Ja&VikarID=<%=strVikarID%>&viskode=<% =viskode %>>Slett</A></td>
											<% 
										Else 
											%>
											<td>&nbsp;</td>
											<% 
										End If 
										%>
										<td><%=rsVikar("Avdeling")%></td>
									</tr>
									<% 
									rsVikar.MoveNext
								loop 
								rsVikar.Close 
								set rsVikar = nothing
								%>
								<tr>
									<form action="Vikar_varl_vis3.asp?Endre=Ny&VikarID=<%=strVikarID%>&Update=Ja&OppdragID=<%=strOppdragID%>&status=<%=stat%>&Loenndato=<%=Loenndato%>" method="post" ID="Form2">
										<td>&nbsp;</td>
									</form>
									<td class="center">
										<%
										If stat = 1 Then
											%>
											<font color="red">
											<%
										Else 
											%>
											<font color="green">
											<% 
										End If 
										%>
										<strong>|||||||||</strong>
										</font>
									</td>
									<%
									If (frakode2 = 1 OR frakode2 = 2 OR frakode2 = 3) Then 'fra timeliste
										%>
										<FORM ACTION="Vikar_varl_refresh.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&visKode=<% =viskode  %>" METHOD="POST" ID="Form1">
											<td colspan="1" class="center">
												<INPUT TYPE=SUBMIT VALUE="Hent på ny" ID="Submit3" NAME="Submit3">
											</td>
										</form>
										<% 
									else
										%>
										<td colspan="1" class="center">
											<INPUT TYPE="button" disabled VALUE="Hent på ny" >
										</td>
										<%									
									End If 
									%>	
									<form action="Vikar_varl_lagre3.asp?VikarID=<%=strVikarID%>&oppdragID=<%=strOppdragID%>&firmaID=<%=strfirmaID%>" method="post" ID="Form3">
										<td colspan="3" class="center"><input type="submit" value="Godkjenn lønn" <%If Not (bGodkjennLoenn) Then Response.Write "disabled"%> ID="Submit2" NAME="Submit2"></td>
										<%
										If Not (bGodkjennLoenn) Then
											%>
											<td colspan="3" class="center">
												<font color="red">Skattekort ikke mottatt/gyldig
											</td>
											<%
										End If
										%>
									</form>	
									<td colspan="2">
										&nbsp;
									</td>																									
								</tr>
							</table>
						</div>
					<% 
				End If 
				%>
				<p>
					<input type="button" value="Til behandle timelister" onclick="javascript:window.self.close()" ID="btnClose" NAME="btnClose">
				</p>
				<p>
					<form action="Vikar_timeliste_vis3.asp?viskode=2&VikarID=<%=strVikarID%>&OppdragID=<% =strOppdragID%>&frakode=3" method="post">
						<input type="submit" value="Tilbake til timelisten">
					</form>
				</p>
				<%
				strSQL = "SELECT DISTINCT OppdragId FROM DAGSLISTE_VIKAR " &_
					"WHERE vikarID = " & strVikarID &_
					" AND Loennstatus < 3 " &_
					" AND TimelisteVikarStatus = 5" &_
					" AND Dato < " & dbDate(session("limitDato")) &_
					" ORDER BY OppdragID"

				Set rsOppdrag = GetFirehoseRS(strSQL, Conn)
				If Not rsOppdrag.EOF Then
					%>
					<br>
					<p><strong>Timelister/Oppdrag:</strong>&nbsp;

					<% do while Not rsOppdrag.EOF %>
						| <A HREF="Vikar_timeliste_vis3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =rsOppdrag("OppdragID") %>&frakode=<% =frakode %>" ><% =rsOppdrag("OppdragID") %></A>
						<% rsOppdrag.MoveNext
					loop
					rsOppdrag.Close
					set rsOppdrag = Nothing
					%>|</p><%
				End If 'ingen linjer i rsOppdrag

				if trim(strOppdragID) <> "" then
				strSQL = "SELECT DISTINCT NotatOkonomi FROM OPPDRAG " &_
					"WHERE notatOkonomi <> '' AND notatokonomi is not null AND oppdragID = " & strOppdragID
					Set rsBeskjed = conn.Execute(strSQL)
					If Not rsBeskjed.EOF Then
						Response.Write "<br>Beskjeder:<br>"
						do while Not rsBeskjed.EOF
							Response.Write rsBeskjed("NotatOkonomi") & "<br>"
							rsBeskjed.MoveNext
						loop
						rsBeskjed.Close
						Set rsBeskjed = Nothing
					End If 'no rows (beskjeder)
				end if
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>