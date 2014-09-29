<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Integration.Contact.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Function OnlyDigits( strString )
		dim idx
		dim strNewstring
		dim Digit
		' Remove all non-nummeric signs FROM string
		If Not IsNull(strString) Then
			For idx = 1 To Len( strString) Step 1
				Digit = Asc( Mid( strString, idx, 1 ) )
				If (( Digit > 47 ) AND ( Digit < 58 )) Then
					strNewstring = strNewString & Mid(strString, idx, 1)
				End If
			Next
		End If
		Onlydigits = strNewString
	End Function

	dim strSQL
	dim Conn
	dim rsNavn
	dim strKontaktNavn
	dim strKontaktID
	dim SOKontaktID
	dim strFirmaNavn
	dim rsFaktLinjer
	dim KontaktSQL
	dim erXisKontakt : erXisKontakt = false
	
%>	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<title>Faktura</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<!--#INCLUDE FILE="../includes/Library.inc"-->
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
		<%
		' Get a database connection
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))

		' prosessing parameters
		strKontaktID = Request.QueryString("Kontakt")
		SOKontaktID = Request.QueryString("SOKontakt")
		strFirmaId = Request.QueryString("FirmaID")
		vikarId = Request.QueryString("vikarID")
		strEndre = Request.QueryString("Endre")
		strNy = Request.QueryString("Ny")
		strLInje = Request.QueryString("Linje")
		StartDato = Request.QueryString("StartDato")
		strAvdeling = Request.QueryString("Avdeling")
		splittuke = Request.QueryString("splittuke")
		splittopt2 = Request.QueryString("splittopt2")
		splittopt = Request.QueryString("splittopt")

		' check parameters
		'se strKontaktID & "<br>"
		'se strFirmaID & "<br>"
		'se strEndre & "<br>"
		'se strLinje & "<br>"

		if (len(strKontaktID) = 0) then
			kontaktSQL = " SOKontakt = " & SOKontaktID
			erXisKontakt = false
		else
			kontaktSQL = " Kontakt = " & strKontaktID
			erXisKontakt = true
		end if

		if (erXisKontakt = false) then
			strKontaktNavn = GetSOPersonName(SOKontaktID)
		else
			' SQL for finding firma AND kontakt
			strSQL = "SELECT Navn=(Fornavn + ' ' + Etternavn) FROM KONTAKT WHERE KontaktID = " & strKontaktID
			Set rsNavn = GetFirehoseRS(strSQL, Conn)
			strKontaktNavn = rsNavn("Navn")
			rsNavn.Close
			set rsNavn = nothing
		end if

		strSQL = "SELECT Firma FROM Firma WHERE FirmaID = " & strFirmaID
		Set rsNavn = GetFirehoseRS(strSQL, Conn)
		strFirmaNavn = rsNavn("Firma")
		rsNavn.Close
		set rsNavn = nothing

		' SQL to get data
		strSQL = "SELECT Linje = FakturaLinjeID, Kontakt, SOKontakt, FirmaID, OppdragID, ArtikkelNr, VikarID, " &_
			"Tekst, Antall, Status, Enhetspris, LinjeSum, LinjeNr, Split, NyLinje " &_
			", Fakturanr, Fakturadato " &_
			"FROM FAKTURAGRUNNLAG " &_
			"WHERE" &_
			kontaktSQL &_
			" AND Avdeling = " & strAvdeling  &_
			" AND Status < 3 " &_
			" ORDER BY VikarID, LinjeNr"

		Set rsFaktLinjer = GetFirehoseRS(strSQL, Conn)

		If rsFaktLinjer.EOF Or rsFaktLinjer.BOF Then
			AddErrorMessage("Ingen fakturalinjer for " & strFirmaNavn & ", " & strKontaktNavn )
			call RenderErrorMessage()
		Else
			strOppdragID = rsFaktLinjer("OppdragID")
			' form to register AND change

			' If edit find the right row to display
			If strLinje <> "" Then ' edit
				strSQL = "SELECT Linje=FakturaLinjeID, VikarID, Kontakt, SOKontakt, FirmaID, OppdragID, ArtikkelNr, VikarID, Tekst, Antall, Enhetspris, LinjeSum " &_
				"FROM FAKTURAGRUNNLAG " &_
				"WHERE FakturaLinjeID = " & strLinje
               response.write(strSql)
				Set rsLinje = GetFirehoseRS(strSQL, Conn)
			End If
			%>
			<div class="contentHead1"><h1><% =strFirmanavn %>, <% = strKontaktNavn %></h1></div>
			<div class="content"></div>
			<%
			' form to register AND change
			If strEndre = "Endre" Or strEndre = "Ny" Or strEndre = "Insert" Then
				%>
				<div class="contentHead"><h2>Redigere linje</h2></div>
				<div class="content">
					<form NAVN="EDITER" ACTION="Faktura_DB.asp?Linje=<% =strLinje %>&Endre=<% =strEndre %>&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>" METHOD="POST" ID="Form1">
						<input name="VikarID" TYPE="HIDDEN" VALUE="<%=vikarid%>" ID="Hidden1">
						<input name="LinjeNr" TYPE="HIDDEN" VALUE="<%=Request("LinjeNr")%>" ID="Hidden2">
						<input name="split" TYPE="HIDDEN" VALUE="<%=rsFaktlinjer("split")%>" ID="Hidden3">
						<input name="Fakturanr" TYPE="HIDDEN" VALUE="<%=rsFaktlinjer("Fakturanr")%>" ID="Hidden4">
						<input name="Fakturadato" TYPE="HIDDEN" VALUE="<%=rsFaktlinjer("Fakturadato")%>" ID="Hidden5">
						<input name="Avdeling" TYPE="hidden" VALUE="<%=Request("Avdeling") %>" ID="Hidden6">
						<input name="Status" TYPE="hidden" VALUE="<% =rsFaktlinjer("Status") %>" ID="Hidden7">
						<table cellpadding='0' cellspacing='1' ID="Table1">
							<tr>
								<th>Artikkelnr</th>
								<th>Tekst</th>
								<th>Antall</th>
								<th>Enhetspris</th>
							</tr>
							<tr>
								<td><input type="text" size="6" NAME="ArtikkelNr" maxlength="10" <% If strLinje <> "" Then Response.Write "VALUE='" & rsLinje("ArtikkelNr") & "'" %>  ID="Text1"></td>
								<td><input type="text" size="40" NAME="Tekst" maxlength="40" <% If strLinje <> "" Then Response.Write "VALUE='" & rsLinje("Tekst") &"'" %> ID="Text2"></td>
								<td><input type="text" size="6" NAME="Antall" maxlength="6" <% If strLinje <> "" Then Response.Write "VALUE='" & rsLinje("Antall") & "'" %>  ID="Text3"></td>
								<td><input type="text" size="6" NAME="Enhetspris" maxlength="6" <% If strLinje <> "" Then Response.Write "VALUE='" & rsLinje("Enhetspris") & "'" %> ID="Text4"></td>
							</tr>
						</table>
						<INPUT TYPE=SUBMIT VALUE="Lagre" ID="Submit1" NAME="Submit1">&nbsp;<INPUT TYPE=RESET value="Tilbakestill" ID="Reset1" NAME="Reset1">
					</form>
					<br>
				</div>
				<% 
				If strLinje <> "" Then 
					rsLinje.Close 
				end if
			End If  ' Endre = ja

			' Display data
			status = rsFaktLinjer("Status")
			%>
			<div class="content">
				<div class="listing">
					<table cellpadding='0' cellspacing='1' ID="Table2">
						<tr>
							<th>Endre</th>
							<th>Artikkelnr</th>
							<th>Tekst</th>
							<th>Antall</th>
							<th>Pris</th>
							<th>Sum</th>
							<th>Slett</th>
							<th><A HREF="Faktura_Vis.asp?Endre=Insert&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&LinjeNr=0&VikarID=<% =rsFaktLinjer("VikarID") %>&avdeling=<% =strAvdeling %>" >Ny linje</A></th>
						</tr>
						<%
						VID = rsFaktLinjer("VikarID")
						VikarerHer = VID
						splitOption = False
						faktNr = rsFaktLinjer("Fakturanr")
						sumsum = 0
						do while NOT rsFaktLinjer.EOF
							If rsFaktLinjer("VikarID") <> VID Then
								splitOption = True
								VID = rsFaktLinjer("VikarID")
								VikarerHer = VikarerHer & "," & VID
								If rsFaktLinjer("split") = 1 Then 
									%>
									<tr>
										<td></td><td></td><th class="right">Sum</th><td></td><td></td><th class="right"><% =sumsum %><% sumsum = 0 %></th>
									</tr>
									<tr>
										<th colspan="8"><% =strFirmanavn %>, <% = strKontaktNavn %> </th>
									</tr>
									<tr>
										<th>Endre</th>
										<th>Artikkelnr</th>
										<th>Tekst</th>
										<th>Antall</th>
										<th>Pris</th>
										<th>Sum</th>
										<th>Slett</th>
										<td><A HREF="Faktura_Vis.asp?Endre=Insert&avdeling=<% =strAvdeling %>&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&VikarID=<% =rsFaktLinjer("VikarID") %>" >Ny linje</A></td>
									</tr>
									<%
								End If
							End If
							If Not IsNull(rsFaktLinjer("LinjeSum")) Then
								sumsum = sumsum + rsFaktLinjer("LinjeSum")
							End If
							%>
							<tr>
								<td>
								<% 
								If rsFaktLinjer("NyLinje") = 1 Then 
									%>
									<a HREF="Faktura_Vis.asp?Endre=Endre&Linje=<% =rsFaktLinjer("Linje") %>&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&VikarID=<% =rsFaktLinjer("VikarID") %>&Avdeling=<% =strAvdeling %>" >Endre</a>
									<% 
								Else 
									%>
									<a HREF="Vikar_timeliste_vis3.asp?VikarID=<% =VID %>&OppdragID=<% =rsFaktlinjer("OppdragID") %>&frakode=3"  >Timeliste</a>
									<% 
								End If 
								%>
								</td>
								<td><% =rsFaktLinjer("ArtikkelNr") %></td>
								<% 
								txt = UCase(rsFaktLinjer("Tekst"))
								If Left(txt, 3) = "KON" Or Left(txt, 3) = "VIK" Then 
									%>
									<td><a HREF="Faktura_vis.asp?Endre=Endre&Linje=<% =rsFaktLinjer("Linje") %>&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&VikarID=<% =rsFaktLinjer("VikarID") %>&Avdeling=<% =strAvdeling %>"><% =txt %></a></td>
									<% 
								Else 
									%>
									<td><% =txt %></td>
									<% 
								End If 
								%>
								<td class=right><% =rsFaktLinjer("Antall") %></td>
								<td class=right><% =rsFaktLinjer("Enhetspris") %></td>
								<td class=right><% =rsFaktLinjer("LinjeSum") %></td>
								<td class="center">
								<% 
								If rsFaktLinjer("NyLinje") = 1 Then 
									%>
									<a href="Faktura_DB.asp?Endre=Slett&Linje=<% =rsFaktLinjer("Linje") %>&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&Avdeling=<% =strAvdeling %>" ><img src="/xtra/Images/icon_delete.gif" alt="Slette linje"></a>
									<% 
								End If 
								%>
								</td>
								<td><a HREF="Faktura_Vis.asp?Endre=Insert&Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&LinjeNr=<% =rsFaktLinjer("LinjeNr") %>&VikarID=<% =rsFaktLinjer("VikarID") %>&avdeling=<% =strAvdeling %>" >Ny linje</a></td>
							</tr>
							<%
							delt = rsFaktLinjer("split")
							rsFaktLinjer.MoveNext
						loop
						%>
						<tr>
							<th>
								<% 
								If status = 1 Then 
									%>
									<font class="warning">
									<% 
								ElseIf status = 2 Then 
									%>
									<font>
									<% 
								End if
								%>
								||||||||||||||||</font>
							</th>
							<th colspan="4" class="right">Sum netto:</th>
							<th class="right"><% =sumsum %></th>
							<th colspan="2">&nbsp;</th>
						</tr>
						<%
					End If
					%>
				</table>
				<p>
					<form ACTION="Faktura_lagre.asp?Kontakt=<% =strKontaktID %>&SOKontakt=<% =SOKontaktID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>&kode=1&fakturanr=<% =faktNr %>&VikarerHer=<% =VikarerHer %>&Avdeling=<% =strAvdeling %>" METHOD=POST ID="Form2">
						<INPUT TYPE=SUBMIT VALUE="Godkjenn faktura" ID="Submit2" NAME="Submit2">
						<% If splitOption Then %>
							<% If delt = 1 Then %>
								<INPUT class="checkbox" TYPE="CHECKBOX" NAME="splitt" VALUE="Ja" CHECKED ID="Checkbox1">Splitt faktura?
							<% Else %>
								<INPUT class="checkbox" TYPE="CHECKBOX" NAME="splitt" VALUE="Ja" ID="Checkbox2">Splitt faktura?
							<% End If %>
						<% End If %>
					</form>
				</p>
				<p>
					<input type="button" value="Til behandle timelister" onclick="javascript:window.self.close()" ID="Button1" NAME="Button1">
				</p>
				<br>
				<%
				strSQL = "SELECT DISTINCT OppdragID " &_
					"FROM FAKTURAGRUNNLAG " &_
					"WHERE " &_
					KontaktSQL &_
					" AND Avdeling = " & strAvdeling  &_
					" AND Status < 3 "

					Set rsBeskjed = GetFirehoseRS(strSQL, Conn)

					If hasRows(rsBeskjed) Then
						do while Not rsBeskjed.EOF
							strSQL = "SELECT Notatokonomi " &_
							"FROM OPPDRAG " &_
							"WHERE OPPDRAG.OppdragID = " & rsBeskjed("OppdragID")

							Set rsBeskjed2 = GetFirehoseRS(strSQL, Conn)
							If Not rsBeskjed2.EOF Then
 								beskjed = Trim(rsBeskjed2("Notatokonomi"))
 								If beskjed <> "" Then
									Response.Write "Til økonomi: (" & rsBeskjed("OppdragID") & ") " & beskjed
 								End If
							End If
							rsBeskjed.MoveNext
						loop
						rsBeskjed.Close
					End If 'no rows (beskjder)
					Set rsBeskjed = Nothing					
					
					if (erXisKontakt = true) then
						KontaktSQL = "D.BestilltAv = " & strKontaktID
					else
						KontaktSQL = "D.SOBestilltAv = " & SOKontaktID
					end if
					
					strSQL = "SELECT DISTINCT VikarID, D.OppdragId FROM " &_
						"DAGSLISTE_VIKAR D, OPPDRAG " &_
						"WHERE " &_
						KontaktSQL &_
						" AND D.OppdragID = OPPDRAG.OppdragID" &_
						" AND Fakturastatus < 3 " &_
						" AND OPPDRAG.AvdelingID = " & strAvdeling &_
						" AND Dato < " & dbDate(session("limitDato")) &_
						" ORDER BY VikarID"

					set rsOppdrag = GetFireHoseRS(strSQL, Conn)
					If (HasRows(rsOppdrag)) Then
						%>
						<br>
						<p><strong>Vikarer/timelister:</strong>&nbsp;
						<%
						do while Not rsOppdrag.EOF
							If CStr(rsOppdrag("VikarID")) = strVikarID Then %>
								| <% =rsOppdrag("VikarID") %>
							<% Else %>
								| <A HREF="Vikar_timeliste_vis3.asp?VikarID=<% =rsOppdrag("VikarID") %>&OppdragID=<% =rsOppdrag("OppdragID") %>&frakode=<% =frakode %>" ><% =rsOppdrag("VikarID") %></A>
							<% End If
							rsOppdrag.MoveNext
						loop
						%> |</p><%
						rsOppdrag.Close
					End If 'ingen linjer i rsOppdrag
					set rsOppdrag = Nothing				
					%>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>