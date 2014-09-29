<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Reports.Utils.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim Conn
	dim faktFradato
	dim faktTildato
	Dim AntallAvd
	
	Function toDecimaler(tall)
		txt = CStr(tall)
		If txt = "" then
			txt = "0"
		End if

		pos = InStr(txt,",")
		lengde = Len(txt)
		If pos > 0 Then
			If pos = lengde-1 Then
				txt = txt & "0"
			ElseIf pos = lengde-3 Then
				txt = Mid(txt, lengde-1)
			End If
		Else
			txt = txt & ",00"
		End If

		toDecimaler = txt
	End Function

	Function MAXAntAvd()
		Set rsAntallAvd = conn.Execute ("SELECT storste=MAX(avdelingID) FROM avdeling")
		If HasRows(rsAntallAvd) then
			MAXAntAvd = rsAntallAvd("storste")
			rsAntallAvd.close
		Else
			MAXAntAvd = 1
		End if
		Set rsAntallAvd = Nothing
	End function


	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	AntallAvd = MAXAntAvd()
	faktFradato = Request("faktFradato")
	faktTildato = Request("faktTildato")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<title>Kontrollrapport lønn</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript" src="../Js/javaScript.js"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Kontrollrapport for overføring av faktura <% =date %></h1>
			</div>
			<div class="content">
				<FORM NAME="FromEn" ACTION=Faktura_kontroll_vis.asp METHOD=POST>
					<table>
						<tr>
							<td>Fra dato: <INPUT TYPE="TEXT" size="8" maxlength="8" NAME="faktFradato" VALUE="<% If faktFradato <> "" Then Response.Write faktFradato %>" ONBLUR="dateCheck(this.form, this.name)" ></td>
							<td>Til dato: <INPUT TYPE="TEXT" size="8" maxlength="8" NAME="faktTildato" VALUE="<% If faktTildato <> "" Then Response.Write faktTildato %>" ONBLUR="dateCheck(this.form, this.name)" ></td>
							<td><INPUT TYPE=SUBMIT VALUE="Søk"></td>
						</tr>
					</table>
				</form>
				<%
				If faktFradato <> "" AND faktTildato <> "" Then

					Dim Total_timer(32)
					Dim Total_utlegg(32)

					For i = 1 To AntallAvd
						avd = UCase(HentAvdNavn(i))

						Response.Write "<h2><strong>Avdeling:</strong> " & avd & "<br>Fakturagrunnlag overført til Rubicon i tidsrommet " & faktFradato & " - " & faktTildato & "</h2>"
						avdnr = i

						' sql for faktura
						strSQL = "SELECT F.FirmaID, FI.Firma," &_
						" sum_timer=SUM(case WHEN F.artikkelnr <> '999910' then f.linjesum else 0 end)"&_
						", sum_utlegg=SUM(case WHEN F.artikkelnr = '999910' then f.linjesum else 0 end)" &_
							" FROM FAKTURAGRUNNLAG F, FIRMA FI" &_
							" WHERE  status = 3" &_
							" AND Fakturadato >= " & dbDate(faktFradato) &_
							" AND Fakturadato <= " & dbDate(faktTildato) &_
							" AND F.FirmaID = FI.FirmaID" &_
							" AND F.Avdeling = " & avdnr &_
							" GROUP BY F.FirmaID, FI.Firma"

						set rsFakt = GetFirehoseRS(strSQL, Conn)
						tot_timer = 0
						tot_utlegg = 0

						' Vis data
						%>
						<div class="listing">
							<table>
								<tr>
									<th>kontaktid</th>
									<th>Kontakt</th>
									<th>Timer</th>
									<th>Viderefakt. utlegg</th>
									<th>Totalt</th>
								</tr>
								<%
								If HasRows(rsFakt) Then
									do while not rsFakt.EOF 
										%>
										<tr>
											<td><% =rsFakt("FirmaID") %></td>
											<td><% =rsFakt("Firma") %></td>
											<td class="right"><% =toDecimaler(rsFakt("sum_timer")) %></td>
											<td class="right"><% =toDecimaler(rsFakt("sum_utlegg")) %></td>
											<td class="right"><% =toDecimaler(rsFakt("sum_utlegg")+rsfakt("sum_timer")) %></td>
										</tr>
										<%
										Tot_timer = tot_timer + rsFakt("sum_timer")
										Tot_utlegg = tot_utlegg + rsFakt("sum_utlegg")
										rsFakt.MoveNext
									loop
									Total_timer(i) = tot_timer
									Total_utlegg(i) = tot_utlegg
									rsFakt.Close
								End If ' ingen rader
								Set rsFakt = Nothing
								%>
								<tr>
									<td colspan="2"><strong>SUM FAKTURERT</strong></td>
									<td class="right"><% =toDecimaler(tot_timer) %></td>
									<td class="right"><% =toDecimaler(tot_utlegg) %></td>
									<td class="right"><% =toDecimaler(tot_timer + tot_utlegg) %></td>
								</tr>
							</table>
						</div>
						<%
					Next 'avdeling
					%>
					<div class="listing">
						<table>
							<tr>
								<th>Totalt fakturert</th>
								<th>Timer</th>
								<th>Viderefakt. utlegg</th>
								<th> Totalt</th>
							</tr>
							<%
							For teller= 1 to AntallAvd
								%>
								<tr>
									<td> <% =UCase(HentAvdNavn(teller)) %></td>
									<td class="right"><% =toDecimaler(total_timer(teller)) %></td>
									<td class="right"><% =toDecimaler(total_utlegg(teller)) %></td>
									<td class="right"><% =toDecimaler(total_timer(teller)+ total_utlegg(teller) ) %></td>
								</tr>
								<%
								GrandTot_timer = GrandTot_timer + total_timer(teller)
								GrandTot_utlegg = GrandTot_utlegg + total_utlegg(teller)
							next 
							%>
							<tr>
								<td><strong>SUM FAKTURERT</strong></td>
								<td class="right"><% =toDecimaler(GrandTot_timer) %></td>
								<td class="right"><% =toDecimaler(GrandTot_utlegg) %></td>
								<td class="right"><% =toDecimaler(GrandTot_timer + GrandTot_utlegg ) %></td>
							</tr>
						</table>
					</div>
					<%
				End If 'dato er lagt inn 
				%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>