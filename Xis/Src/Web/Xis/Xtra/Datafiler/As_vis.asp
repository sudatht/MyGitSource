<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit 
Response.Expires = 0
%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	' NB!!! oppgraderingen av lønnstatus på AS er fullstendig omarbeider av Kjetil Borg  27.12.99. De 
	' eneste filene som  nå er i bruk for denne funksjonen, er As_vis.asp og AS_oppgrad_Lstat_DB.asp
	' Oppgraderingen skjer nå slik at _alle_ linjer i Dagsliste_vikar og vikar_ukeliste oppgraderes til 
	' loennstatus=3 når timelistene har timelistestatus=5. Dette skjer i bulk fordi AS ikke skal inn i 
	' H&L, men blir lønnet ved utbetaling av faktura som andre kreditorer. Når timelistestatus=5 innebærer 
	' dette at faktura fra konsulenten er ferdig kontrollert og godkjent. Alle linjer med loennstatus=1 og 
	' timelistestatus=5 kan derfor trygt oppgraderes til loennstatus=3. Dette i henhold til spesifikasjon 
	' fra XP.

	' declare variables
	Dim Conn
	Dim strSQL
	Dim rsAS

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' søkeresultat
	strSQL = "SELECT distinct D.VikarID, V.Etternavn, V.Fornavn, VIKAR_ANSATTNUMMER.ansattnummer,  maxDato=max(D.dato) " &_
		" FROM DAGSLISTE_VIKAR D, VIKAR V LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON V.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " &_
		" WHERE D.TimelisteVikarStatus = 5 "  &_
		" AND D.Loennstatus = 1 " &_
		" AND V.TypeID = 3 " &_
		" AND D.VikarID = V.VikarID" &_
		" GROUP BY D.VikarID, v.etternavn, v.fornavn, VIKAR_ANSATTNUMMER.ansattnummer "&_
		" ORDER BY V.Etternavn"

	set rsAS = conn.execute(strSQL)
	%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"  "http://www.w3.org/TR/REC-html40/loose.dtd">
	<html>
		<head>
			<title>AS</title>
			<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
			<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		</head>
		<body>
			<div class="pageContainer" id="pageContainer">
				<div class="contentHead1">
					<h1>AS-vikarer som ligger med godkjente timelister (T.stat=5) og ikke godkjent lønn (L.stat = 1)</h1>
				</div>
				<div class="content">
					<%
					If hasRows(rsAS) = false Then 
						response.write "<p>Det er ingen AS-vikarer som ligger med T.stat=5 og L.stat = 3</p>"
					else
						%>			
						<div class="listing">
							<form action="AS_oppgrad_Lstat_DB.asp" method="post">
								<table>
									<tr>
										<th>Ansattnummer</th>
										<th>Vikar</th>
										<th>Til Dato</th>
									</tr>
									<%
									' søkeresultat
									do while not rsAS.EOF 
										%>
										<tr>
											<td>
												<input type="hidden" name="vikarID" value="<%=rsAS("VikarID")%>">
												<% =rsAS("ansattnummer") %>
											</td>
											<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsAs( "VikarID" ), rsAs("Etternavn") & " " & rsAS("Fornavn"), "Vis vikar " & rsAs("Etternavn") & " " & rsAS("Fornavn") )%></td>
											<td><% =rsAs("maxDato") %></td>
										</tr>
										<%
										rsAs.MoveNext
									loop
									rsAS.Close
									%>
								</table><br>
								<input type="submit" value="Oppgrader lønnstatus til 3" id="Submit1" name="Submit1">	
							</form>
						</div>					
						<%
					End If 'rader
					Set rsAS = Nothing
					%>
				</div>
			</div>
		</body>
	</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>