<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
%>
<html>
	<head>
		<title>Ukeliste</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">		
		<script language="javaScript" src="Js/javaScript.js" type="text/javascript"></script>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Sammendrag</h1>
				<div class="contentMenu">
					&nbsp;
				</div>
			</div>
			<div class="content">
				<%
				' Connect to database
				Set Conn = GetConnection(GetConnectionstring(XIS, ""))

				' prosessing parameters
				strVikarID = Request("VikarID")
				strOppdragID = Request("OppdragID")
				strFirmaID = Request("FirmaID")
				strId = Request("ID")
				kode = Request("kode")
				tilgang = Request("tilgang")
				strUkeNr = Request("UkeNr")
				endre = Request("endre")

				' SQL for displaying data
				strSQL = "SELECT Id, VikarId, OppdragID, FirmaID, UkeNr, StatusID, " &_
					"Antall, Sats, Belop, Dato, Loenndato, bestilltAv, SOBestilltAv, notat, " &_
					"Fakturatimer, Fakturapris, Fakturabeloep, Fakturadato, " &_
					"Loennsartnr, Overfort_fakt_status, Overfort_loenn_status, BonusStatus" &_
					" FROM VIKAR_UKELISTE" &_
					" WHERE VikarID = " & strVikarID &_
					" and OppdragID = " & strOppdragID &_
					" and StatusID < 6" &_
					" and StatusID > 1" &_
					" order by UkeNr, Loennsartnr DESC"

				Set rsVikar = GetFirehoseRS(strSQL, Conn)

				' SQL for finding name
				strSQL = "SELECT Navn=(Fornavn + ' ' + Etternavn) FROM [VIKAR] WHERE [Vikarid] = " & strVikarID

				Set rsNavn = GetFirehoseRS(strSQL, Conn)

				strNavn = rsNavn("Navn")
				rsNavn.Close
				set rsNavn = nothing

				' If no record exsists
				If (not hasRows(rsVikar)) Then
					Response.Write "<p class=""warning"">Ingen godkjente timelister!</p>"				
				Else 
					' Display data
					%>
					<H3>Sammendrag (ukeliste) for <% =strNavn %></H3>
					<div class="listing">
						<table border="1" ID="Table1">
							<tr class="right">
								<th WIDTH=30>Ukenr</th>
								<th>D</th>
								<th WIDTH=30>L.art</th>
								<th WIDTH=50>L.timer</th>
								<th WIDTH=30>L.sats</th>
								<th WIDTH=20>Lønn</th>
								<th WIDTH=20>L.stat</th>
								<th WIDTH=30>L.dato</th>
								<th WIDTH=20>F.timer</th>
								<th WIDTH=20>F.sats</th>
								<th WIDTH=20>Fakt</th>
								<th WIDTH=20>F.stat</th>
								<th WIDTH=30>F.dato</th>
								<th WIDTH=20>T.stat</th>
								<th WIDTH=20>Bonus</th>
							</tr>
							<%
							ukeTeller = rsVikar("Ukenr")
							sum = 0
							sumsum = 0
							fsum = 0
							fsumsum = 0
							ftim = 0
							ltim = 0
							i=0
							' LOOP
							do while not rsVikar.EOF 
								If rsVikar("notat") = "Null" Then delt = "N" Else delt = rsVikar("Notat")
								
								If ukeTeller <> rsVikar("Ukenr") Then 
									ukeTeller = rsVikar("Ukenr")
								End If

								ltim = ltim + rsVikar("Antall")
								ftim = ftim + rsVikar("Fakturatimer")
								sum = sum + rsVikar("Belop")
								fsum = fsum + rsVikar("Fakturabeloep")
								%>
								<TR class=right>
									<th><%=rsVikar("Ukenr")%>&nbsp;</th>
									<th><% =Ukedel%>&nbsp;</th>
									<th><%=rsVikar("Loennsartnr")%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("Antall"),2)%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("Sats"),2)%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("Belop"),2)%>&nbsp;</th>
									<th><%=rsVikar("Overfort_loenn_status") %>&nbsp;</th>
									<th><%=rsVikar("Loenndato")%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("Fakturatimer"),2)%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("Fakturapris"),2)%>&nbsp;</th>
									<th><%=formatNumber(rsVikar("FakturaBeloep"),2)%>&nbsp;</th>
									<th><%=rsVikar("Overfort_fakt_status") %>&nbsp;</th>
									<th><%=rsVikar("Fakturadato")%>&nbsp;</th>
									<th><%=rsVikar("StatusID") %>&nbsp;</th>
									<th><%=rsVikar("BonusStatus")%>&nbsp;</th>
								</tr>
								<%
								rsVikar.MoveNext
							loop
							' END LOOP
							%>
							<tr>
								<th>Sum:</th>
								<th></th>
								<th></th>
								<th class="right"><%=formatNumber(ltim,2)%></th>
								<th></th>
								<th class="right"><%=formatNumber(sum,2)%></th>
								<th><th></th>
								<th class="right"><%=formatNumber(ftim,2)%></th>
								<th></th>
								<th class="right"><%=formatNumber(fsum,2)%></th>
							</tr>
						</table>
					</div>							
					<%  
					rsVikar.Close 
				End If 
				%>
				<INPUT TYPE=Button onClick="javascript:history.back()" VALUE="                  Tilbake                        " ID="Button1" NAME="Button1">
				<% 
				set rsVikar = nothing
				CloseConnection(Conn)
				set Conn = nothing
				%>
			</div>
		</div>
	</body>
</html>