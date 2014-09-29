<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Response.Expires = 0
%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Integration.Contact.inc"-->
<!--#INCLUDE FILE="..\includes\CRM.Integration.Contact.inc"-->

<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	' Dim'er variabler
	Dim Conn 'connectionobjekt for DB
	Dim firmaID 
	Dim firma 'firmanavn
	Dim postadresse 'firmaadresse
	Dim poststed 'for firma
	Dim strSQL 'variabel for SQL-queries
	Dim avdelingID
	Dim periode 'fakturaperiode
	Dim rsKontakt 'henter kontaktpersoner for kunde med fakturert oppdrag
	Dim rsFirmaInfo 'henter basisopplysninger om kunde
	' henter inputparametre
	firmaID = request("valg")
	avdelingID = request("avdelinger")
	periode = request("periode")

	'oppretter databaseforbindelse
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	' sub for å hente kundeopplysninger
	sub hentKundeInfo()

		strSQL ="SELECT f.firmaid, f.firma, a.adresse, a.postnr, a.poststed "&_
			" FROM firma f, adresse A "&_
			" WHERE f.firmaid IN (" & firmaID & ")" &_
			" AND a.adresseRelID = f.firmaid "

		Set rsFirmaInfo = GetFirehoseRS(strSQL, Conn)
		
		if HasRows(rsFirmaInfo) Then
			firma = rsFirmaInfo("firma")
			postadresse = rsFirmaInfo("adresse")
			poststed = rsFirmaInfo("postnr") & "  " & rsFirmaInfo("poststed")
			%>
			<div class="contentHead1">
				<h1><%=firma%></h1>
			</div>
			<div class="content">
				<%=postadresse & "  " & poststed%>
			<%			
			rsFirmaInfo.Close
		end if
		Set rsFirmaInfo = nothing
	end sub

	' sub for å hente utvalg med kontaktpersoner med oppdrag hos kunde 
	sub hentKontakt_oppdrag()
		strSQL = " SELECT DISTINCT v.firmaid, f.firma, O.bestilltAv, O.SOPeID, "&_
			" kontakt=(K.fornavn+' '+ K.etternavn ) "&_
			" FROM firma f, vikar_ukeliste V, oppdrag O, kontakt K "&_
			" WHERE V.faktperiode  = " & periode &_
			" AND V.firmaid=f.firmaid " &_
			" AND O.bestilltav *= K.kontaktID " &_
			" AND O.oppdragID = V.oppdragid " &_
			" AND O.avdelingID IN (" & avdelingID & ") " &_
			" AND V.firmaID =" & firmaID &_
			" ORDER BY V.FirmaID, O.bestilltav  "

	
		Set rsKontakt = GetFirehoseRS(strSQL, Conn)
		if HasRows(rskontakt) = false Then
			%>
			<tr>
				<th colspan="2"><p class="warning">Det er ikke fakturert oppdrag i valgt periode</p></th>
			</tr>
			<%
			response.end
		end if
		dim personId
		dim kontakt
		dim rdoPersonID
		do until rsKontakt.EOF
			kontakt	= rsKontakt("kontakt")
			rdoPersonID = rsKontakt("SOPeID")
			if(isnull(kontakt) and not isnull(rdoPersonID)) then
				'kontakt = GetSOPersonName(rdoPersonID)
				kontakt = GetCRMContactPersonName(rdoPersonID)
				rdoPersonID = "SO_" & rdoPersonID
			else
				rdoPersonID = "XIS_" & rsKontakt("bestilltAv")
			end if
			%>
			<tr>
				<td><% =kontakt %></td>
				<td><INPUT class="radio" TYPE="radio" NAME="valg" VALUE="<%=rdoPersonID%>"</td>
			</tr>
			<%
			rsKontakt.MoveNext
		loop
		rsKontakt.close
		set rsKontakt = nothing		
		%>
		<tr>
			<td colspan="2"><input type="submit" value="Hent rapport"></td>
		</tr>			
		<%
	end sub

' her bygges HTML-siden opp, og sub'er kalles for å fylle inn innhold
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
		<body>
			<div class="pageContainer" id="pageContainer">
				<% 
					call hentKundeInfo 'header for basisinfo om kunde
				%>	
				<p>Periode: <%=" " & periode%></p>
				<form action="faktgrlagKunde.asp">
					<input type="hidden" name="periode" value="<%=periode%>">
					<input type="hidden" name="avdelinger" value="<%=avdelingID%>">
					<input type="hidden" name="firmaID" value="<%=firmaID %>">
					<div class="listing">
					<table cellpadding='0' cellspacing='1'>
						<tr>
							<th colspan="2">Kontaktpersoner</th>
					<%
						call hentKontakt_oppdrag()' innholdet i table hentes og tegnes opp i sub'en
					%>
					</table>
				</form>
			</div>
		</div>
	</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>