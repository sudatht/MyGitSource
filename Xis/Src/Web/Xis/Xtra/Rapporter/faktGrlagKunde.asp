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
	Dim strSQL 'variabel for sql-queries
	Dim avdelingID
	Dim periode 'fakturaperiode
	Dim rsRapport 'henter fakturagrunnlag
	Dim rsFirmaInfo 'henter basisopplysninger om kunde
	Dim kontaktID
	Dim kontaktNavn
	Dim ukedag
	Dim startTid
	Dim slutTid
	Dim sumTotal
	Dim addition
	Dim additionTotal
	Dim timerTotal
	Dim fakturaNr 'bruker dette da det sikrer at rapporten trekker ut kun linjer for den enkelte fakturering
	Dim gmlFakturaNr ' for å splitte rapport hvis ny fakturering har skjedd i perioden
	Dim FaktDato
	Dim GmlFaktDato
	Dim l_vikarID
	Dim Forrige_vikarID
	Dim loennsart
	Dim rsSum 'summerer beløp i footer
	Dim tmpBelop
	Dim tmpTimer
	Dim l_FaktDato
	dim blnIsXisContact 
	dim ForrigeVikarID
	dim label
	dim aXmlHTTP

	'initialiserer variable
	gmlfakturaNr = "start"

	' henter inputparametre
	firmaID = request("firmaID")
	avdelingID = request("avdelinger")
	periode = request("periode")
	kontaktID = request("valg")
	if(LEFT(kontaktID, 3) = "XIS") then
		blnIsXisContact = true
		kontaktID = mid(KontaktID, 5)
	else
		blnIsXisContact = false
		kontaktID = mid(KontaktID, 4)
	end if
	
	'oppretter databaseforbindelse
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	' sub for å hente kundeopplysninger
	sub HentKundeInfo()

		if(blnIsXisContact = true) then
			strSQL = " SELECT F.firma, F.SOCuID, A.adresse, A.postnr, A.poststed, kontaktperson = (K.Fornavn +  ' ' + K.Etternavn) "&_
				" FROM firma AS F, Kontakt AS K, adresse A  "&_
				" WHERE F.firmaid = " & firmaID &_
				" AND F.FirmaID *= K.FirmaID " &_
				" AND A.adresseRelID = f.firmaid " &_ 
				" AND K.KontaktID = " & kontaktID
		else
			strSQL = " SELECT F.firma, F.SOCuID " &_
				" FROM firma AS F "&_
				" WHERE F.firmaid = " & firmaID
				
			'kontaktNavn = GetSOPersonName(kontaktID)
			kontaktNavn = GetCRMContactPersonName(kontaktID)
		end if

		Set rsFirmaInfo = GetFirehoseRS(strSQL, Conn)
		if HasRows(rsFirmaInfo) Then
			firma = rsFirmaInfo("firma")
			if(blnIsXisContact = true) then
				kontaktNavn = rsFirmaInfo("kontaktperson")
				postadresse = rsFirmaInfo("adresse")
				poststed = rsFirmaInfo("postnr") & " " & rsFirmaInfo("poststed")				
			end if
			
			'Dim cts 
			'Set cts = server.CreateObject("Integration.SuperOffice")

			if(not isnull(rsFirmaInfo("SOCuID"))) then
				'dim rsAddress
				'set rsAddress = cts.GetAddressByContactId(clng(rsFirmaInfo("SOCuID")), 1)
				'if (hasRows(rsAddress)) then
				'	postadresse = rsAddress("address1")
				'	poststed = rsAddress("zipcode") & " " & rsAddress("city")
				'end if		
				'set rsAddress = nothing
				Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
				aXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetAccountAddress&Socuid=" + Cstr(rsFirmaInfo("SOCuID")) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
				aXmlHTTP.send ""
				postadresse = aXmlHTTP.responseText
			end if
			'set cts = nothing
			%>
			<h2>Firma: <%=" " & firma%></h2>
			<p>Adresse: <%=" " & postadresse & "  "&poststed%></p>
			<p>Kontaktperson: <%=" " & kontaktNavn %></p>
			<%
			rsFirmaInfo.Close
		end if
		Set rsFirmaInfo = nothing
	end sub


	'sub for å lage sideskift og skrive ut overskrift på nytt
	sub NySide()
		%>
		<h1 class="pageBreakBefore">Fakturagrunnlagsrapport</h1>
		<%
		call HentKundeInfo
		%>
		<p>Periode: <%=" " & periode%></p>
		<%
	End sub

	' sub for å skrive ut rapportheader
	sub RapportHeader()
		%>
		<div class="listing">
			<table>
				<tr>
					<th>Vikar</th>
					<th>Dag</th>
					<th>Dato</th>
					<th>Vikarnr.</th>
					<th>OppdNr.</th>
					<th>Tid</th>
					<th>F.tim</th>
					<th>Pris</th>
					<th>F.dato</th>
				</tr>
		<%
	end sub

	sub TotalHeader()
		%>
		<div class="listing">
			<table>
		<%

	end sub

	sub RapportAdditionsHeader()
		%>
		<tr><td>
		<div class="listing">
			<table>
				<tr>
					<th>Art.nr.</th>
					<th>Tillegg type</th>
					<th>Antall</th>
					<th>Beløp</th>
					<th>Sum</th>
				</tr>
		<%
	end sub

	' sub for å hente rapportlinjer
	sub HentRapport()
		dim kontaktSQL
		
		if(blnIsXisContact = true) then
			kontaktSQL = " AND o.bestilltav = " & kontaktID 
		else
			kontaktSQL = " AND o.SOPeID = " & kontaktID 
		end if
		
		strSQL = " SELECT konsulent=(v.Fornavn + '  ' + v.etternavn), d.vikarID, "&_
			" d.fakturadato, d.fakturanr, d.oppdragID, "&_
			" d.dato, d.faktperiode, d.starttid, d.sluttid, d.fakturatimer, "&_
			" d.fakturapris "&_
			" FROM dagsliste_vikar d, vikar v, oppdrag o "&_
			" WHERE faktperiode="& periode  &_
			" AND d.firmaID = " & firmaID  &_
			" AND d.vikarID = v.vikarID "&_
			" AND d.oppdragID = o.oppdragID "&_
			kontaktSQL &_
			" AND d.fakturatimer <> 0 "&_
			" ORDER BY d.fakturadato, d.vikarid, d.oppdragid, d.dato "

		set rsRapport = GetFirehoseRS(strSQL, Conn)
		l_VikarID = 0
		FaktDato = Null
		do until rsRapport.EOF
			FaktDato = rsRapport("fakturadato")
			l_VikarID = rsRapport("VikarID")

			if (not FaktDato = gmlFaktDato or l_vikarID <> ForrigeVikarID) AND not ForrigeVikarID = "" Then
				call rapportFooter(GmlFaktDato, forrigeVikarID)
				If not FaktDato = gmlFaktDato Then
					call NySide()
				End if
				call rapportHeader()
				sumTotal = 0
				timerTotal = 0
			end if

			ukedag = weekday(rsRapport("dato"), 2)
			SELECT	CASE ukedag
					CASE 1
						ukedag = "Man"
					CASE 2
						ukedag = "Tir"
					CASE 3
						ukedag = "Ons"
					CASE 4
						ukedag = "Tor"
					CASE 5
						ukedag = "Fre"
					CASE 6
						ukedag = "Lør"
					CASE 7
						ukedag = "Søn"
			END SELECT

			startTid = Left(TimeValue(rsRapport("Starttid")), 5)
			sluttid  = Left(TimeValue(rsRapport("sluttid")), 5)
			fakturaNr = rsRapport("fakturanr")
			FaktDato = rsRapport("fakturadato")
			l_VikarID = rsRapport("VikarID")
			%>
			<tr>
				<td><% =rsRapport("konsulent") %></td>
				<td><% =ukedag %></td>
				<td><% =rsRapport("dato") %></td>
				<td><% =rsRapport("vikarID") %></td>
				<td><% =rsRapport("oppdragID") %></td>
				<td><% =startTid & "-" & sluttid %></td>
				<td class="right"><% =formatNumber(rsRapport("fakturatimer"), 2) %></td>
				<td class="right"><% =rsRapport("fakturapris") %> kr</td>
				<td><% =rsRapport("fakturadato") %></td>
			</tr>
			<%
			gmlFakturaNr = fakturaNr
			ForrigeVikarID = l_vikarID
			GmlFaktDato = FaktDato
			
			rsRapport.MoveNext
		loop
		call RapportEnd()
		'call RapportAdditionSum(l_VikarID)
		Call TotalHeader()
		call RapportFooter(FaktDato, l_VikarID)
		call TotalEnd()
		rsRapport.close
		set rsRapport = nothing
	end sub

	' sub for rapportfooter
	sub RapportFooter(a_faktdato, a_VikarID)
		%>
		<tr>
			<td colspan="9" class="noTDBorder">&nbsp;</td>
		</tr>
		<%
		call RapportAdditionSum(a_VikarID)
		call BeregnOgVisSum(a_faktdato, a_VikarID)
		%>
		<tr>
			<td colspan="6"><strong>Sum tillegg :</strong></td>
			<td class="right"></strong></td>
			<td class="right"><%=formatNumber(additionTotal, 2)& " "%>kr</strong></td>
			<td>&nbsp;</td>
		</tr>		
		<tr>
			<td colspan="6"><strong>Sum Total:</strong></td>
			<td class="right"><strong><% =formatNumber(timerTotal, 2) & " "%>timer</strong></td>
			<td class="right"><strong><%=formatNumber(sumTotal, 2)& " "%>kr</strong></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
		</tr>
		<%
	end sub

	' sub for å beregne overtidslinjer
	sub BeregnOgVisSum(a_fakturadato, a_VikarID)
		dim kontaktSQL
		L_faktDato = dbDate(a_fakturadato)

		if(blnIsXisContact = true) then
			kontaktSQL = " AND O.bestilltav = " & kontaktID 
		else
			kontaktSQL = " AND O.SOPeID = " & kontaktID 
		end if

		strSQL =" SELECT belop=sum(fakturabeloep), timer=sum(fakturatimer),  isnull(groupkey,0) as FakturaProsent  "&_
			" FROM vikar_ukeliste v " &_
			" INNER JOIN Oppdrag AS O ON V.OppdragID = O.OppdragID " &_
			" WHERE v.fakturadato = "& l_faktdato &_
			" AND v.fakturaSats >=  1.0"  &_
			" AND v.FirmaID = " & firmaID &_
			kontaktSQL &_
			" AND v.vikarID = "& a_VikarID &_
			" GROUP BY v.fakturadato, v.groupkey"

		set rsSum = GetFirehoseRS(strSQL, Conn)

		if hasRows(rsSum) Then
			while not rsSum.EOF
				
				if(rsSum("FakturaProsent") = "step1") then
					label = "Overtid 1 "
				else 
					if (rsSum("FakturaProsent") = "step2") then
						label = "Overtid 2 "
					else

						label = "Sum vanlige timer "
					end if
				end if

				tmpBelop = rsSum("belop")
				tmpTimer = rsSum("timer")
				%>
				<tr>
					<td colspan="6"><strong><%=label%>:</strong></td>
					<td class="right"><% =formatNumber(tmpTimer, 2) & " " %>timer</td>
					<td class="right"><%=formatNumber(tmpBelop, 2) & " "%>kr</td>
					<td>&nbsp;</td>
				</tr>				
				<%
				timerTotal = TimerTotal + tmpTimer
				sumTotal = sumTotal + tmpBelop				
				rsSum.MoveNext
			wend		
			rsSum.close
		end if
		set rsSum = nothing		
	end sub
	
	sub RapportEnd()
		%>
			</table>
		</div>
		<%
	end sub

	sub TotalEnd()
		%>
			</table>
		</div>
		<%
	end sub

	sub RapportAdditionEnd()
		%>
			</table></td></tr>
		</div>
		<%
	end sub

	sub RapportAdditionSum(a_VikarID)

		strSQL =" SELECT additions=sum(InvTotal),A.ArticleID,H.description,InvRate,InvUnits  "&_
			" FROM Addition A " &_
			" INNER JOIN Oppdrag AS O ON A.OppdragID = O.OppdragID " &_
			" INNER JOIN H_ADDITIONS_ARTICLES AS H ON A.ArticleID = H.ArticleID " &_
			" WHERE A.Invperiod = "& periode &_
			" AND o.firmaid = " & firmaID &_
			" AND A.vikarID = "& a_VikarID &_
			" GROUP BY H.description,InvRate,InvUnits,A.ArticleID "


		set rsSum = GetFirehoseRS(strSQL, Conn)
		if hasRows(rsSum) Then
			call RapportAdditionsHeader()
			additionTotal = 0
			while not rsSum.EOF
				addition = rsSum("additions")
				
				%>
				<tr>
					<td><% =rsSum("ArticleID") %></td>
					<td><% =rsSum("description") %></td>
					<td><% =formatNumber(rsSum("InvUnits"),2) %></td>
					<td><% =formatNumber(rsSum("InvRate"),2) %></td>
					<td class="right"><% =formatNumber(rsSum("additions"),2) %> kr</td>
				</tr>

				<%
				
				rsSum.MoveNext
				sumTotal = sumTotal + addition
				additionTotal = additionTotal + addition
			wend	
			rsSum.close
			call RapportAdditionEnd()
		end if
		set rsSum = nothing

	end sub
	
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title>Fakturagrunnlagsrapport</title>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Fakturagrunnlagsrapport</h1>
			</div>
			<div class="content">
				<%
					call hentKundeInfo() 'header for basisinfo om kunde
				%>
				<p>Periode: <%=" " & periode%></p>
				<%
					call RapportHeader()
					call HentRapport() 'henter rapportlinjer
					
				%>
    		</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing

%>