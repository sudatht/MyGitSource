<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Economics.Constants.inc"-->
<%

	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim foundRecords : foundRecords = false
	dim strSQL
	dim rsRapport
	Sub beregnOvertid(fraUkenr, tilUkenr, vikarID, oppdragID, typeID, Db, oms, ln)

		dim strSQL
		dim rsOvertid
		dim overtidLonnProsent
		dim overtidFakturaProsent
		dim Fbelop
		dim Lbelop
		dim LTimer
		dim FTimer
		dim DekningsbidragOvertid

		strSQL = " SELECT " &_
			" loennSats = (l.Loennrate * 100)  - 100, " &_
			" fakturaSats = (F.fakturaSats * 100)  - 100," &_
			" FTimer=SUM(V.Fakturatimer)," &_
			" FBelop=(SUM(V.Fakturabeloep)/(F.fakturaSats*100))*((F.fakturaSats*100)-100), " &_
			" LBelop=SUM(V.Belop), " &_
			" LBelop=(SUM(V.Belop)/(L.Loennrate*100))*((L.Loennrate*100)-100), " &_
			" LTimer=SUM(V.antall) " &_
			" FROM VIKAR_UKELISTE V, h_loennsArt L, H_faktura_type F "& _
			" WHERE " &_
			" V.Loennsartnr = L.Loennsartnr " &_
			" AND V.fakturatype = F.fakturatype " &_
			" AND L.LoennRate > 1.0 " &_			
			" AND V.Fakturapris > 0 "& _
			" AND V.ukenr <= "& tilUkenr & _
			" AND V.ukenr >= "& fraUkenr & _
			" AND V.vikarID = "& vikarID & _
			" AND V.oppdragid = "& oppdragID & _
 			" AND NOT V.ID IN(SELECT ID FROM vikar_ukeliste WHERE (V.ukenr= " & tilUkenr & " AND notat like '2') "& _
			"		  OR (V.ukenr = " & fraUkenr & " AND notat LIKE '1')) " & _
			" GROUP BY L.Loennrate, F.fakturaSats"

		set rsOvertid = GetFirehoseRS(strSQL, Conn)

		DO WHILE (not rsOvertid.EOF)

			overtidLonnProsent 	= rsOvertid("loennSats")
			overtidFakturaProsent = rsOvertid("fakturaSats")
			
			Fbelop 			= rsOvertid("FBelop")
			Lbelop 			= rsOvertid("LBelop")
			LTimer 			= rsOvertid("LTimer")
			FTimer 			= rsOvertid("FTimer")

			If (typeID = 1) Then
				DekningsbidragOvertid = Fbelop - ( Lbelop * XIS_FACTOR )
			Else
				DekningsbidragOvertid = Fbelop -  Lbelop
			End If

			Dekningsbidrag = (Db + DekningsbidragOvertid)
			Omsetning = (oms + Fbelop)
			loenn = (ln + Lbelop)

			Response.Write "<tr><td></td><td colspan=4>Overtid Lønn "& overtidLonnProsent & "%, Faktura " & overtidFakturaProsent & "%</td><td class=right>"& formatNumber(FTimer,2) &"</td><td class=right>"& formatNumber(Fbelop,2) &"</td><td class=right>"& formatNumber(LTimer,2) &"</td><td class=right>"& formatNumber(Lbelop,0)&"</td><td></td><td class=right>"& formatNumber(Fbelop,0)&"</td><td class=right>"& formatNumber(DekningsbidragOvertid,0)&"</td></tr>"
			rsOvertid.MoveNext
		loop
		rsOvertid.close
		set rsOvertid = nothing
	End Sub

	Sub TotaltFooter( Omsetning, Bidrag, AntallOppdrag, Loenn, AntTimer, FaktTimer )

		dim DiffTimer

		if ((omsetning <> 0) AND (bidrag <> 0) AND (omsetning <> bidrag)) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if

		DiffTimer = ( AntTimer - FaktTimer )

		Response.Write "<tr><TD COLSPAN=2><strong>Sum totalt</strong></td>"
		Response.Write "<TD COLSPAN=6></td>"
		Response.Write "<TD COLSPAN=3><strong>Total omsetning:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Omsetning, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total bidrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Bidrag, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total lønn:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( loenn, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Antall oppdrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & AntallOppdrag & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Faktor:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber(Faktor, 2 ) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber(DiffTimer, 0) & "</strong></td>"

	End Sub

	Sub AvdelingHeader( rsRapport )
		' Create heading on avdeling
		Response.Write "<tr>"
		Response.Write "<td colspan='3'><h4>Avdeling: " & rsRapport("Avdeling") & "</h4></td>"
		Response.Write "<td colspan='9'>&#160;</td>"
		Response.Write "</tr>"

	End Sub

	Sub AvdelingFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )
		if (omsetning <> 0 AND bidrag <> 0 AND omsetning <> bidrag) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if

		DiffTimer = AntTimer - FaktTimer

		Response.Write "<tr><TD COLSPAN=2><strong>Sum avdeling</strong></td>"
		Response.Write "<TD COLSPAN=6></td>"
		Response.Write "<TD COLSPAN=3><strong>Total omsetning:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Omsetning, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total bidrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Bidrag, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total lønn:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( loenn, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Antall oppdrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & AntallOppdrag & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Faktor:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Faktor, 2 ) & "&nbsp;</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( DiffTimer, 0) & "</strong></td>"
	End Sub

	Sub MedarbeiderHeader( rsRapport  )
		' Create table heading
		Response.Write "<tr>"
		Response.Write "<th COLSPAN='3'><h5>Ansvarlig " & rsRapport("Medarbeider") & "</h5></th>"
		Response.Write "<td COLSPAN='9'>&#160;</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<th>Opp.nr.</th>"
		Response.Write "<th>Kunde</th>"
		Response.Write "<th>Vikar (Ansattnr.)</th>"
		Response.Write "<th>Startdato</th>"
		Response.Write "<th>Sluttdato</th>"
		Response.Write "<th>F.timer</th>"
		Response.Write "<th>Pris</th>"
		Response.Write "<th>L.timer</th>"
		Response.Write "<th>Lønn</th>"
		Response.Write "<th>Faktor</th>"
		Response.Write "<th>Oms.</th>"
		Response.Write "<th>DB</th>"
		Response.Write "</tr>"
	End Sub

	Sub MedarbeiderFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )
		dim DiffTimer

		if (omsetning <> 0 AND bidrag <> 0 AND omsetning<>bidrag) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if

		DiffTimer = (AntTimer - FaktTimer)

		Response.Write "<tr><TD COLSPAN=2><strong>Sum ansvarlig</strong></td>"
		Response.Write "<TD COLSPAN=6></td>"
		Response.Write "<TD COLSPAN=3><strong>Total omsetning:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Omsetning, 0 ) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total bidrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Bidrag, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Total lønn:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( loenn, 0) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Antall oppdrag:</strong></td>"
		Response.Write "<TD class=right><strong>" & AntallOppdrag & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Faktor:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber( Faktor, 2 ) & "</strong></td>"

		Response.Write "<tr><TD COLSPAN=8></td>"
		Response.Write "<TD COLSPAN=3><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<TD class=right><strong>" & FormatNumber(DiffTimer, 0) & "</strong></td>"
	End Sub

	Sub OppdragFooter( OppdragID, Firma, Vikar, VikarID,  Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn, LoennsTimer, TypeID, faktTimer, ansattnummer)
		'Create row
		Response.Write "<tr>"
		Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & OppdragID, OppdragID, "Vis oppdrag "  ) & "&nbsp;</td>"
		Response.Write "<td>" & Firma & "</td>"
		Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & VikarID, Vikar, "Vis vikar " & Vikar )  & " (" & ansattnummer & ")</td>"
		Response.Write "<td>" & Fradato & "</td>"
		Response.Write "<td>" & Tildato & "</td>"
		Response.Write "<TD class=right>"& FormatNumber(faktTimer, 2) & "</td>"
		If Fakturapris <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Fakturapris, 0 )  & "</td>"
		Else
			Response.Write "<TD class=right>0</td>"
		End If
		Response.Write "<TD class=right>"& FormatNumber(LoennsTimer, 2) &"</td>"
		If Timelonn <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Timelonn, 0 )  & "</td>"
		Else
			Response.Write "<TD class=right>0</td>"
		End If
		If  ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR ) <> 0 Then
			Faktor = Omsetning  / ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR )
			Response.Write "<TD ALIGN=center>" & FormatNumber( Faktor , 2 )  & "</td>"
		Else
			Response.Write "<TD class=right>&nbsp;</td>"
		End If

		If Omsetning <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Omsetning, 0 )  & "</td>"
		Else
			Response.Write "<TD class=right>0</td>"
		End If

		If Dekningsbidrag <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Dekningsbidrag, 0 ) & "</td>"
		Else
			Response.Write "<TD class=right>0</td>"
		End If

		Response.Write "</tr>"
	End Sub

		' Is this first time to show this page
		If Request.Form( "tbxPageNo") <> "" Then

			' Add values from current page
			Fradato     = Request.Form( "tbxFradato" )
			Tildato     = Request.Form( "tbxTildato" )

			if (Fradato <> "" AND Tildato <> "") then
				if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
					AddErrorMessage("Fradato kan ikke være senere enn tildato!")
					call RenderErrorMessage()
				end if
			end if

			if Fradato <> "" then
				FraDatoUke = Fradato
				FraDatoAar = Fradato
				call KorrigerUke(FraDatoUke)
				call KorrigerAar(FraDatoAar)
				FraDatoAarsuke = FraDatoAar & FraDatoUke
			end if

			if Tildato <> "" then
				TilDatoUke = Tildato
				TilDatoAar = Tildato
				call KorrigerUke(TilDatoUke)
				call KorrigerAar(TilDatoAar)
				TilDatoAarsuke=TilDatoAar&TilDatoUke
			end if

			SelectAvdelingID = Request.form("dbxAvdeling")

		End If

		' Open database connection
		Set Conn = GetConnection(GetConnectionstring(XIS, ""))

		' First time page called AND search value exist ?
		If ((Fradato <> "") AND (Tildato <> ""))  Then

			If SELECTAvdelingID > 0 Then
				strSelectAvdeling = " AND O.AvdelingID = " & SELECTAvdelingID
			End If

			strSQL = "SELECT A.AvdelingID, M.MedID, DV.OppdragID, DV.OppdragVikarID, DV.VikarID, A.Avdeling, Medarbeider=M.Etternavn+' '+M.Fornavn, F.Firma, Vikar=V.Etternavn, V.TypeID, " &_
				" O.Fradato, O.Tildato, DV.Fakturapris, DV.Fakturatimer, DV.AntTimer, DV.Timelonn, DV.Dato, VIKAR_ANSATTNUMMER.ansattnummer " &_
				"FROM DAGSLISTE_VIKAR DV, OPPDRAG O, FIRMA F, VIKAR V LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON V.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, AVDELING A, MEDARBEIDER M " &_
				"WHERE DV.Dato >= " & DbDate( fradato) &_
				" AND DV.Dato <= " & DbDate( Tildato) &_
				" AND DV.Anttimer > 0 " &_
				" AND DV.OppdragID = O.OppdragID " &_
				" AND DV.OppdragID > 0 " &_
				strSelectAvdeling &_
				" AND O.AvdelingID = A.AvdelingID " &_
				" AND O.AnsMedID = M.MedID " &_
				" AND DV.VikarID = V.VikarID " &_
				" AND DV.FirmaID = F.FirmaID " &_
				"ORDER BY A.AvdelingID, M.MedID, F.Firma, DV.OppdragID, DV.VikarID, Dv.Fakturapris, DV.Timelonn"

			set rsRapport = GetFirehoseRS(strSQL, Conn)

			'Records found ?
			If (HasRows(rsRapport) = true)  Then
				FoundRecords = true
			End If
		End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script language="javascript" src="../Js/javascript.js" type="text/javascript"></script>
	<title>Omsetning pr. avdeling/Ansvarlig/Oppdrag</title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>Omsetning pr. avdeling/Ansvarlig/Oppdrag</h1>
		</div>
		<div class="content">
			<p>gir kun riktig overtid hvis datointervallet er en hel måned</p>
			<form name="formEn" ACTION="omsetningAvdelingTimer.asp" METHOD="POST">
				<input type="hidden" NAME="tbxPageNo" value="1">
				<table>
					<tr>
						<td>Fra dato:</td>
						<td><INPUT NAME="tbxFraDato" TYPE=TEXT SIZE=10 MAXLENGTH=10 Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)"> </td>
						<td>Til dato:</td>
						<td><INPUT NAME="tbxTilDato" TYPE=TEXT SIZE=10 MAXLENGTH=10 Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)"> </td>
						<td>Avdeling:</td>
						<td>
							<select NAME="dbxAvdeling">
								<OPTION VALUE="0">
								<%
								' Get avdeling
								strSQL = "SELECT AvdelingID, Avdeling FROM Avdeling ORDER BY avdeling"
								set rsAvdeling = GetFirehoseRS(strSQL, Conn)
								Do Until (rsAvdeling.EOF)
									If (CInt(rsAvdeling("AvdelingID")) = CInt(Request("dbxAvdeling"))) Then
										sel = " SELECTED"
									Else
										sel = ""
									end if
									%>
									<option value="<% =rsAvdeling("AvdelingID") %>" <% =sel %>><% =rsAvdeling("Avdeling") %></option>
									<%  	
									rsAvdeling.MoveNext
								Loop
								' Close AND release recordset
								rsAvdeling.Close
								Set rsAvdeling = Nothing
								%>
							</select>
						</td>
						<td><input type="submit" name="pbnDataAction" value="     Søk    "></td>
					</tr>
				</table>
			</form>
			<div class="listing">
				<%
				' Create table only when records found
				If  (foundRecords)  Then
					' Create table
					Response.Write "<table>"
					Do Until (rsRapport.EOF)
						TypeID = rsrapport("TypeID")
						' Break on Avdeling ?
						If rsRapport( "AvdelingID") <> AvdelingID Then
							' Do we have a Oppdrag ?
							If OppdragID <> "" or VikarID <> "" Then
								' Create avdeling heading
								Call OppdragFooter( OppdragID, Firma, Vikar, VikarID, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn, AntTimerOppdrag, TypeID, AntFTimerOppdrag, strAnsattnummer )
								call beregnOvertid(FraDatoAarsuke, TilDatoAarsuke, vikarID, oppdragID, TypeID, Dekningsbidrag, Omsetning, loenn)

								AntTimerOppdrag = 0
								AntFTimerOppdrag = 0

								' Accumulate va
								OmsMedarbeider = OmsMedarbeider + Omsetning
								BidragMedarbeider = BidragMedarbeider + Dekningsbidrag
								LoennMedarbeider = LoennMedarbeider + Loenn

								AntTimerMedarbeider = AntTimerMedarbeider + AntTimer
								FaktTimerMedarbeider = FaktTimerMedarbeider + FaktTimer

								' Reset values
								OppdragID = ""
								VikarID = ""
								Omsetning = 0
								Loenn = 0
								Dekningsbidrag = 0

								AntTimer = 0
								FaktTimer = 0
							End If
							' Do we have a Medarbeider ?
							If (MedID <> "") Then
								' Create footer
								Call MedarbeiderFooter( OmsMedarbeider, BidragMedarbeider, AntallOppdrag, LoennMedarbeider , AntTimerMedarbeider, FaktTimerMedarbeider  )
								' Reset values
								MedID = ""
								OmsAvdeling = OmsAvdeling + OmsMedarbeider
								BidragAvdeling = BidragAvdeling + BidragMedarbeider
								AntallAvdeling = AntallAvdeling + AntallOppdrag
								LoennAvdeling = LoennAvdeling + LoennMedarbeider
								AntTimerAvdeling = AntTimerAvdeling + AntTimerMedarbeider
								FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerMedarbeider
								AntallOppdrag = 0
								OmsMedarbeider = 0
								BidragMedarbeider = 0
								LoennMedarbeider = 0

								AntTimerMedarbeider = 0
								FaktTimerMedarbeider = 0
							End If
							' Do we have a Avdeling ?
							If AvdelingID <> "" Then
								' Create footer
								Call AvdelingFooter( OmsAvdeling, BidragAvdeling, AntallAvdeling, LoennAvdeling, AntTimerAvdeling, FaktTimerAvdeling  )

								AntallTotalt = AntallTotalt + AntallAvdeling
								OmsTotalt = OmsTotalt + OmsAvdeling
								BidragTotalt = BidragTotalt + BidragAvdeling
								LoennTotalt = LoennTotalt + LoennAvdeling

								AntTimerTotalt = AntTimerTotalt + AntTimerAvdeling
								FaktTimerTotalt = FaktTimerTotalt + FaktTimerAvdeling

								OmsAvdeling = 0
								BidragAvdeling = 0
								AntallAvdeling = 0
								LoennAvdeling = 0

								AntTimerAvdeling = 0
								FaktTimerAvdeling = 0
							End If
							' Create avdeling heading
							Call AvdelingHeader( rsRapport  )
						End If
						' Break on Ansvarlig medarbeider ?
						If rsRapport( "MedID") <> MedID Then
							' Do we have a Oppdrag ?
							If OppdragId <> "" Or VikarID <> "" Then
								' Create avdeling heading
								Call OppdragFooter( OppdragID, Firma, Vikar, VikarID, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn, AntTimerOppdrag, TypeID, AntFTimerOppdrag, strAnsattnummer )
								call beregnOvertid(FraDatoAarsuke, TilDatoAarsuke, vikarID, oppdragID, TypeID, Dekningsbidrag, Omsetning, loenn)

								AntTimerOppdrag = 0
								AntFTimerOppdrag= 0
								OmsMedarbeider = OmsMedarbeider + Omsetning
								BidragMedarbeider = BidragMedarbeider + Dekningsbidrag
								LoennMedarbeider = LoennMedarbeider + Loenn
								AntTimerMedarbeider = AntTimerMedarbeider + AntTimer
								FaktTimerMedarbeider = FaktTimerMedarbeider + FaktTimer

								' Set new value
								Omsetning = 0
								Dekningsbidrag = 0
								Loenn = 0
								AntTimer = 0
								FaktTimer = 0
								OppdragID = ""
								VikarID = ""
							End If
							' Do we have a Medarbeider ?
							If MedID <> "" Then
								' Create footer
								Call MedarbeiderFooter( OmsMedarbeider, BidragMedarbeider, AntallOppdrag, LoennMedarbeider , AntTimerMedarbeider, FaktTimerMedarbeider  )
								OmsAvdeling = OmsAvdeling + OmsMedarbeider
								BidragAvdeling = BidragAvdeling + BidragMedarbeider
								AntallAvdeling = AntallAvdeling + AntallOppdrag
								LoennAvdeling = LoennAvdeling + LoennMedarbeider

								AntTimerAvdeling = AntTimerAvdeling + AntTimerMedarbeider
								FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerMedarbeider

								OmsMedarbeider = 0
								BidragMedarbeider = 0
								LoennMedarbeider = 0

								AntTimerMedarbeider = 0
								FaktTimerMedarbeider = 0
								' Reset values
								AntallOppdrag = 0
							End If
							' Create header
							Call MedarbeiderHeader( rsRapport )
						End If
						' Break on oppdragid
						If rsRapport("OppdragID") <> OppdragID Or rsRapport("VikarID") <> VikarID or rsRapport("FakturaPris") <> FakturaPris or rsRapport("Timelonn") <> Timelonn Then
							' Do we have a Oppdrag ?
							If OppdragID <> "" Or VikarID <> "" Then
								' Create avdeling heading
								Call OppdragFooter( OppdragID, Firma, Vikar, VikarID, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn, AntTimerOppdrag, TypeID, AntFTimerOppdrag, strAnsattnummer )
								AntTimerOppdrag = 0
								AntFTimerOppdrag = 0
								'Skriv ut overtid bare dersom det er et nytt oppdrag (ikke ved to ulike priser)
								If rsRapport("OppdragID") <> OppdragID Or rsRapport("VikarID") <> VikarID then
									call beregnOvertid(FraDatoAarsuke, TilDatoAarsuke, vikarID, oppdragID, TypeID, Dekningsbidrag, Omsetning, loenn)
								End if
							End If

							OmsMedarbeider = OmsMedarbeider + Omsetning
							BidragMedarbeider = BidragMedarbeider + Dekningsbidrag
							LoennMedarbeider = LoennMedarbeider + Loenn

							AntTimerMedarbeider = AntTimerMedarbeider + AntTimer
							FaktTimerMedarbeider = FaktTimerMedarbeider + FaktTimer

							Omsetning = 0
							Dekningsbidrag = 0
							Loenn = 0
							AntTimer = 0
							FaktTimer = 0
							' Set new value
							strAnsattnummer = rsRapport("ansattnummer").Value
							OppdragID = rsRapport("OppdragID")
							VikarID = rsRapport("VikarID")
							Firma = rsRapport( "Firma")
							Vikar = rsRapport( "Vikar")
							Fradato = rsRapport( "Fradato")
							Tildato = rsRapport( "TilDato")
							' tell antall oppdrag
							AntallOppdrag = AntallOppdrag + 1
						End If
						Omsetning = Omsetning + ( rsRapport("FakturaTimer") *  rsRapport("Fakturapris") )
						Loenn = Loenn + ( rsRapport("AntTimer") * rsRapport("Timelonn") )
						If rsRapport("TypeID") = 1 Then
							Dekningsbidrag = Dekningsbidrag + ( ( rsRapport("Fakturapris") * rsRapport("FakturaTimer") ) - ( rsRapport("Timelonn") * rsRapport("AntTimer") * XIS_FACTOR ) )
						Else
							Dekningsbidrag = Dekningsbidrag + ( ( rsRapport("Fakturapris") * rsRapport("FakturaTimer") ) - ( rsRapport("Timelonn") * rsRapport("AntTimer") ) )
						End If
						AntTimer  = rsRapport( "Anttimer")
						FaktTimer = rsRapport( "Fakturatimer")

						' Set new value
						strAnsattnummer = rsRapport("ansattnummer").Value
						AvdelingID = rsRapport("AvdelingID")
						MedID = rsRapport("MedID")
						VikarID = rsRapport("VikarID")
						OppdragID = rsRapport("OppdragID")

						' This will correct for each record
						Fakturapris = rsRapport("Fakturapris")
						Timelonn = rsRapport("Timelonn")
						AntTimerOppdrag = AntTimerOppdrag + rsRapport("antTimer")
						AntFTimerOppdrag= AntFTimerOppdrag + rsRapport("Fakturatimer")

						' Get next record
						rsRapport.MoveNext
					Loop
					' Do we have a Oppdrag ?
					If VikarID <> "" Then
						' Create avdeling heading
						Call OppdragFooter(  OppdragID, Firma, Vikar, VikarID, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn, AntTimerOppdrag, TypeID, AntFTimerOppdrag, strAnsattnummer )
						call beregnOvertid(FraDatoAarsuke, TilDatoAarsuke, vikarID, oppdragID, TypeID, Dekningsbidrag, Omsetning, loenn)
							AntTimerOppdrag = 0
							AntFTimerOppdrag= 0
						OmsMedarbeider = OmsMedarbeider + Omsetning
						BidragMedarbeider = BidragMedarbeider + Dekningsbidrag
						LoennMedarbeider = LoennMedarbeider + Loenn
						AntTimerMedarbeider = AntTimerMedarbeider + AntTimer
						FaktTimerMedarbeider = FaktTimerMedarbeider + FaktTimer
						AntallOppdrag = AntallOppdrag + 1
					End If
					' Do we have a Medarbeider ?
					If MedID <> "" Then
						' Create footer
						Call MedarbeiderFooter( OmsMedarbeider, BidragMedarbeider, AntallOppdrag, LoennMedarbeider , AntTimerMedarbeider, FaktTimerMedarbeider )
						OmsAvdeling = OmsAvdeling + OmsMedarbeider
						BidragAvdeling = BidragAvdeling + BidragMedarbeider
						AntallAvdeling = AntallAvdeling + AntallOppdrag
						LoennAvdeling = LoennAvdeling + LoennMedarbeider
						AntTimerAvdeling = AntTimerAvdeling + AntTimerMedarbeider
						FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerMedarbeider
					End If

					' Do we have a Avdeling ?
					If AvdelingID <> "" Then

						' Create footer
						Call AvdelingFooter( OmsAvdeling, BidragAvdeling, AntallAvdeling, LoennAvdeling, AntTimerAvdeling, FaktTimerAvdeling )

						AntallTotalt = AntallTotalt + AntallAvdeling
						OmsTotalt = OmsTotalt + OmsAvdeling
						BidragTotalt = BidragTotalt + BidragAvdeling
						LoennTotalt = LoennTotalt + LoennAvdeling

						AntTimerTotalt = AntTimerTotalt + AntTimerAvdeling
						FaktTimerTotalt = FaktTimerTotalt + FaktTimerAvdeling

					End If
					' Create footer
					Call TotaltFooter( OmsTotalt, BidragTotalt, AntallTotalt, LoennTotalt, AntTimerTotalt, FaktTimerTotalt )
					' Close recordset
					rsRapport.Close
					' Clear recordset
					set rsRapport = Nothing
					' End table
					Response.Write "</table>"
				End If
				%>
					<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
				</div>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>