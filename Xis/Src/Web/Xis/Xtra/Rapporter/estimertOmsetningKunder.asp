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
	
	Response.Expires = 0	
	dim foundRecords : foundRecords = false
	
	Sub TotaltFooter( Omsetning, Bidrag, AntallOppdrag, Loenn, AntTimer, FaktTimer )
		if (omsetning <> 0 AND bidrag <> 0 AND omsetning <> bidrag) then
				Faktor = Omsetning / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if
		DiffTimer = ( AntTimer - FaktTimer )

		Response.Write "<tr>"
		Response.Write "<td colspan='9'><hr></td>"
		Response.Write "<tr><td colspan='2'><strong>Sum totalt</strong></td>"
		Response.Write "<td colspan='3'></td>"
		Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber(Faktor, 2 ) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber(DiffTimer, 0) & "</strong></td></tr>"
	End Sub

	Sub AvdelingHeader( rsRapport  )
		' Create heading on avdeling
		Response.Write "<tr>"
		Response.Write "<td colspan='3'><h2>Avdeling: " & rsRapport("Avdeling") & "</h2></td>"
		Response.Write "</tr>"
	End Sub

	Sub AvdelingFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )

		If  ( ( Omsetning - Bidrag ) / XIS_FACTOR ) <> 0 Then
			Faktor = Omsetning  / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
		Else
			Faktor = 0
		End If
		
		'Faktor =  Omsetning /  ( (Omsetning - Bidrag ) / XIS_FACTOR  )
		DiffTimer = AntTimer - FaktTimer

		Response.Write "<tr>"
		Response.Write "<td colspan='9'><hr></td>"
		Response.Write "<tr><td colspan='2'><strong>Sum avdeling</strong></td>"
		Response.Write "<td colspan='3'></td>"
		Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Faktor, 2 ) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( DiffTimer, 0) & "</strong></td></tr>"
	End Sub

	Sub FirmaHeader( rsRapport  )
		' Create table heading
		Response.Write "<tr>"
		Response.Write "<td colspan='3'><h2>Kunde: " & rsRapport("Firma") & "</h2></td>"
		Response.Write "</tr>"
		Response.Write "<td colspan='9'><hr></td>"
		Response.Write "</tr>"
		Response.Write "<tr><th>Opp.nr.</th>"
		Response.Write "<th>Ansvarlig</th>"
		Response.Write "<th>Vikar</th>"
		Response.Write "<th>Timer</th>"
		Response.Write "<th>Pris</th>"
		Response.Write "<th>Lønn</th>"
		Response.Write "<th>Faktor</th>"
		Response.Write "<th>Oms.</th>"
		Response.Write "<th>DB</th></tr>"
	End Sub

	Sub FirmaFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer)
		If  ( ( Omsetning - Bidrag ) / XIS_FACTOR ) <> 0 Then
			Faktor = Omsetning  / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
		Else
			Faktor = 0
		End If

		Response.Write "<tr>"
		Response.Write "<td colspan='9'><hr></td>"
		Response.Write "<tr><td colspan='2'><strong>Sum kunde</strong></td>"
		Response.Write "<td colspan='3'></td>"
		Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0 ) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
		Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( Faktor, 2 ) & "</strong></td></tr>"

		Response.Write "<tr><td colspan='5'></td>"
		Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
		Response.Write "<td class='right'><strong>" & FormatNumber( DiffTimer, 0) & "</strong></td></tr>"
	End Sub

	Sub OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )

            Response.Write "<tr>"
            Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & OppdragID, OppdragID , "Vis Oppdrag"  ) & "</td>"
            Response.Write "<td>" & Medarbeider & "</td>"
            Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & vikarID, Vikar, "Vis vikar " & Vikar ) & "</td>"

			if Omsetning <> 0 OR Fakturapris <> 0 then
				Response.Write "<td>"& FormatNumber(Omsetning/Fakturapris, 1) & "</td>"
			else
				Response.Write "<td></td>"
			end If
			
            If Fakturapris <> 0 Then
               Response.Write "<td class='right'>" & FormatNumber( Fakturapris, 0 )  & "</td>"
            Else
               Response.Write "<td class='right'>" & "0" & "</td>"
            End If
            
            If Timelonn <> 0 Then
               Response.Write "<td class='right'>" & FormatNumber( Timelonn, 0 )  & "</td>"
            Else
               Response.Write "<td class='right'>" & "0" & "</td>"
            End If
            
            If  ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR ) <> 0 Then
               Faktor = Omsetning  / ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR )
               Response.Write "<td class='right'>" & FormatNumber( Faktor , 2 )  & "</td>"
            Else
               Response.Write "<td class='right'>" & "</td>"
            End If
            
            If Omsetning <> 0 Then
               Response.Write "<td class='right'>" & FormatNumber( Omsetning, 0 )  & "</td>"
            Else
               Response.Write "<td class='right'>" & "0" & "</td>"
            End If

            If Dekningsbidrag <> 0 Then
               Response.Write "<td class='right'>" & FormatNumber( Dekningsbidrag, 0 ) & "</td>"
            Else
               Response.Write "<td class='right'>" & "0" & "</td>"
            End If
            Response.Write "</tr>"
End Sub

	dim strSQL
	dim Conn

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

' Is this first time to show this page
If Request.Form( "tbxPageNo") <> "" Then
   ' Add values FROM current page
   Fradato = Request.Form( "tbxFradato" )
   Tildato = Request.Form( "tbxTildato" )

   SelectAvdelingID = Request.form("dbxAvdeling")
   
	' First time page called AND search value exist ?
	If Fradato <> "" AND Tildato <> ""  Then
		if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
			Response.write "<p class'warning'>Fradato kan ikke være senere enn tildato!</p>"
			Response.End
		end if

		If SelectAvdelingID > 0 Then
			strSelectAvdeling = " AND O.AvdelingID = " & SelectAvdelingID
		End If
		
		' Get all
		strSQL = "SELECT A.AvdelingID, M.MedID, DV.OppdragID, DV.OppdragVikarID, DV.VikarID, A.Avdeling, Medarbeider=M.Etternavn+' '+M.Fornavn, F.FirmaID, F.Firma, Vikar = V.fornavn + ' ' + V.Etternavn, V.VikarID, V.TypeID, " &_
				" O.Fradato, O.Tildato, DV.FakturaPris, DV.FakturaTimer, DV.AntTimer, DV.Timelonn " &_
				"FROM DAGSLISTE_VIKAR DV, OPPDRAG O, FIRMA F, VIKAR V, AVDELING A, MEDARBEIDER M " &_
				"WHERE DV.Dato >= " & DbDate( fradato) &_
				" AND DV.Dato <= " & DbDate( Tildato) &_
				" AND DV.OppdragID = O.OppdragID " &_
				" AND DV.OppdragID > 0 " &_
				" AND DV.Anttimer > 0 " &_
					strSelectAvdeling &_
				" AND O.AvdelingID = A.AvdelingID " &_
				" AND O.AnsMedID = M.MedID " &_
				" AND DV.VikarID = V.VikarID " &_
				" AND DV.FirmaID = F.FirmaID " &_
				" ORDER BY A.AvdelingID, F.Firma, DV.FirmaID, DV.OppdragID, DV.VikarID, Dv.Fakturapris, DV.Timelonn"

		Set rsRapport = GetFirehoseRS(strSQL, Conn)

		' No records found ?
		If (not rsRapport.EOF) Then
			foundRecords = true
		End If
	End If   
End If

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
		<script language="javaScript" src="../Js/javaScript.js" type="text/javascript"></script>
		<title>Omsetning pr. avdeling</title>
	</head>
	<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>Estimert omsetning pr. kunder/ansvarlig/oppdrag</h1>
		</div>
		<div class="content">
			<form name="formEn" ACTION="EstimertOmsetningKunder.asp" METHOD="POST">
				<input type="hidden" NAME="tbxPageNo" VALUE="1">
				<table>
					<tr>
						<td>Fra dato:</td>
						<td><INPUT NAME="tbxFraDato" TYPE=TEXT  SIZE=10 MAXLENGTH=10 Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name) "> </td>
						<td>Til dato:</td>
						<td><INPUT NAME="tbxTilDato" TYPE=TEXT SIZE=10 MAXLENGTH=10 Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)"> </td>
						<td>Avdeling:</td>
						<td>
							<SELECT NAME="dbxAvdeling">
								<option value="0"></option>
								<%
								' Get avdeling
								strSQL = "SELECT AvdelingID, Avdeling FROM Avdeling ORDER BY  avdeling"
								Set rsAvdeling = GetFirehoseRS( strSQL, Conn )
								Do Until rsAvdeling.EOF
									If CInt(rsAvdeling("AvdelingID")) = CInt(Request("dbxAvdeling")) Then sel = " SELECTED" Else sel = "" %>
									<OPTION VALUE="<% =rsAvdeling("AvdelingID") %>" <% =sel %>><% =rsAvdeling("Avdeling") %>
									<%   
									rsAvdeling.MoveNext
								Loop
								' Close AND release recordset
								rsAvdeling.Close
								Set rsAvdeling = Nothing
								%>
							</select>
						</td>
						<td><input type="submit" name="pbnDataAction" value="Søk"></td>
					</tr>
				</table>
			</form>
			<%
			' Create table only when records found
			If  (foundRecords)  Then
				' Create table
				Response.Write "<TABLE border = 0>"
				Do Until rsRapport.EOF
					' Break on Avdeling ?
					If rsRapport( "AvdelingID") <> AvdelingID Then

						' Do we have a Oppdrag ?
						If OppdragID <> "" Or VikarID <> "" Then

							' Create avdeling heading

							Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn , Fakturapris, Timelonn  )

							OmsFirma = OmsFirma + Omsetning
							BidragFirma = BidragFirma + Dekningsbidrag
							LoennFirma = LoennFirma + Loenn
							AntTimerFirma = AntTimerFirma + AntTimer
							FaktTimerFirma = FaktTimerFirma + FaktTimer

							' Set new value
							OppdragID = ""
							VikarID = ""
							Omsetning = 0
							Dekningsbidrag = 0
							Loenn = 0
							AntTimer = 0
							FaktTimer = 0

						End If

						' Do we have a Firma ?
						If FirmaID <> "" Then
							' Create footer
							Call FirmaFooter( OmsFirma, BidragFirma, AntallOppdrag, LoennFirma , AnttimerFirma, FaktTimerFirma  )

							FirmaID = ""

							OmsAvdeling = OmsAvdeling + OmsFirma
							BidragAvdeling = BidragAvdeling + BidragFirma
							AntallAvdeling = AntallAvdeling + AntallOppdrag
							LoennAvdeling = loennAvdeling + LoennFirma
							AntTimerAvdeling = AntTimerAvdeling + AntTimerFirma
							FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerFirma

							AntTimerFirma = 0
							FaktTimerFirma = 0
							' Reset values
							AntallOppdrag = 0
							OmsetningFirma = 0
							BidragFirma = 0
							LoennFirma = 0
						End If

						' Do we have a Avdeling ?
						If AvdelingID <> "" Then

							' Create footer
							Call AvdelingFooter( OmsAvdeling, BidragAvdeling, AntallAvdeling, LoennAvdeling , AnttimerAvdeling, FaktTimerAvdeling  )

							AntallTotalt = AntallTotalt + AntallAvdeling
							OmsTotalt = OmsTotalt + OmsAvdeling
							BidragTotalt = BidragTotalt + BidragAvdeling
							LoennTotalt = LoennTotalt + LoennAvdeling
							AntTimerTotalt = AntTimerTotalt + AntTimerAvdeling
							FaktTimerTotalt = FaktTimerTotalt + FaktTimerAvdeling

							OmsAvdeling = 0
							BidragAvdeling = 0
							LoennAvdeling = 0
							AntallAvdeling = 0
							AntTimerAvdeling = 0
							FaktTimerAvdeling = 0
							Omsetning = 0
							OmsFirma = 0
						End If
						' Create avdeling heading
						Call AvdelingHeader( rsRapport  )

					End If
					' break on firma
					If rsRapport( "FirmaID") <> FirmaID Then
						' Do we have a Oppdrag ?
						If OppdragID <> "" Or VikarID <> "" Then
							' Create avdeling heading
							Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning , Loenn , Fakturapris, Timelonn )
							OmsFirma = OmsFirma + Omsetning
							BidragFirma = BidragFirma + Dekningsbidrag
							LoennFirma = LoennFirma + Loenn
							AntTimerFirma = AntTimerFirma + AntTimer
							FaktTimerFirma = FaktTimerFirma + FaktTimer
							' Set new value
							Omsetning = 0
							Dekningsbidrag = 0
							Loenn = 0
							OppdragID = ""
							VikarID = ""
							AntTimer = 0
							Fakttimer = 0

						End If

						If FirmaID <> "" Then

							' Create footer
							Call FirmaFooter( OmsFirma, BidragFirma, AntallOppdrag, LoennFirma , AnttimerFirma, FaktTimerFirma  )

							OmsAvdeling = OmsAvdeling + OmsFirma
							BidragAvdeling = BidragAvdeling + BidragFirma
							LoennAvdeling = LoennAvdeling + LoennFirma
							AntallAvdeling = AntallAvdeling + AntallOppdrag
							AntTimerAvdeling = AntTimerAvdeling + AntTimerFirma
							FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerFirma

							' Reset values
							AntTimerFirma = 0
							FaktTimerFirma = 0
							AntallOppdrag = 0
							OmsFirma = 0
							BidragFirma = 0
							LoennFirma = 0
							AntallFirma = 0

						End If

						' Create header
						Call FirmaHeader( rsRapport )

					End If

					' Break on oppdragid
					If rsRapport("OppdragID") <> OppdragID Or rsRapport("VikarID") <> VikarID or rsRapport("FakturaPris") <> FakturaPris or rsRapport("Timelonn") <> Timelonn Then

						' Do we have a Oppdrag ?
						If OppdragID <> "" Or VikarID <> "" Then
							' Create avdeling heading
							Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning , Loenn, Fakturapris, Timelonn)
						End If

						OmsFirma = OmsFirma + Omsetning
						BidragFirma = BidragFirma + Dekningsbidrag
						LoennFirma = LoennFirma + Loenn

						AntTimerFirma = AntTimerFirma + AntTimer
						FaktTimerFirma = FaktTimerFirma + FaktTimer

						Omsetning = 0
						Dekningsbidrag = 0
						Loenn = 0
						AntTimer = 0
						FaktTimer = 0

						' Set new value
						VikarID = rsRapport("VikarID")
						OppdragID = rsRapport("OppdragID")
						Firma = rsRapport( "Firma")
						Vikar = rsRapport( "Vikar")
						Fradato = rsRapport( "Fradato")
						Tildato = rsRapport( "Tildato")

						' accumulate
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
					AvdelingID = rsRapport("AvdelingID")
					FirmaID = rsRapport("FirmaID")
					VikarID = rsRapport("VikarID")
					OppdragID = rsRapport("OppdragID")
					' This will correct for ech record
					Fakturapris = rsRapport("Fakturapris")
					Timelonn = rsRapport("Timelonn")
					Medarbeider = rsRapport("Medarbeider")
					' Get next record
					rsRapport.MoveNext
				Loop

				' Do we have a Oppdrag ?
				If OppdragID <> "" Or VikarID <> "" Then

					' Create avdeling heading
					Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning , Loenn, Fakturapris, Timelonn)
					OmsFirma = OmsFirma + Omsetning
					BidragFirma = BidragFirma + Dekningsbidrag
					LoennFirma = LoennFirma + Loenn
					AntTimerFirma = AntTimerFirma + AntTimer
					FaktTimerFirma = FaktTimerFirma + FaktTimer
					AntallOppdrag = AntallOppdrag + 1
				End If

				' Do we have a Oppdrag ?
				If FirmaID <> "" Then
					' Create footer
					Call FirmaFooter( OmsFirma, BidragFirma, AntallOppdrag, LoennFirma , AnttimerFirma, FaktTimerFirma )
					OmsAvdeling = OmsAvdeling + OmsFirma
					BidragAvdeling = BidragAvdeling + BidragFirma
					LoennAvdeling = LoennAvdeling + LoennFirma
					AntTimerAvdeling = AntTimerAvdeling + AntTimerFirma
					FaktTimerAvdeling = FaktTimerAvdeling + FaktTimerFirma
					AntallAvdeling = AntallAvdeling + AntallOppdrag
				End If

				' Do we have a Avdeling ?
				If AvdelingID <> "" Then
					' Create footer
					Call AvdelingFooter( OmsAvdeling, BidragAvdeling, AntallAvdeling, LoennAvdeling , AnttimerAvdeling, FaktTimerAvdeling  )

					AntallTotalt = AntallTotalt + AntallAvdeling
					OmsTotalt = OmsTotalt + OmsAvdeling
					BidragTotalt = BidragTotalt + BidragAvdeling
					LoennTotalt = LoennTotalt + LoennAvdeling
					AntTimerTotalt = AntTimerTotalt + AntTimerFirma
					FaktTimerTotalt = FaktTimerTotalt + FaktTimerFirma
				End If
				' Create footer
				Call TotaltFooter( OmsTotalt, BidragTotalt, AntallTotalt, LoennTotalt , AnttimerTotalt, FaktTimerTotalt )
				' Close recordset
				rsRapport.Close
				' End table
				Response.Write "</table>"
			End If
			set rsRapport = Nothing
			%>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>
