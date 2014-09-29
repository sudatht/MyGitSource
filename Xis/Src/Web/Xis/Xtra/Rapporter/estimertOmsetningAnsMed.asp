<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Reports.Utils.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Economics.Constants.inc"-->
<%
	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim Conn
	dim foundRecords : foundRecords = false
	
	Sub TotaltFooter( Omsetning, Bidrag, AntallOppdrag, Loenn, AntTimer, FaktTimer )

		if (omsetning<>0 AND bidrag<>0 AND omsetning<>bidrag) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if
        'Faktor = Omsetning / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
        DiffTimer = ( AntTimer - FaktTimer )

        Response.Write "<tr>"
        Response.Write "<td colspan='10'><hr></td></tr>"
        Response.Write "<tr><td colspan='2'><strong>Sum totalt</strong></td>"
        Response.Write "<td colspan='4'></td>"
        Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber(Faktor, 2 ) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber(DiffTimer, 0) & "</strong></td></tr>"

	End Sub

	Sub AvdelingHeader( rsRapport )
         Response.Write "<tr>"
         Response.Write "<td colspan='3'><h4>Avdeling: " & rsRapport("Avdeling") & "</h4></td>"
         Response.Write "</tr>"
	End Sub

	Sub AvdelingFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )

		if (omsetning <> 0 AND bidrag <> 0 AND omsetning <> bidrag) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor = 0
		end if

        'Faktor =  Omsetning /  ( (Omsetning - Bidrag ) / XIS_FACTOR  )
        DiffTimer = AntTimer - FaktTimer

        Response.Write "<tr>"
        Response.Write "<td colspan='10'><hr></td></tr>"
        Response.Write "<tr><td colspan='2'><strong>Sum avdeling</strong></td>"
        Response.Write "<td colspan='4'></td>"
        Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Faktor, 2 ) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( DiffTimer, 0) & "</strong></td></tr>"
	End Sub


	Sub MedarbeiderHeader( rsRapport  )
		' Create table heading
		Response.Write "<tr>"
		Response.Write "<td colspan='3'><h2>Ansvarlig " & rsRapport("Medarbeider") & "</h2></td>"
		Response.Write "</tr>"
		Response.Write "<td colspan='10'><hr></td>"
		Response.Write "</tr>"

		Response.Write "<tr><th>Opp.nr.</th>"
		Response.Write "<th>Kunde</th>"
		Response.Write "<th>Vikar</th>"
		Response.Write "<th>Startdato</th>"
		Response.Write "<th>Sluttdato</th>"
		Response.Write "<th>Pris</th>"
		Response.Write "<th>Lønn</th>"
		Response.Write "<th>Faktor</th>"
		Response.Write "<th>Oms.</th>"
		Response.Write "<th>DB</th></tr>"
	End Sub

	Sub MedarbeiderFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )

		if (omsetning <> 0 AND bidrag <> 0 AND omsetning <> bidrag) then
			Faktor = Omsetning / ( (Omsetning - Bidrag ) / XIS_FACTOR )
		else
			Faktor=0
		end if

        DiffTimer = AntTimer - FaktTimer

        Response.Write "<tr>"
        Response.Write "<td colspan='10'><hr></td>"
        Response.Write "<tr><td colspan='2'><strong>Sum ansvarlig</strong></td>"
        Response.Write "<td colspan='4'></td>"
        Response.Write "<td colspan='3'><strong>Total omsetning:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Omsetning, 0 ) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total bidrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Bidrag, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Total lønn:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( loenn, 0) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Antall oppdrag:</strong></td>"
        Response.Write "<td class='right'><strong>" & AntallOppdrag & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Faktor:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber( Faktor, 2 ) & "</strong></td></tr>"

        Response.Write "<tr><td colspan='6'></td>"
        Response.Write "<td colspan='3'><strong>Timer lønn - timer fakt.:</strong></td>"
        Response.Write "<td class='right'><strong>" & FormatNumber(DiffTimer, 0) & "</strong></td></tr>"

	End Sub


	Sub OppdragFooter( OppdragID, Firma, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )
            ' Create row
            Response.Write "<tr>"
            Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "WebUI/OppdragView.aspx?OppdragID=" & OppdragID, OppdragID , "Vis Oppdrag"  ) & "</td>"
            Response.Write "<td>" & Firma & "</td>"
            Response.Write "<td>" & CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & vikarid, vikar, "Vis vikar " & vikar ) & "</td>"
            Response.Write "<td>" & Fradato & "</td>"
            Response.Write "<td>" & Tildato & "</td>"
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

	' Check input values

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))


	' Is this first time to show this page
	If Request.Form("tbxPageNo") <> "" Then

		' Add values FROM current page
		Fradato = Request.Form( "tbxFradato" )
		Tildato = Request.Form( "tbxTildato" )
		SelectAvdelingID = Request.form("dbxAvdeling")

		' First time page called AND search value exist ?
		If Fradato <> "" AND Tildato <> ""  Then

			if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
				AddErrorMessage("Fradato kan ikke være senere enn tildato!")
				call RenderErrorMessage()
			end if

			If SelectAvdelingID > 0 Then
				strSelectAvdeling = " AND O.AvdelingID = " & SelectAvdelingID
			End If

			' Get all
			strSQL = "SELECT A.AvdelingID, M.MedID, DV.OppdragID, DV.OppdragVikarID, DV.VikarID, A.Avdeling, Medarbeider = M.Etternavn+' '+M.Fornavn, F.Firma, Vikar = V.fornavn + ' ' + V.Etternavn, V.TypeID, " &_
					" O.Fradato, O.Tildato, DV.Fakturapris, DV.Fakturatimer, DV.AntTimer, DV.Timelonn, DV.Dato " &_
					"FROM DAGSLISTE_VIKAR DV, OPPDRAG O, FIRMA F, VIKAR V, AVDELING A, MEDARBEIDER M " &_
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
					" ORDER BY A.AvdelingID, M.MedID, F.Firma, DV.OppdragID, DV.VikarID, Dv.Fakturapris, DV.Timelonn"

			Set rsRapport = GetFirehoseRS( strSql, Conn )

			' No records found ?
			If (HasRows(rsRapport) = true) Then
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
		<script language="javaScript" type="text/javascript" src="/xtra/Js/javaScript.js"></script>
		<title>Estimert omsetning pr. avdeling/ansvarlig/oppdrag</title>
	</head>
	<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>Estimert omsetning pr. avdeling/ansvarlig/oppdrag</h1>
		</div>
		<div class="content">
			<form name="formEn" ACTION="EstimertomsetningAnsMed.asp" METHOD="POST" ID="Form1">
					<input type="hidden" NAME="tbxPageNo" VALUE="1" ID="Hidden1">
					<table>
						<tr>
							<td>Fra dato:</td>
							<td><input name="tbxFradato" type="text" size="10" MAXLENGTH=10 Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)" ID="Text1"></td>
							<td>Til dato:</td>
							<td><input name="tbxTilDato" type="text" size="10" MAXLENGTH=10 Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)" ID="Text2"> </td>
							<td>Avdeling:</td>
							<td>
								<select name="dbxAvdeling" ID="Select1">
									<option value="0"></option>
									<%
									' Get avdeling
									strSQL = "SELECT AvdelingID, Avdeling FROM Avdeling ORDER BY avdeling"
									Set rsAvdeling = GetFireHoseRS(strSQL, Conn)
									
									Do Until rsAvdeling.EOF
										If CInt(rsAvdeling("AvdelingID")) = CInt(Request("dbxAvdeling")) Then sel = " SELECTED" Else sel = "" %>
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
							<td><input type="submit" name="pbnDataAction" value="Søk" ID="Submit1"></td>
						</tr>
					</table>
				</form>
				<%
				' Create table only when records found
				If (foundRecords = true)  Then

					' Create table
					Response.Write "<div class='listing'><table>"

					Do Until rsRapport.EOF

						' Break on Avdeling ?
						If rsRapport( "AvdelingID") <> AvdelingID Then

							' Do we have a Oppdrag ?
							If OppdragID <> "" or VikarID <> "" Then

								' Create avdeling heading
								Call OppdragFooter( OppdragID, Firma, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )

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
							If MedID <> "" Then
								' Create footer
								Call MedarbeiderFooter( OmsMedarbeider, BidragMedarbeider, AntallOppdrag, LoennMedarbeider , AntTimerMedarbeider, FaktTimerMedarbeider  )

								MedID = ""

								' Reset values
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
								Call OppdragFooter( OppdragID, Firma, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning , loenn, Fakturapris, Timelonn )

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
								Call OppdragFooter( OppdragID, Firma, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )

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
							OppdragID = rsRapport("OppdragID")
							VikarID = rsRapport("VikarID")
							Firma = rsRapport( "Firma")
							Vikar = rsRapport( "Vikar")
							Fradato = rsRapport( "Fradato")
							Tildato = rsRapport( "TilDato")

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
						MedID = rsRapport("MedID")
						VikarID = rsRapport("VikarID")
						OppdragID = rsRapport("OppdragID")

						' This will correct for ech record
						Fakturapris = rsRapport("Fakturapris")
						Timelonn = rsRapport("Timelonn")

						' Get next record
						rsRapport.MoveNext

					Loop

					' Do we have a Oppdrag ?
					If VikarID <> "" Then

						' Create avdeling heading
						Call OppdragFooter( OppdragID, Firma, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )

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
					Response.Write "</table></div>"
				End If
				%>
				<p id="sideskift" class="pageBreakAfter">&nbsp;</p>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>