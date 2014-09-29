<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Integration.Contact.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.Economics.Constants.inc"-->
<%

	If (HasUserRight(ACCESS_REPORT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Sub TotaltFooter( Omsetning, Bidrag, AntallOppdrag, Loenn, AntTimer, FaktTimer )

        Faktor = Omsetning / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
        DiffTimer = ( AntTimer - FaktTimer )

        Response.Write "<tr>"
        Response.Write "<TD colspan=10><HR></TD>"
        Response.Write "<tr><TD colspan=2>Sum totalt</TD>"
        Response.Write "<TD colspan=4></TD>"
        Response.Write "<TD colspan=3>Total omsetning:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( Omsetning, 0) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Total bidrag:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( Bidrag, 0) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Total lønn:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( loenn, 0) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Antall oppdrag:</TD>"
        Response.Write "<TD colspan=1 class=right>" & AntallOppdrag & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Faktor:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber(Faktor, 2 ) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Timer lønn - timer fakt.:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber(DiffTimer, 0) & "</TD>"

	End Sub

	Sub AvdelingHeader( rsRapport  )
		' Create heading on avdeling
		Response.Write "<tr>"
		Response.Write "<TD colspan=3><H4>Avdeling: " & rsRapport("Avdeling") & "</H4></TD>"
		Response.Write "</TR>"
	End Sub


	Sub AvdelingFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer )

		Faktor =  Omsetning /  ( (Omsetning - Bidrag ) / XIS_FACTOR  )
		DiffTimer = AntTimer - FaktTimer

		Response.Write "<tr>"
		Response.Write "<TD colspan=10><HR></TD>"
		Response.Write "<tr><TD colspan=2>Sum avdeling</TD>"
		Response.Write "<TD colspan=4></TD>"
		Response.Write "<TD colspan=3>Total omsetning:</TD>"
		Response.Write "<TD colspan=1 class=right>" & FormatNumber( Omsetning, 0) & "</TD>"

		Response.Write "<tr><TD colspan=6></TD>"
		Response.Write "<TD colspan=3>Total bidrag:</TD>"
		Response.Write "<TD colspan=1 class=right>" & FormatNumber( Bidrag, 0) & "</TD>"

		Response.Write "<tr><TD colspan=6></TD>"
		Response.Write "<TD colspan=3>Total lønn:</TD>"
		Response.Write "<TD colspan=1 class=right>" & FormatNumber( loenn, 0) & "</TD>"

		Response.Write "<tr><TD colspan=6></TD>"
		Response.Write "<TD colspan=3>Antall oppdrag:</TD>"
		Response.Write "<TD colspan=1 class=right>" & AntallOppdrag & "</TD>"

		Response.Write "<tr><TD colspan=6></TD>"
		Response.Write "<TD colspan=3>Faktor:</TD>"
		Response.Write "<TD colspan=1 class=right>" & FormatNumber( Faktor, 2 ) & "</TD>"

		Response.Write "<tr><TD colspan=6></TD>"
		Response.Write "<TD colspan=3>Timer lønn - timer fakt.:</TD>"
		Response.Write "<TD colspan=1 class=right>" & FormatNumber( DiffTimer, 0) & "</TD>"
	End Sub

	Sub FirmaHeader( rsRapport  )

         ' Create table heading
         Response.Write "<tr>"
         Response.Write "<TD  colspan=3><H5>Kontakt: " & rsRapport("Firma") & "</H5></TD>"
         Response.Write "</TR>"
         Response.Write "<TD colspan=10><HR></TD>"
         Response.Write "</TR>"

         Response.Write "<th>Opp.nr.</th>"
         Response.Write "<th>Ansvarlig</th>"
         Response.Write "<th>Vikar</th>"
         Response.Write "<th>Startdato</th>"
         Response.Write "<th>Sluttdato</th>"
         Response.Write "<th>Pris</th>"
         Response.Write "<th>Lønn</th>"
         Response.Write "<th>Faktor</th>"
         Response.Write "<th>Oms.</th>"
         Response.Write "<th>DB</th>"
	End Sub

	Sub FirmaFooter( Omsetning, Bidrag, AntallOppdrag, Loenn , AntTimer, FaktTimer)

        If  ( ( Omsetning - Bidrag ) / XIS_FACTOR ) <> 0 Then
            Faktor = Omsetning  / ( ( Omsetning - Bidrag ) / XIS_FACTOR )
        Else
            Faktor = 0
        End If

        Response.Write "<tr>"
        Response.Write "<TD colspan=10><HR></TD>"
        Response.Write "<tr><TD colspan=2>Sum kontakt</TD>"
        Response.Write "<TD colspan=4></TD>"
        Response.Write "<TD colspan=3>Total omsetning:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( Omsetning, 0 ) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Total bidrag:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( Bidrag, 0) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Total lønn:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( loenn, 0) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Antall oppdrag:</TD>"
        Response.Write "<TD colspan=1 class=right>" & AntallOppdrag & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Faktor:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( Faktor, 2 ) & "</TD>"

        Response.Write "<tr><TD colspan=6></TD>"
        Response.Write "<TD colspan=3>Timer lønn - timer fakt.:</TD>"
        Response.Write "<TD colspan=1 class=right>" & FormatNumber( DiffTimer, 0) & "</TD>"
	
	End Sub

	Sub OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn, Fakturapris, Timelonn )
		' Create row
		Response.Write "<tr>"
		Response.Write "<TD> <A Href='oppdragvis.asp?OppdragID=" & OppdragID & "'>" & OppdragID & "</A></TD>"
		Response.Write "<TD>" & Medarbeider & "</TD>"
		Response.Write "<TD>" & Vikar & "</TD>"
		Response.Write "<TD>" & Fradato & "</TD>"
		Response.Write "<TD>" & Tildato & "</TD>"
		If Fakturapris <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Fakturapris, 0 )  & "</TD>"
		Else
			Response.Write "<TD class=right>" & "0" & "</TD>"
		End If

		If Timelonn <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Timelonn, 0 )  & "</TD>"
		Else
			Response.Write "<TD class=right>" & "0" & "</TD>"
		End If

		If  ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR ) <> 0 Then
			Faktor = Omsetning  / ( ( Omsetning - Dekningsbidrag ) / XIS_FACTOR )
			Response.Write "<TD class=right>" & FormatNumber( Faktor , 2 )  & "</TD>"
		Else
			Response.Write "<TD class=right>" & "</TD>"
		End If

		If Omsetning <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Omsetning, 0 )  & "</TD>"
		Else
			Response.Write "<TD class=right>" & "0" & "</TD>"
		End If

		If Dekningsbidrag <> 0 Then
			Response.Write "<TD class=right>" & FormatNumber( Dekningsbidrag, 0 ) & "</TD>"
		Else
			Response.Write "<TD class=right>" & "0" & "</TD>"
		End If

		Response.Write "</TR>"
	End Sub


	' Check input values

	' Is this first time to show this page
	If Request.Form( "tbxPageNo") <> "" Then
	' Add values FROM current page
	Fradato			= Request.Form( "tbxFradato" )
	Tildato			= Request.Form( "tbxTildato" )
	SelectAvdelingID = Request.form("dbxAvdeling")
	End If

	' Open database connection 
	SET Conn = GetConnection(GetConnectionstring(XIS, ""))	

' First time page called and search value exist ?
If Fradato <> "" And Tildato <> ""  Then

	if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
		AddErrorMessage("Fradato kan ikke være senere enn tildato!")
		call RenderErrorMessage()
	end if

    If SelectAvdelingID > 0 Then
       strSelectAvdeling = " And O.AvdelingID = " & SelectAvdelingID
   End If

   ' Get all
   strSql = "SELECT A.AvdelingID, M.MedID, DV.OppdragID, DV.OppdragVikarID, DV.VikarID, A.Avdeling, Medarbeider=M.Etternavn+' '+M.Fornavn, F.FirmaID, F.Firma, Vikar=V.Etternavn, V.TypeID, " &_
                " O.Fradato, O.Tildato, DV.FakturaPris, DV.FakturaTimer, DV.AntTimer, DV.Timelonn " &_
                "FROM DAGSLISTE_VIKAR DV, OPPDRAG O, FIRMA F, VIKAR V, AVDELING A, MEDARBEIDER M " &_
                "WHERE DV.Dato >= " & DbDate( fradato) &_
                " And DV.Dato <= " & DbDate( Tildato) &_
                " And DV.OppdragID = O.OppdragID " &_
                " And DV.OppdragID > 1 " &_
                " And DV.Anttimer > 0 " &_
                 strSelectAvdeling &_
                " And O.AvdelingID = A.AvdelingID " &_
                " And O.AnsMedID = M.MedID " &_
                " And DV.VikarID = V.VikarID " &_
                " And DV.FirmaID = F.FirmaID " &_
                 " Order by A.AvdelingID, F.Firma, DV.FirmaID, DV.OppdragID, DV.VikarID, Dv.Fakturapris, DV.Timelonn"

 ' Response.write strSql

   Set rsRapport = Conn.Execute( strSql )

   ' No records found ?
   If rsRapport.BOF = True And rsRapport.EOF = True Then
      RecordsFound = 0
   Else
      RecordsFound = 1
   End If

Else

   ' No records found
   RecordsFound = 0

End If
%>

<html>
<head>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script type="text/javascript" langauage="javaScript" src="javaScript.js"></script>
	<title>Omsetning pr. avdeling</title>
</head>

<body>
	<div class="pageContainer" id="pageContainer">

<h1>Omsetning pr kontakt <!--- <% =session("avdkontorNavn") %>--></h1>

<form ACTION="omsetning-kunde-timer.asp" METHOD="POST">
  <input type="hidden" NAME="tbxPageNo"          VALUE="1">

  <table cellpadding='0' cellspacing='0'>
    <tr>
     <TD>Fra dato:</TD>
     <TD><input name="tbxFraDato" TYPE=TEXT SIZE=10 MAXLENGTH=10 Value="<%=Fradato%>" ONBLUR="dateCheck(this.name), dateInterval(this.name)"> </TD>
     <TD>Til dato:</TD>
     <TD><input name="tbxTilDato" TYPE=TEXT SIZE=10 MAXLENGTH=10 Value="<%=Tildato%>" ONBLUR="dateCheck(this.name), dateInterval(this.name)"> </TD>
     <TD>Avdeling:</TD>
     <TD>
    <SELECT NAME="dbxAvdeling">
    <OPTION VALUE=0>
<%
   ' Get avdeling
   Set rsAvdeling = Conn.Execute("SELECT AvdelingID, Avdeling FROM Avdeling order by avdeling")

      Do Until rsAvdeling.EOF
	If CInt(rsAvdeling) = CInt(Request("dbxAvdeling"))Then sel = " SELCETED" Else sel = "" %>
    <OPTION VALUE="<% =rsAvdeling("AvdelingID") %>" <% =sel %>><% =rsAvdeling("Avdeling") %><% =Request("dbxAvdeling")%>
<%   rsAvdeling.MoveNext
   Loop

   ' Close and release recordset
   rsAvdeling.Close
   Set rsAvdeling = Nothing
 %>
   </SELECT>
   </TD>
      <td><input type="submit" name="pbnDataAction" value="     Søk    "></td>
    </tr>
  </table>
</form>

<%
' -----------------------------------------------
' Create table only when records found
' -----------------------------------------------

If  RecordsFound = 1  Then

   ' Create table
   Response.Write "<TABLE border = 0>"

   Do Until rsRapport.EOF

      ' Break on Avdeling ?
      ' *****************************************
      If rsRapport( "AvdelingID") <> AvdelingID Then

         ' Do we have a Oppdrag ?
         If OppdragID <> "" Or VikarID <> "" Then

            ' Create avdeling heading

            Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning, loenn , Fakturapris, Timelonn)

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
     ' ****************************************

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

            AntTimerFirma = 0
            FaktTimerFirma = 0

            ' Reset values
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
     ' *********************************************
      If rsRapport("OppdragID") <> OppdragID Or rsRapport("VikarID") <> VikarID or rsRapport("FakturaPris") <> FakturaPris or rsRapport("Timelonn") <> Timelonn Then

         ' Do we have a Oppdrag ?
         If OppdragID <> "" Or VikarID <> "" Then

            ' Create avdeling heading

            Call OppdragFooter( OppdragID, Medarbeider, Vikar, Fradato, Tildato, Faktor, Dekningsbidrag, Omsetning , Loenn, Fakturapris, Timelonn )

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

   ' Clear recordset
   set rsRapport = Nothing

   ' End table
   Response.Write "</table>"

End If
%>

</body>
</html>
