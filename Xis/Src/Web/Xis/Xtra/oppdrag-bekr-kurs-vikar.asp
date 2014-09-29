<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/Library.inc"-->
<HTML>
<HEAD>
<TITLE>Bekreftelse av oppdrag</TITLE>
<link rel="stylesheet" href="/xtra/css/
stilsett.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</HEAD>
<body>
	<div class="pageContainer" id="pageContainer">
<%
profil = Session("Profil")

' Open database connection
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'-----------------------------------------------------------------------------
' Find chosen vikarID
'-----------------------------------------------------------------------------
i = 0
first = True
val2 = Request("vid1")


do while Not val2=""
   i = i + 1
   name = "opt" & i
   name2 = "vid" & i
   val = Request(name)
   val2 = Request(name2)
   If val = "CHECKED" Then
	If Not first Then
		%><P ID="sideskift" STYLE="page-break-after='always'">_____________________________<br>sign.</P><%
	Else
		first = False
	End If

	lVikarID = val2


'-----------------------------------------------------------------------------
' Check input values
'-----------------------------------------------------------------------------

' Check parameter Fradato
If Request("Fradato") <> "" Then
   strFradato = Request.Querystring("Fradato")
Else
   Response.Write "Fradato mangler"
   Response.End
End If

' Check parameter Tildato
If Request("Tildato") <> "" Then
   strTildato = Request.Querystring("Tildato")
Else
   Response.Write "Tildato mangler"
   Response.End
End If

strOldfraDato = ""


' Get information from database
' --------------------------

' Get Vikar information
strSql = "Select VikarId, Vikarnavn=Fornavn + ' '+ Etternavn, Adresse, PostAdr=Postnr+' '+Poststed " &_
                            "from VIKAR V, ADRESSE A " &_
                            "where VikarID = " & lVikarID &_
                            " and V.VikarID = A.adresseRelID " & _
                           " and A.AdresseRelasjon = 2 and A.AdresseType = 1 "
Set rsVikar  = Conn.Execute( strSQL )

' Error from database ?
If Conn.Errors.Count > 0 then
   Call SqlError()
End if

' ************************************
' Loop on lines to create days
' ************************************

' create SQL statement
strSQL = "Select OV.OppdragVikarId, OV.OppdragID, OV.VikarID, OV.FirmaID, OV.Fradato, OV.Tildato " &_
                  " from OPPDRAG_VIKAR OV, Oppdrag O" &_
                  " where OV.VikarID = " & lVikarID &_
                  " and OV.StatusID Not in ( 3, 6 ) " &_
                  " and OV.OppdragID = O.OppdragID " &_
                  " and O.Oppdragskode = 1 " &_
                 " and ( ( OV.Fradato <= " & DbDate( strFradato ) & " and  OV.Tildato >= " & DbDate( strTildato )  & " ) " &_
                      " or (  OV.Fradato >= " & DbDate( strFradato )  & " and  OV.Fradato <= " & DbDate( strTildato ) & " ) " &_
                     " or (  OV.Tildato >= " & DbDate( strFradato )  & " and  OV.Tildato <= " & DbDate( strTildato ) & " ) ) "

' Get oppdrags vikarer
Set rsOppdragVikar = Conn.Execute( strSQL )

' Delete all rows in database
Conn.Execute( "Delete from BEKREFTELSE_KURS where BrukerID = " & Session("brukerID") & " and VikarID = " & lVikarID  )

' Error from database
If Conn.Errors.Count > 0 then
    Call SqlError()
End if

' Read all accepted vikar and create timeliste
Do Until rsOppdragVikar.EOF

       ' Create sql-statement for Procedure Lag_timeliste
      strSQL = "Execute Lag_Bekreftelse_kurs " & rsOppdragVikar("OppdragVikarID") &_
			"," & rsOppdragVikar("OppdragID") &_
			"," & rsOppdragVikar("VikarID") &_
			"," & rsOppdragVikar("FirmaID") &_
			"," & DbDate( rsOppdragVikar("Fradato") ) &_
			"," & DbDate( rsOppdragVikar("TilDato") ) &_
			"," & Session("brukerID")

      ' Run Lag_timeliste Procedure
      Conn.Execute( strSQL )

       ' Error from database
       If Conn.Errors.Count > 0 then
          Call SqlError()
       End if

      ' Get next OPPDRAG_VIKAR
      rsOppdragVikar.MoveNext
Loop

' Close and release recordset
rsOppdragVikar.Close
Set rsOppdragVikar = Nothing

' Get all details
strSql = "Select BK.Dato, OV.Fradato, OV.Timeloenn, OV.Timepris, O.Deltagere, OV.OppdragID, OV.FirmaID, F.Firma, F.Telefon, O.ArbAdresse, "&_
                "OV.Tildato, OV.Frakl, OV.Tilkl, Ansvarlig=M.Fornavn+' '+M.Etternavn, Kontaktperson=K.Fornavn+' '+K.Etternavn, " &_
                "O.Beskrivelse, O.Notatvikar, Program=KO.KTittel, Kompniva=KL.KLevel, D.Dokumentasjon, T.KursType " &_
                "from BEKREFTELSE_KURS BK, OPPDRAG_VIKAR OV, OPPDRAG O, FIRMA F, MEDARBEIDER M , KONTAKT K," &_
                "H_KOMP_TITTEL KO, H_KOMP_LEVEL KL, H_KURS_DOK D, H_KURS_TYPE T " &_
                "where BK.BrukerID = " & Session("brukerID")   &_
                " and BK.VikarID = " & lVikarID &_
                " and BK.Dato >= " & dbDate( strFraDato ) &_
                " and BK.Dato <= " & dbDate( strTilDato ) &_
                " and BK.OppdragVikarID = OV.OppdragVikarID " &_
               " and BK.FirmaID = F.FirmaID "&_
               " and BK.OppdragID = O.OppdragID " &_
               " and O.Oppdragskode = 1 "&_
               " and O.bestilltav *= K.KontaktID " &_
               " and O.AnsMedID *= M.MedID " &_
               " and O.ProgramID *= KO.K_TittelID and KO.K_TypeID=3 " &_
               " and O.Kompniva *= KL.K_LevelID " &_
               " and O.DokID *= D.OppdragdokID " &_
               " and O.TypeID  *= T.KurstypeID " &_
               " order by BK.dato, BK.OppdragID"

'Response.write strSql

 Set rsOppdragVikar = Conn.Execute( strSql )

' Error from database ?
If Conn.Errors.Count > 0 then
   Call SqlError()
End if


'Hent stedsnavn fra pålogget brukers avdelingskontor
strSQL = "SELECT [Lokasjon].[Navn] " & _
"FROM [Lokasjon] " & _
"INNER JOIN [Avdelingskontor] ON [Avdelingskontor].[LokasjonID] = [Lokasjon].[LokasjonID] " & _
"INNER JOIN [medarbeider] ON [medarbeider].[AvdelingskontorID] = [Avdelingskontor].[ID] " & _
"WHERE [Medarbeider].[MedID] =" & Session("medarbID")


Set rsAvdKontor = Conn.Execute (strSql)

' Error from database ?
If Conn.Errors.Count > 0 then
   Call SqlError()
End if

If not rsAvdKontor.EOF then
	Sted = rsAvdKontor("navn") & ",&nbsp;"
Else
	Sted = ""
End If
rsAvdKontor.close
set rsAvdKontor = Nothing

' Create Oppdragsbeksreftelse
' ----------------------------
%>


<table align=center WIDTH="670">
<tr>
    <td WIDTH="50"></td>
    <td><img align="right" src="http://www.xtra.no/images/xis/xtra_logo.gif" alt="Xtra logo"></td>
   <tr>
   <td></td>
   <tr>
   <td><p> <br><p> <br></td>
  <tr>
    <td></td>
    <td>
    <font Size="+1">
    <b><%=rsVikar("Vikarnavn")%></b><br>
    <%=rsVikar("Adresse")%><br>
    <%=rsVikar("PostAdr")%>
<%
   ' Close and release recordset
   rsVikar.Close
   Set rsVikar = Nothing
   %>

    </font>
    </td>
  <tr>
   <td></td>
   <td ALIGN="RIGHT"><%=Sted %> <%=Date() %><p> <br></td>

</table>
<p>
<h4 ALIGN="CENTER">Oversikt over kurs i perioden <%=strFradato %> til <%=strTildato %> fra Xtra</h4>

<table WIDTH="650" ALIGN="CENTER">

<%
Do Until rsOppdragVikar.EOF

   ' Do we have a new date ?
   If rsOppdragVikar("dato") <> strOldfraDato Then

	  ' Print new date line
	  Response.write "<TR><TD COLSPAN=5><B><HR></B></TD>"
	  Response.write "<TR><TD>" & WeekdayName( WeekDay(rsOppdragVikar("dato"))) &"</TD>"
	  Response.write "<TD>" & rsOppdragVikar("dato") &"</TD>"

      ' Set new olddate
      strOldfraDato=rsOppdragVikar("dato")

    Else
       ' Do we have a new Oppdrag ?
       If rsOppdragVikar("OppdragID") <> strOldOppdragID Then
          ' Print new date line
          Response.write "<TR><TD></TD><TD COLSPAN=4><HR></TD>"

         ' Set new olddate
         strOldOppdragID=rsOppdragVikar("OppdragID")

       End If

    End If

   %>
   <tr>
    <td></td>
    <th ALIGN="LEFT">Oppdrag:</th>
    <td COLSPAN="3"><%=rsOppdragVikar("OppdragID") & " " & rsOppdragVikar("Beskrivelse")%></td>

   <tr>
    <td></td>
    <td></td>
    <td></td>
    <th ALIGN="LEFT">Lønn/honorar:</th>
    <td><%=rsOppdragVikar("Timeloenn")%></td>

   <tr>
    <td></td>
    <th ALIGN="LEFT">Kontakt:</th>
    <td><%=rsOppdragVikar("Firma")%></td>
    <th ALIGN="LEFT">Program:</th>
    <td><%=rsOppdragVikar("Program")%></td>
    <td></td>

   <tr>
    <td></td>
    <th ALIGN="LEFT">Arbeidstid:</th>
    <td><%=FormatDateTime( rsOppdragVikar("FraKl"), 4)%> - <%=FormatDateTime( rsOppdragVikar("TilKl"), 4)%></td>
    <th ALIGN="LEFT">Versjon:</th>
    <td><%=rsOppdragVikar("Deltagere")%></td>
   <tr>
    <td></td>
    <th ALIGN="LEFT">Arbeidsadresse:</th>
    <td><%=rsOppdragVikar("ArbAdresse")%></td>
    <th ALIGN="LEFT">Nivå:</th>
    <td><%=rsOppdragVikar("KompNiva")%></td>
   <tr>
    <td></td>
    <th ALIGN="LEFT">Kontaktperson:</th>
    <td><%=rsOppdragVikar("Kontaktperson")%></td>
    <th ALIGN="LEFT">Dokumentasjon:</th>
    <td><%=rsOppdragVikar("Dokumentasjon")%></td>
   <tr>
    <td></td>
    <th ALIGN="LEFT">Telefon Kontakt:</th>
    <td><%=rsOppdragVikar("Telefon")%></td>
    <th ALIGN="LEFT">Type kurs:</th>
    <td><%=rsOppdragVikar("Kurstype")%></td>

   <tr>
    <td COLSPAN="5" ALIGN="CENTER"><b>Oppdragsansvarlig hos Xtra: </b><%=rsOppdragVikar("Ansvarlig")%> </td>


<%
   ' Get next record
   rsOppdragVikar.MoveNext
Loop

rsOppdragVikar.Close: Set rsOppdragVikar = Nothing
%>

<TR><TD COLSPAN=5><B><HR></B></TD>
</table>

<br>
<table WIDTH="650" ALIGN="CENTER">

<tr><td><B>Kommentar til vikar:</B><BR><BR></td>

<%
' Get all comments conneceted to selected Oppdrag
'  -----------------------------------------------------
' create SQL statement
strSql = "Select O.Oppdragid, O.Beskrivelse, O.NotatVikar " &_
             " from BEKREFTELSE_KURS BK, Oppdrag O" &_
             " where BK.BrukerID = " & Session("brukerID")   &_
             " and BK.VikarID = " & lVikarID &_
             " and BK.Dato >= " & dbDate( strFraDato ) &_
             " and BK.Dato <= " & dbDate( strTilDato ) &_
             " and BK.OppdragID = O.OppdragID " &_
             " and (NOT notatvikar IS NULL)"

' Get Comments in period
Set rsNotat = Conn.Execute( strSql )

OID = 0
Do Until rsNotat.EOF

	If OID <> rsNotat("OppdragID") Then
		OID = rsNotat("OppdragID")
		felt = Trim(rsNotat("NotatVikar"))

		If Not felt = "" And Not felt = " " Then %>
			<TR>
			<TH ALIGN=LEFT>Oppdrag: <% =rsNotat("OppdragID") %>&nbsp;<% =rsNotat("Beskrivelse") %></TH>
			<TR><TD><% =felt %></TD>
			<TR><TD><BR>
		<% End If
	End If

	rsNotat.MoveNext
Loop

rsNotat.Close: Set rsNotat = Nothing %>

</TABLE>


<TABLE WIDTH="650" ALIGN="CENTER">
<TR><TD><B><HR></B></TD>
<TR><TD>
<TEXTAREA COLS=77 ROWS=8>
Oppdragsbekreftelsen vil være grunnlag for honorar/lønn. Dersom denne avviker fra dine lister må du kontakte Xtra umiddelbart.
Faktura sendes en - 1 - gang pr. mnd., senest fem - 5 - virkedager før den 12. i hver mnd.
Honorar/lønn utbetales hver 12. i måneden.
</TEXTAREA>
</TD>
</TABLE>
<%

End If 'noen er avkrysset

loop

If lVikarID = "" Then
   Response.write "Du har ikke krysset av noen!"
   Response.End
End If

%>
<P>&nbsp;</P>
    </div>
</body>
</html>
