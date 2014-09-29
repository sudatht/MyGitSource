<!--#INCLUDE FILE="includes\Library.inc"-->
<%

'Variabler
oppdragID = Request("tbxOppdragID")
frakode = Request("frakode")

If frakode = 3 Then
'	se frakode & "tilbake"
'	se (session("antLister") + 2)
End If

session("frakode") = frakode

' Open database connection
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")


'	innlegg av nye timelister
If session("hvorerjegNr") > 0 Then

If frakode = 1 Then   'første gangs kall (fra avdkontor)

'-----------------sette faste parametere---------------------------

	session("teller") = 0
	session("oppdragid") = Request("tbxoppdragID")
	session("osloURL") = "http://intranett.xtra.no/xtra/Timeliste_mellom_oslo.asp"
	session("redir") = session("osloURL") &_
		"?avdkontor=" & session("hvorerjegNR") &_
		"&backURL=" & session("backURL")

'-------------------hente linjer fra Oppdrag_vikar--------------------

	strSQL = "Select O.OppdragVikarId " &_
	" from OPPDRAG_VIKAR O, VIKAR V" &_
	" where O.OppdragID = " & session("OppdragID") &_
	" and  O.StatusID = 4 and O.Timeliste = 0 " &_
	" and O.VikarID = V.VikarID" 

'se strSQL
	Set rsOppdragVikar = Conn.Execute( strSQL )

	i = 0
	do while not rsOppdragVikar.EOF
		i = i + 1
		Redim Preserve idListe(i)
		idListe(i - 1) = rsOppdragVikar("OppdragvikarID")
		rsOppdragvikar.MoveNext
	loop: rsOppdragVikar.Close: Set rsOppdragvikar = Nothing
	session("idListe") = idListe
	session("antLister") = i

	frakode = frakode + 1

End If 'frakode = 1

'---------------------------------------------------------------------------------
'	kjøres for hver linje i oppdragvikar det skal lages tliste for
'---------------------------------------------------------------------------------

'se frakode & " " & session("antLister") & " " & idListe(0)

If Int(frakode) > 1 And Int(frakode) < (session("antLister") + 2) Then 'kjøres hvergang (andre gang etter kall fra oslo)

	Dim idListe2: idListe2 = session("idListe")
	'for i = 0 To session("antLister") - 1
	' se idListe2(i) & " " & i
	'next
'--------------------------hente linje i oppdrag_vikar---------------------------------

	strSQL = "Select O.OppdragVikarId, O.OppdragID, O.VikarID, V.TypeID, O.FirmaID, O.Fradato, O.Tildato, " &_
	"O.Frakl, O.TilKl, O.AntTimer, O.Timeloenn, O.Timepris, O.Lunch " &_
	" from OPPDRAG_VIKAR O, VIKAR V" &_
	" where O.OppdragID = " & session("OppdragID") &_
	" and  O.StatusID = 4 and O.Timeliste = 0 " &_
	" and O.VikarID = V.VikarID" &_
	" and O.OppdragvikarID = " & idListe2(frakode-2)

'se strSQL
	
	Set rsOV = Conn.Execute( strSQL )

'-------------------------sette parameter til timelisten------------------------------ 

	strSQL = "Select BestilltAv from OPPDRAG where OppdragID = " & rsOV("OppdragID") 
	Set rsOppdrag = Conn.Execute( strSQL )
	bestilltav = rsOppdrag("BestilltAv"):rsOppdrag.Close: Set rsOppdrag = Nothing
	
	If rsOV("AntTimer") = "" Then TimerPrDag = 0 Else TimerPrDag = rsOV("AntTimer")
	Call fjernKomma(TimerPrDag)
	If rsOV("TypeID") = 1 Then Vikartype = 1 Else Vikartype = 0
	FraKl = FormatDateTime( rsOV( "FraKl" ), 3 )
	TilKl = FormatDateTime( rsOV( "TilKl" ), 3 )
	Lunch = FormatDateTime( rsOV( "Lunch" ), 3 )

	timelisteparametre = "&OppdragVikarID=" & rsOV("OppdragVikarID") &_ 
		"&oppdragID=" & session("OppdragID") &_
		"&vikarID=" & rsOV("VikarID") &_
		"&firmaID=" & rsOV("FirmaID") &_
		"&fradato=" & rsOV("Fradato") &_
		"&tildato=" & rsOV("TilDato") &_
		"&FraKl=" & Frakl &_
		"&TilKl=" & Tilkl &_
		"&TimerPrDag=" & TimerPrDag &_
		"&Lunch=" & Lunch &_
		"&Timeloenn=" & rsOV( "Timeloenn" ) &_
		"&Timepris=" & rsOV( "Timepris" ) &_
		"&vikarType=" & VikarType &_
		"&Bestilltav=" & Bestilltav &_
		"&TimelsiteVikarstatus=0" 

'se timelisteparametre
	rsOV.Close
End If 

'---------------------------------------------------------------------------------
'	oppdatering av oppdrag_vikar etter at timeliste er laget
'---------------------------------------------------------------------------------
If Int(frakode) > 2 And Int(fraKode) <= (session("antLister") + 2) Then 'kjøres hvergang (andre gang etter kall fra oslo)

	Dim idListe3
	idListe3 = session("idListe")
	strSQL = "Update oppdrag_vikar set Timeliste = 1 where oppdragvikarid =" & idListe3(frakode-3)
'se strSQL	
	Conn.Execute(strSQL)

'se "oppdatering"
End If

'---------------------------------------------------------------------------------
'	etter siste kall fra oslo
'---------------------------------------------------------------------------------
If Int(frakode) = (session("antLister") + 2) Then 'siste gang dette kalles fra oslo

'se session("frakode")
'se session("teller")
'se "siste kall"

    Response.redirect "oppdragvis.asp?OppdragID=" & session("OppdragID")

Response.End
	
End If

'---------------------------------------------------------------------------------
'	kall til oslo for å lage timelister (frakode = 2, frakode >= antLister)
'---------------------------------------------------------------------------------

session("teller") = session("teller") + 1
	redir = session("redir") & timelisteparametre & "&frakode=" & frakode
'se redir
	Response.redirect redir

'---------------------------------------------------------------------------------
'	slutt
'---------------------------------------------------------------------------------
Else 'jeg er ikke avdkontor

    Response.redirect "oppdragdb.asp?OppdragID=" & Request("tbxOppdragID") & "&pbnDataAction=[W] Lag Timeliste"

End If 'hvorerjeg > 0 (avdkontor)

%>