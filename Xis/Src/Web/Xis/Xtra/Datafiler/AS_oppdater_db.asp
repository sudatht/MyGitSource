<!--#INCLUDE FILE="../includes/Library.inc"--> 
<%

'--------------------------------------------------------------------------------------------------
' PARAMETERS
'--------------------------------------------------------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'seVar

dato = Request("oppdateringsDato")
VikarID = Request("VikarID")
splittuke = Request("Splittuke")
If splittuke = "" Then splittuke = "NULL"
ukenr = Request("Ukenr")
aar = Datepart("yyyy", dato)
ukenr = ukenr + (aar * 100)
'--------------------------------------------------------------------------------------------------
' oppdatere dagsliste
'--------------------------------------------------------------------------------------------------
sql = "update DAGSLISTE_VIKAR set" &_
	" Loennstatus = 3" &_
	", Loenndato = " & dbDate(date) &_
	" where dato <= " & dbDate(dato) &_
	" and fakturastatus = 3" &_
	" and loennstatus < 3" &_
	" and VikarID = " & VikarID
'se sql
conn.execute(sql)

'--------------------------------------------------------------------------------------------------
' oppdatere ukeliste
'--------------------------------------------------------------------------------------------------
ukenr2 = ukenr - 1

sql = "update VIKAR_UKELISTE set" &_
	" Overfort_loenn_status = 3" &_
	", Loenndato = " & dbDate(date) &_
	" where VikarID = " & VikarID &_
	" and Overfort_loenn_status < 3" &_
	" and Overfort_fakt_status = 3" &_
	" and ukenr <= " & ukenr2

'se sql
conn.execute(sql)



sql = "update VIKAR_UKELISTE set" &_
	" Overfort_loenn_status = 3" &_
	", Loenndato = " & dbDate(date) &_
	" where VikarID = " & VikarID &_
	" and Overfort_loenn_status < 3" &_
	" and Overfort_fakt_status = 3" &_
	" and ukenr = " & ukenr &_
	" and (notat = " & splittuke &_
	" or notat = ' ')"

'se sql
conn.execute(sql)

Response.redirect "As_detaljer_vis.asp?VikarID=" & VikarID & "&tildato=" & session("ASTildato")

%>