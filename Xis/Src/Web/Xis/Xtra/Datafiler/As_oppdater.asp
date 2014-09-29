<html>
<head>
	<title>AS</title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<!--#INCLUDE FILE="../includes/Library.inc"--> 
<%
'seVAR
'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' prosessing parameters
'--------------------------------------------------------------------------------------------------

'strVikarID = Request("VikarID")
ASTildato = Request("ASTildato")

'--------------------------------------------------------------------------------------------------
' form
'--------------------------------------------------------------------------------------------------
%>
Timeliste til AS
<FORM HREF=AS_oppdater.asp METHOD=POST>
<input type="text" name=ASTildato >
<INPUT TYPE=SUBMIT VALUE="   Søk  ">
</form>

<%
'--------------------------------------------------------------------------------------------------
' søkeresultat
'--------------------------------------------------------------------------------------------------
If ASTildato <> "" Then
	session("ASTildato") = ASTildato

	SQL = "select distinct D.VikarID Etternavn, Fornavn " &_
	"from  DAGSLISTE_VIKAR D, VIKAR, FIRMA " &_
	"where D.TimelisteVikarStatus < 6 "  &_
	" and D.Loennstatus < 3 " &_
	" and VIKAR.TypeID = 3" &_
	" and D.VikarID = VIKAR.VikarID " &_
	"and D.Dato < " &DbDate(session("ASTildato"))  &_
	"and D.FirmaID = FIRMA.FirmaID " &_
	" order by VIKAR.TypeID, Etternavn"

se SQL

set rsAS = conn.execute(sql)
If Not rsAS.EOF Then

'--------------------------------------------------------------------------------------------------
' kolonner
'--------------------------------------------------------------------------------------------------
%>

<table cellpadding='0' cellspacing='0'>
<tr>
<th>VikarID
<th>Navn
<th>Kontakt

<%
'--------------------------------------------------------------------------------------------------
' søkeresultat
'--------------------------------------------------------------------------------------------------

do while not rsAS.EOF
%>
	<tr>
	<TD><% =rsAS("VikarID") %>
	<TD><A HREF="As_detaljer_vis.asp<% =rsAs("Etternavn") & " " & rsAS("Fornavn") %>
	<TD><% =rsAs("Firma") %>
	<%
	rsAs.MoveNext
loop







