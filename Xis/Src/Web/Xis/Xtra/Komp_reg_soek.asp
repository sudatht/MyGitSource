<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<title>Kompetanse - gjennomtrekk</title>
	<!--#INCLUDE FILE="../Library.inc"--> 
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script language="javascript" src=../javascript.js" type="text/javascript"></script>
<script language="javaScript" type="text/javascript">

function flytt(felt){

	felt2 = felt + '2';
//alert(felt2);

	k = felt.indexOf("Ja");
	p = felt.indexOf("Nei");
	if (k > 0 || p > 0){
		document.all[felt2].value = document.all[felt].checked;
	}else {
		document.all[felt2].value = document.all[felt].value;
	}
//alert(document.all[felt2].value);
}

</script>

</head>
<body>
	<div class="pageContainer" id="pageContainer">

<p><a href="Komp_index.asp">Tilbake til meny</a></p>

<h1>Søk etter vikar</h1>

<%
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

VikarID = Request("VikarID")
Etternavn = Request("Etternavn")

KompetanseID = Request("KompetanseID")
KompStatusID = Request("KompStatusID")
LevelID = Request("LevelID")
Prior = Request("Prior")
BekreftetNei = Request("BekreftetNei")
BekreftetJa = Request("BekreftetJa")
BevisNei = Request("BevisNei")
BevisJa = Request("BevisJa")
soek = Request("soek")

KompetanseID2 = Request("KompetanseID")
KompStatusID2 = Request("KompStatusID")
LevelID2 = Request("LevelID")
Prior2 = Request("Prior")
BekreftetNei2 = Request("BekreftetNei")
BekreftetJa2 = Request("BekreftetJa")
BevisNei2 = Request("BevisNei")
BevisJa2 = Request("BevisJa")

If KompetanseID2 = "" Then kompetanseID2 = 0
If KompStatusID2 = "" Then kompstatusID2 = 0
If LevelID2 = "" Then LevelID2 = 0
'If Prior2 = "" Then Prior2 = 0
If BekreftetJa2 = "CHECKED" Then BekreftetJa2 = true Else BekreftetJa2 = false
If BekreftetNei2 = "CHECKED" Then BekreftetNei2 = true Else BekreftetNei2 = false
If BevisJa2 = "CHECKED" Then BevisJa2 = true Else BevisJa2 = false
If BevisNei2 = "CHECKED" Then BevisNei2 = true Else BevisNei2 = false

'--------------------------------------------------------------------------------------------------
' show parameters
'--------------------------------------------------------------------------------------------------

'se VikarID & "<br>"
'se Etternavn & "<br>"

'se KompetanseID & "<br>"
'se LevelID & "<br>"
'se KompStatusID & "<br>"

'se Prior & "<br>"
'se BevisJa & "<br>"
'se BevisNei & "<br>"
'se BekreftetJa & "<br>"
'se BekreftetNei & "<br>"

'se BevisJa2 & "<br>"
'se BevisNei2 & "<br>"
'se BekreftetJa2 & "<br>"
'se BekreftetNei2 & "<br>"

'se soek & "<br>"

'--------------------------------------------------------------------------------------------------
' search in VIKAR
'--------------------------------------------------------------------------------------------------
%>

<form name="formEn" ACTION=Komp_reg_soek.asp?soek=1 METHOD=POST>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Vikarnr</th>
		<th><INPUT NAME=VikarID TYPE=TEXT SIZE=5 VALUE=<% =VikarID %>></th>
		<th>Etternavn</th>
		<th><INPUT NAME=Etternavn TYPE=TEXT SIZE=20 VALUE=<% =Etternavn %>></th>
		<th><INPUT TYPE=SUBMIT VALUE="SØK"></th>
	</tr>
</table>
</form>

<%
'--------------------------------------------------------------------------------------------------
' search in KOMPETANSE_REG
'--------------------------------------------------------------------------------------------------
%>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Program</th>
		<th>Nivå</th>
		<th>Status</th>
		<th>Prior</th>
		<th>Bekrftet</th>
		<th>Bevis</th>
		
		<td><form name="formEn" ACTION="Komp_reg_rapport1.asp?rapp=1&VikarID=<% =VikarID %>" METHOD=POST>
<INPUT TYPE=HIDDEN NAME=KompetanseID2 VALUE=<% =KompetanseID2 %>>
<INPUT TYPE=HIDDEN NAME=LevelID2 VALUE=<% =LevelID2 %>>
<INPUT TYPE=HIDDEN NAME=KompStatusID2 VALUE=<% =KompStatusID2 %>>
<INPUT TYPE=HIDDEN NAME=prior2 VALUE=<% =prior2 %>>
<INPUT TYPE=HIDDEN NAME=bekreftetJa2 VALUE=<% =bekreftetJa2 %>>
<INPUT TYPE=HIDDEN NAME=bekreftetNei2 VALUE=<% =bekreftetNei2 %>>
<INPUT TYPE=HIDDEN NAME=bevisJa2 VALUE=<% =bevisJa2 %>>
<INPUT TYPE=HIDDEN NAME=bevisNei2 VALUE=<% =bevisNei2 %>>
<INPUT TYPE=SUBMIT VALUE="Rapport"></td>
</form>

<form name="formEn" ACTION=Komp_reg_soek.asp?soek=2 METHOD=POST>
	<tr>
		<td>
<%
strSQL = "select * from H_KOMP_TITTEL where K_TypeID > 2"
Set rsKompID = conn.execute(strSQL) %>
<SELECT NAME=KompetanseID ONBLUR=flytt(this.form, this.name)>
	<OPTION VALUE=0>
	<% do while not rsKompID.EOF 
	 id = rsKompID("K_TittelID")
	 If Int(KompetanseID) = Int(id) Then val = id & " SELECTED" Else val = id %>
	<OPTION VALUE=<% =val %>><% =rsKompID("KTittel") %>	
	<% rsKompID.MoveNext 
	 loop
	 rsKompID.Close %>
</select>
<td>
<%
If Endre = "Ny" Then LevelID = 3
strSQL = "select * from H_KOMP_LEVEL"
Set rsLevelID = conn.execute(strSQL) %>
<SELECT NAME=LevelID ONBLUR=flytt(this.form, this.name)>
	<OPTION VALUE=0>
	<% do while not rsLevelID.EOF 
	 id = rsLevelID("K_LevelID")
	 If Int(LevelID) = Int(id) Then val = id & " SELECTED" Else val = id %>
	<OPTION VALUE=<% =val %>><% =rsLevelID("KLevel") %>
	<% rsLevelID.MoveNext 
	 loop
	 rsLevelID.Close %>
</select>
<td>
<%
strSQL = "select * from H_KOMPETANSE_STATUS"
Set rsStatusID = conn.execute(strSQL) %>
<SELECT NAME="KompStatusID" ONBLUR=flytt(this.form, this.name)>
	<OPTION VALUE=0>
	<% do while not rsStatusID.EOF 
	 id = rsStatusID("KompstatusID")
	 If Int(KompStatusID) = Int(id) Then val = id & " SELECTED" Else val = id %>
	<OPTION VALUE=<% =val %>><% =rsStatusID("Kompstatus") %>	
	<% rsStatusID.MoveNext 
	 loop 
	 rsStatusID.Close %>
</select>
<td><INPUT TYPE=TEXT SIZE=2 NAME=prior VALUE="<% =prior %>" ONBLUR=flytt(this.form, this.name)>
<td>Ja<INPUT TYPE=CHECKBOX NAME=bekreftetJa <% =bekreftetJa %> VALUE="CHECKED" ONBLUR=flytt(this.form, this.name)>
Nei<INPUT TYPE=CHECKBOX NAME=bekreftetNei <% =bekreftetNei %> VALUE="CHECKED" ONBLUR=flytt(this.form, this.name)>
<td>Ja<INPUT TYPE=CHECKBOX NAME=bevisJa <% =bevisJa %> VALUE="CHECKED" ONBLUR=flytt(this.form, this.name)>
Nei<INPUT TYPE=CHECKBOX NAME=bevisNei <% =bevisNei %> VALUE="CHECKED" ONBLUR=flytt(this.form, this.name)>
<td><INPUT TYPE=SUBMIT VALUE="    SØK    "></td>
</form>
</table>
<%
'--------------------------------------------------------------------------------------------------
' SQL
'--------------------------------------------------------------------------------------------------
If soek = 1 Then
If VikarID <> "" Or Etternavn <> "" Then

	where = ""
	If VikarID <> "" Then where = where & "and VikarID = " & VikarID
	If Etternavn <> "" Then	where = where & "and Etternavn Like '" & Etternavn & "%'"
	where = Mid(where,5)

strSQL = "select VikarID, Etternavn, Fornavn " &_
	"from VIKAR " &_
	"where " & where

	se strSQL & "<br>"

End If 'vikarID = ""
End If 'soek = 1

If soek = 2 Then

	where = ""
	If KompetanseID <> "0" Then where = where & " and KompetanseID = " & KompetanseID
	If LevelID <> "0" Then where = where & " and LevelID = " & LevelID
	If KompstatusID <> "0" Then where = where & " and KompStatusID = " & KompStatusID
	If Prior <> "" Then where = where & " and Prioritet = " & Prior
	If BekreftetJa <> "" Then where = where & " and NOT BekrftDato IS NULL "
	If BekreftetNei <> "" Then where = where & " and BekrftDato IS NULL "
	If BevisJa <> "" Then where = where & " and NOT BevisDato IS NULL "
	If BevisNei <> "" Then where = where & " and BevisDato IS NULL "

strSQl = "select distinct K.LinjeID, KTittel, KLevel, K.Prioritet, V.VikarID, V.Etternavn, V.Fornavn " &_
	"from KOMPETANSE_REG K, H_KOMP_TITTEL T, H_KOMP_LEVEL L " &_
	"where K.KompetanseID = T.K_TittelID " &_
	"and K.LevelID = L.K_LevelID " &_
	where

	se strSQL & "<br>"

End If 'soek = 2

	Set rsVikar = conn.execute(strSQL)

If soek <> "" Then
If not rsVikar.EOF Then
'--------------------------------------------------------------------------------------------------
' show list
'--------------------------------------------------------------------------------------------------
%>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Vikarnr</th>
		<th>Etternavn</th>
		<th>Fornavn</th>
		<th>Prior</th>
		<th>Reg</th>
<%
do while not rsVikar.EOF
%>
	</tr>
	<tr>
		<td><% =rsVikar("VikarID") %></td>
		<td><% =rsVikar("Etternavn") %></td>
		<td><% =rsVikar("Fornavn") %></td>

<% If Soek = 2 Then %>

		<td><% =rsVikar("KTittel") %></td>
		<td><% =rsVikar("KLevel") %></td>
		<td><% =rsVikar("Prioritet") %>
<% Etternavn = rsVikar("Etternavn") %>
<% param = "?VikarID=" & rsVikar("VikarID") & "&LinjeID=" & rsVikar("LinjeID") & "&Endre=Endre" %>
<% param = param & "&okovis=1" %>
<td><A HREF=Komp_reg_ny.asp<% =param %>>Se ønsker</a>

<% Else %>
<% param = "?VikarID=" & rsVikar("VikarID")  %>
<td><A HREF=Komp_reg_ny.asp<% =param %>>Se ønsker</a>

<% End If 'soek

rsVikar.MoveNext
loop
rsVikar.Close

'--------------------------------------------------------------------------------------------------
' 
'--------------------------------------------------------------------------------------------------
End If 'no rows
End If 'soek <> ""
%>
    </div>
</body>
</html>

