<!--#INCLUDE FILE="../includes/Library.inc"-->
<html>
<head>
	<title></title>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<%
'------------------------------------------
'	variabler
'------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

status = 3
'status = 2

'------------------------------------------
'	finne neste lønnsnummer
'------------------------------------------
sql = "select maks=max(LoenNr) from Vikar_loen_variable"
Set rsMax = conn.execute(sql)
LNR = rsMax("maks")
rsMax.Close

se "maxnr: " & LNR
'LNR = 0

'------------------------------------------
'	hovedsql
'------------------------------------------
sql = "select VikarID, Loenndato, LoenNr from Vikar_loen_variable where overfor_loenn_status = " & status &_
	" and LoenNr IS NULL" &_
	" order by VikarID, Loenndato"


Set rsLoenn = conn.execute(sql)

If not rsLoenn.EOF Then

VID = rsLoenn("VikarID")
LD = rsLoenn("Loenndato")


'------------------------------------------
'	loop
'------------------------------------------
do while not rsLoenn.EOF

'------------------------------------------
'	sette inn lønnsnummer
'------------------------------------------
	If rsLoenn("vikarID") <> VID Or rsLoenn("Loenndato") <> LD Then
		LNR = LNR + 1
		sql = "update Vikar_loen_variable set LoenNr = " & LNR &_
		" where VikarID = " & VID &_
		" and Loenndato = " & dbDate(LD)
		conn.execute(sql)
		VID = rsLoenn("VikarID")
		LD = rsLoenn("Loenndato")
	End If
	
	rsLoenn.MoveNext
loop
rsLoenn.Close: Set rsLoenn = Nothing

'oppdatere den siste...
LNR = LNR + 1
sql = "update Vikar_loen_variable set LoenNr = " & LNR &_
" where VikarID = " & VID &_
" and Loenndato = " & dbDate(LD)
conn.execute(sql)

'------------------------------------------
' vise data
'------------------------------------------
sql = "select VikarID, Loenndato, LoenNr from Vikar_loen_variable where overfor_loenn_status = " & status &_
	" and LoenNr IS NULL" &_
	" order by VikarID, Loenndato"


Set rsLoenn = conn.execute(sql)
%>
	<table cellpadding='0' cellspacing='0'>
<%
do while not rsLoenn.EOF

%>
	<tr>
		<TD><% =rsLoenn("VikarID") %></td>
		<TD><% =rsLoenn("LoennDato") %></td>
		<TD><% =rsLoenn("LoenNr") %></td>
	</tr>

<%
	rsLoenn.MoveNext
loop
rsLoenn.Close: Set rsLoenn = Nothing

End If 'ingen rader
%>
</table>
</body>
</html>  