<%@ LANGUAGE="VBSCRIPT" %>
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
'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' Check parameters and put into variables
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Display
'--------------------------------------------------------------------------------------------------
%>

<h1>M�nedlig rutine for � fjerne gamle timelister fra vanlig visning.</h1>
<p>Setter timelistestatus til 6 p� timelister der hvor b�de l�nnsstatus og fakturastatus = 3</p>
<FORM ACTION="Timeliste_lag_gamle_db.asp" METHOD=POST >
	<INPUT TYPE=SUBMIT VALUE="Overf�r til gamle">
</form>

<hr>

<h1>Etterfakturering eller etterl�nning der vikar ikke har flere timelister.</h1>
<FORM ACTION="Vikar_timeliste_ny3.asp?frakode=2&gmlEnr=ja" METHOD=POST >
	<p>Vikarnr: <input type="text" name=VIKARID SIZE=4 ></p>
	<p>Oppdragnr: <input type="text" name=OPPDRAGID SIZE=4 ></p>
	<INPUT TYPE=SUBMIT VALUE="S�k">
</form>

</body>
</html>