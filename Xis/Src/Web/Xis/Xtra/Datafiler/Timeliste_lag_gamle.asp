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

<h1>Månedlig rutine for å fjerne gamle timelister fra vanlig visning.</h1>
<p>Setter timelistestatus til 6 på timelister der hvor både lønnsstatus og fakturastatus = 3</p>
<FORM ACTION="Timeliste_lag_gamle_db.asp" METHOD=POST >
	<INPUT TYPE=SUBMIT VALUE="Overfør til gamle">
</form>

<hr>

<h1>Etterfakturering eller etterlønning der vikar ikke har flere timelister.</h1>
<FORM ACTION="Vikar_timeliste_ny3.asp?frakode=2&gmlEnr=ja" METHOD=POST >
	<p>Vikarnr: <input type="text" name=VIKARID SIZE=4 ></p>
	<p>Oppdragnr: <input type="text" name=OPPDRAGID SIZE=4 ></p>
	<INPUT TYPE=SUBMIT VALUE="Søk">
</form>

</body>
</html>