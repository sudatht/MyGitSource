<html>
<head>
	<title>Arbeid med datafiler</title>
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body onload=openval()>
	<div class="pageContainer" id="pageContainer">

<SCRIPT LANGUAGE="VBSCRIPT">

Function openval() 

 'document.forms(0).Starttid.value = "<% =Request.Querystring("Fratid") %>"
 'document.forms(0).Sluttid.value = "<% =Request.Querystring("Tiltid") %>"
 'document.forms(0).Dato.value = "<% =Request.Querystring("Dato") %>"
 'document.forms(0).Linje.value = "<% = Request.Querystring("Linje") %>"
 'document.forms(0).ENDRE.value = "<% = Request.Querystring("Endre") %>"

End Function

</SCRIPT>

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

If Request.QueryString("VikarID") = "" Then
   strVikarID = Request.Form("VikarID")
   strOppdragID = Request.Form("OppdragID")
   strFirmaID = Request.Form("FirmaID")
   strId = Request.Form("ID")
   kode = Request.Form("kode")
   tilgang = Request.Form("tilgang")
Else
   strVikarID = Request.QueryString("VikarID")
   strOppdragID = Request.QueryString("OppdragID")
   strFirmaID = Request.QueryString("FirmaID")
   strId = Request.QueryString("ID")
   kode = Request.QueryString("kode")
   tilgang = Request.QueryString("tilgang")
End IF
   strUkeNr = Request.QueryString("UkeNr")

'Response.write Request.Form("VikarID") & "<br>"
'Response.write Request.QueryString("VikarID") & "<br>"

'Response.write strVikarID & "<br>"
'Response.write strOppdragID & "<br>"
'Response.write strFirmaID & "<br>"
'Response.write strID & "<br>"
'Response.write kode & "<br>"
'Response.write tilgang & "<br>"

'--------------------------------------------------------------------------------------------------
' SQL for finding name
'--------------------------------------------------------------------------------------------------
strSQL = "select Navn=(Fornavn + ' ' + Etternavn) from VIKAR where Vikarid = " & strVikarID

'Response.write strSQL & "<br>"

Set rsNavn = Conn.Execute(strSQL)

strNavn = rsNavn("Navn")

rsNavn.Close


'--------------------------------------------------------------------------------------------------
' SQL for displaying loennsart
'--------------------------------------------------------------------------------------------------

Set rsLoennsart = conn.Execute("select Loennsartnr, Loennsart from H_LOENNSART order by Loennsart")

'--------------------------------------------------------------------------------------------------
' If edit find the right row to display
'--------------------------------------------------------------------------------------------------
strID = Request.QueryString("ID")
strEndre = Request.QueryString("Endre")
strLoennsartnr = ""

If strEndre = "Ja" Then

 strSQL = "select Dato, Loennsartnr, Antall, Sats  " &_ 
	"from VIKAR_UKELISTE " &_
	 "where ID = " & strID

     Set rsID = conn.Execute(strSQL)

strLoennsartnr = rsID("Loennsartnr")
End If 	

'--------------------------------------------------------------------------------------------------
' Form to register lønnsdata
'--------------------------------------------------------------------------------------------------

%>


<!--FORM ACTION="Vikar_varl_db.asp" METHOD=POST>
<input name="VikarID" TYPE=HIDDEN VALUE=<% '=strVikarID %> >
<input name="OppdragID" TYPE=HIDDEN VALUE="<% '=strOppdragID %>" >
<input name="Endre" TYPE=HIDDEN VALUE="<% '=strEndre %>" >
<input name="ID" TYPE=HIDDEN VALUE="<% '=strID %>" >

<table cellpadding='0' cellspacing='0'>

<tr>
<th>Dato
<th>Art
<th>Antall
<th>Sats
<!--TH>Beløp-->
<tr>
<!--TH><input type=text size=6 NAME=Dato <% If strID <> "" Then Response.write "VALUE=" & rsID("Dato")%> -->

<th><!--SELECT NAME="Loennsartnr">
	<OPTION VALUE= "">
	<% Do Until rsLoennsart.EOF
		If rsLoennsart("Loennsartnr") = strLoennsartnr Then
			strSelected = rsLoennsart("Loennsartnr") & " SELECTED"
		Else
			strSelected = rsLoennsart("Loennsartnr")
		End If %>
		<OPTION VALUE=<% '=strSelected %> ><% '=rsLoennsart("Loennsart")%>
	<% rsLoennsart.MoveNext
	   loop
	 rsLoennsart.Close %>
   </SELECT-->
   

<!--TH><input type=text size=6 NAME=Antall <% If strID <> "" Then Response.write "VALUE=" & rsID("Antall")%> >
<th><input type=text size=6 NAME=Sats <% If strID <> "" Then Response.write "VALUE=" & rsID("Sats")%> >
<tr>
<th colspan=6 ><INPUT TYPE=SUBMIT  VALUE="Lagre variable lønnsopplysninger"><INPUT TYPE=RESET>
</table>
</FORM-->

<% If strID <> "" Then rsID.Close %>
<%
'--------------------------------------------------------------------------------------------------
' SQL for displaying data
'--------------------------------------------------------------------------------------------------

strSQL = "select Id, VikarId, OppdragID, FirmaID, UkeNr, Loennsartnr, Antall, Sats, Belop, Notat, Dato, StatusID " &_ 
	 "from VIKAR_UKELISTE " &_
	 "where VikarID = " & strVikarID &_
	" and OppdragID = " & strOppdragID &_
	" order by UkeNr"


Response.write strSQL

Set rsVikar = Conn.Execute(strSQL)


'--------------------------------------------------------------------------------------------------
' If no record exsists
'--------------------------------------------------------------------------------------------------

If rsVikar.BOF= True And rsVikar.EOF = True Then 
   Response.write "<H4><br>Ingen godkjente timelister! <br></H4>"
%>

<% Else %>

<%
'--------------------------------------------------------------------------------------------------
' Display data
'--------------------------------------------------------------------------------------------------
%>
<H4>Sammendrag for <% =strNavn %>.</H4>


<table cellpadding='0' cellspacing='0'>

<TR class=right>
<th WIDTH=30>Dato
<th WIDTH=30>Ukenr
<th WIDTH=30>Lønnsart
<th WIDTH=50>Antall
<th WIDTH=30>Sats
<th WIDTH=20>Belop

<% 
ukeTeller = rsVikar("Ukenr")
sum = 0
sumsum = 0

'--------------------------------------------------------------------------------------------------
' LOOP
'--------------------------------------------------------------------------------------------------

do while not rsVikar.EOF %>
 
<% 
sum = sum + rsVikar("Belop")

If ukeTeller <> rsVikar("Ukenr") Then %>

<tr>
<tr><th class="right" colspan=5>Sum:<th class="right"><% =sum %>

<% ukeTeller = rsVikar("Ukenr") 
sumsum = sumsum + sum
sum = 0
  End If %>

<TR class=right>

<!---------------UPDATE---------------------->
<!--TH><A HREF=Vikar_timeliste_vis2.asp?Endre=Ja&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&ID=<%=rsVikar("ID")%>&FirmaID=<% =strFirmaID %>&UkeNr=<% =strUkeNr %> >Endre</A-->

<th><%=rsVikar("Dato")%>
<th><%=rsVikar("Ukenr")%>
<th><%=rsVikar("LoennsartNr")%>
<th><%=rsVikar("Antall")%>
<th><%=rsVikar("Sats")%>
<th><%=rsVikar("Belop")%>
<!--TH><A HREF=Vikar_timeliste_db2.asp?Slett=Ja&VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&ID=<%=rsVikar("ID")%>>Slett</A-->

<%
 rsVikar.MoveNext
   loop 
'--------------------------------------------------------------------------------------------------
' END LOOP
'--------------------------------------------------------------------------------------------------

%>

<tr><tr>
<tr><th class="right" colspan=5>Sum:<th class="right"><% =sum %>
<tr><th class="right" colspan=5>Sum:<th class="right"><% =sumsum + sum %>

<tr><tr><tr>
</table>
<table cellpadding='0' cellspacing='0'><tr>

<% If tilgang = 2 Then %>
<FORM ACTION="Vikar_timeliste_db3.asp?VikarID=<% =strVikarID %>&OppdragID=<% =strOppdragID %>&FirmaID=<% =strFirmaID %>" METHOD=POST>
<th><INPUT TYPE=SUBMIT VALUE="         Overfør til variabel lønn        "></th>
</table>
</form>
<%  End If %>

<tr><tr><tr><tr><tr>

<%  End If %>

<table cellpadding='0' cellspacing='0'><tr>
<FORM ACTION=Vikar_timeliste_vis.asp?VikarID=<%=strVikarID %>&OppdragID=<%=strOppdragID %>&FirmaID=<% =strFirmaID %>&kode=<% =kode %> METHOD=POST >
<th>
<INPUT TYPE=SUBMIT VALUE="                      Tilbake                         ">
</th>
</table>
</form>

<% rsVikar.Close %>


</body>
</html>

