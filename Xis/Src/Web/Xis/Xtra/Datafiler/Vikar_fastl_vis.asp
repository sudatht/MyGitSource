<html>
<head>
	<title>Arbeid med datafiler</title>
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
' prosessing parameters
'--------------------------------------------------------------------------------------------------

If Request.QueryString("VikarID") = "" Then
   strVikarID = Request.Form("VikarID")
   'strNavn = Request.Form("Navn")
   strAvdeling = Request.Form("Avdeling")
Else
   strVikarID = Request.QueryString("VikarID")
   'strNavn = Request.QueryString("Navn")
   strAvdeling = Request.QueryString("Avdeling")
End IF

'--------------------------------------------------------------------------------------------------
' SQL for finding name
'--------------------------------------------------------------------------------------------------
strSQL = "select Navn=(Fornavn + ' ' + Etternavn) from VIKAR where Vikarid = " & strVikarID

'Response.write strSQL & "<br>"

Set rsNavn = Conn.Execute(strSQL)

strNavn = rsNavn("Navn")

rsNavn.Close


'Response.write strVikarID & "<br>"
'Response.write strNavn & "<br>"
'Response.write strAvdeling & "<br>"

'--------------------------------------------------------------------------------------------------
' SQL for displaying loennsart
'--------------------------------------------------------------------------------------------------

Set rsLoennsart = conn.Execute("select Loennsartnr, Loennsart from H_LOENNSART order by Loennsart")


'--------------------------------------------------------------------------------------------------
' If edit find the right row to display
'--------------------------------------------------------------------------------------------------
strID = Request.QueryString("ID")
strLoennsartnr = ""
If strID <> "" Then
 strEndre="Ja"
 strSQL = "select Loennsart, Antall, Sats, Beloep, Saldo " &_ 
	"from VIKAR_LOENN_FASTE " &_
	 "where ID = " & Request.QueryString("ID")

  Set rsID = conn.Execute(strSQL)

  strLoennsartnr = rsID("Loennsart") 
 
End If 	


'--------------------------------------------------------------------------------------------------
' Form to register faste lønnsdata
'--------------------------------------------------------------------------------------------------

%>

<FORM ACTION="Vikar_fastl_db.asp" METHOD=POST>
<input name="VikarID" TYPE=HIDDEN VALUE=<% =strVikarID %> >
<input name="Navn" TYPE=HIDDEN VALUE="<% =strNavn %>" >
<input name="Avdeling" TYPE=HIDDEN VALUE="<% =strAvdeling %>" >
<input name="Endre" TYPE=HIDDEN VALUE="<% =strEndre %>" >
<input name="ID" TYPE=HIDDEN VALUE="<% =strID %>" >

<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Lønnsart</th>
		<th>Antall</th>
		<th>Sats</th>
	</tr>
	<tr>
		<td>
			<SELECT NAME="Loennsartnr">
				<OPTION VALUE=""></option>
	<% Do Until rsLoennsart.EOF
		If rsLoennsart("Loennsartnr") = strLoennsartnr Then
			strSelected = rsLoennsart("Loennsartnr") & " SELECTED"
		Else
			strSelected = rsLoennsart("Loennsartnr")
		End If %>
				<OPTION VALUE="<% =strSelected %>"><% =rsLoennsart("Loennsart")%></option>
	<% rsLoennsart.MoveNext
	   loop
	 rsLoennsart.Close %>
			</SELECT>
		</td>
		<td><input type="text" NAME="Antall" <% If strID <> "" Then Response.write "VALUE=" & rsID("Antall")%>></td>
		<td><input type=text size=8 NAME=Sats <% If strID <> "" Then Response.write "VALUE=" & rsID("Sats")%> ></td>
		<!--TH><input type=text size=8 NAME=Beloep <% If strID <> "" Then Response.write "VALUE=" & rsID("Beloep")%> -->
	</tr>	
</table>

	<input type="submit" value="Registrer faste lønnsopplysninger"> <input type="reset" value="Tilbakestill">
</form>

<% If strID <> "" Then rsID.Close %>
<%
'--------------------------------------------------------------------------------------------------
' SQL for displaying data
'--------------------------------------------------------------------------------------------------

strSQL = "select Id, Loennsart, Antall, Sats, Beloep, Saldo " &_ 
	 "from VIKAR_LOENN_FASTE " &_
	 "where VikarID = " & strVikarID

'Response.write strSQL

Set rsVikar = Conn.Execute(strSQL)

'--------------------------------------------------------------------------------------------------
' If no record exsists
'--------------------------------------------------------------------------------------------------

If rsVikar.BOF= True And rsVikar.EOF = True Then 
   Response.write "<H4><br>Ingen faste lønnsopplysninger for " & strNavn  & " ! <br></H4>"
%>

<% Else %>

<%
'--------------------------------------------------------------------------------------------------
' Display data
'--------------------------------------------------------------------------------------------------
%>
<h1>Registrerte faste lønnsopplysninger for <% =strNavn %>.</h1>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Lønnsart</th>
		<th>Antall</th>
		<th>Sats</th>
		<th>Beløp</th>
		<th>Saldo</th>
		<th>Slett</th>
	</tr>

<% do while not rsVikar.EOF %>

	<tr>
		<td><!-- UPDATE -->
			<a href="Vikar_fastl_vis.asp?VikarID=<%=strVikarID %>&Avdeling=<%=strAvdeling %>&ID=<%=rsVikar("ID")%>"><%=rsVikar("Loennsart")%></a>
		</td>
		<td><%=rsVikar("Antall")%></td>
		<td><%=rsVikar("Sats")%></td>
		<td><%=rsVikar("Beloep")%></td>
		<td><% =rsVikar("Saldo")%></td>
		<td><!-- DELETE -->
			<a href="Vikar_fastl_db.asp?ID=<%=rsVikar("ID")%>&Slett=Ja&Avdeling=<%=strAvdeling %>&VikarID=<%=strVikarID%>">Slett</a>
		</td>
	</tr>
<% rsVikar.MoveNext
   loop %>
</table>
<% End If %>
<% rsVikar.Close %>

</body>
</html>

