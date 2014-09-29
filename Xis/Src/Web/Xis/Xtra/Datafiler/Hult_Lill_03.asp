<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
	<title>Arbeid med datafiler</title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
	<div class="contentHead1">
		<h1>Nye og endrede vikarer</h1>
	</div>
	<div class="content">
<form action="../vikarsoek.asp?Kode=3" method="post">
	<input type="submit" value="Søk">
</form>

<%
'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' Run different queries
'--------------------------------------------------------------------------------------------------
Set rsVikar = Conn.Execute("select VIKAR.VikarId, Etternavn, Fornavn, Endring, Adresse, Postnr, PostSted, Telefon, MobilTlf, EPost, " &_
			"foedselsdato, personnummer, overfort, VIKAR_ANSATTNUMMER.ansattnummer " &_
			"from VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, ADRESSE " &_
			"where VIKAR.VikarID = ADRESSE.AdresseRelID " &_
			"and ADRESSE.AdresseType = 1 " &_
			"and ADRESSE.Adresserelasjon = 2 " &_
			"and Endring = 2 " &_
			"order by VIKAR.VikarID")


' No records found
If rsVikar.BOF = True And rsVikar.EOF = True Then 
   Response.write "<p>ingen treff!</p>"
Else
%>
<div class="listing">
<FORM ACTION="merk_for_overf_db.asp" METHOD=POST>
<table>
	<tr>
		<th>Ansattnr</th>
		<th>Navn</th>
		<th>Adresse</th>
		<th>Post</th>
		<th>FDato</th>
		<th>Overføres</th>
	</tr>
<% 
'--------------------------------------------------------------------------------------------------
' Show data
'--------------------------------------------------------------------------------------------------
count = 1
Do Until rsVikar.EOF
   strFullName   = rsVikar("Etternavn") & " " & rsVikar("Fornavn")   
   strPostAddress = rsVikar("Postnr") & " " & rsVikar("PostSted")  
   fNr =  rsVikar("foedselsdato") & " " & rsVikar("personnummer")
   optname = "opt" & count 
   optname2 = "oppt" & count
   If rsVikar("overfort") = 1 Then
	merket = "checked"
   Else
	merket = ""    
   End If
%>

	<tr>
		<td><% =rsVikar("ansattnummer") %></td>
		<td><A HREF="Vikar_frames.asp?vikarID=<%=rsVikar("VikarID") %>&tilgang=3&kode=0"><% =strFullName %></A></td>
		<td><% =rsVikar("Adresse") %></td>
		<td><% =strPostAddress %></td>
		<td><% =fNr %></td>
		<td class="center">
			<INPUT class="checkbox" TYPE=CHECKBOX VALUE=checked NAME=<% =optname %> <% =" " & merket %> >
			<INPUT TYPE=HIDDEN SIZE=4 NAME=<% =optname2 %> VALUE=<% =rsVikar("VikarID") %> >
		</td>
	</tr>

<%
rsVikar.MoveNext
count = count + 1
Loop

rsVikar.Close
count = count - 1
%>

</table>
</div>
<input name="COUNT" TYPE="HIDDEN" VALUE="<% =count %>">
<input name="btnOverfoer" TYPE="SUBMIT" VALUE="Merk for overføring">

<% Response.write kode %> 
</form>

<FORM ACTION="Eksp_HL_01.asp?kode=<% =kode %>" METHOD=POST >
	<input name="btnOverfoer" TYPE="SUBMIT" VALUE="Overfør alle oppplysninger til Hult &amp; Lillevik">
</form>


<% End If 'ingen treff %>
    </div>
    </div>
</body>
</html>

