<%@ Language=VBScript %>
<%
'********************************************************************************************
'oppretter databaseforbindelse

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")


vikar=false

'********************************************************************************************
'Henter parametere for valgt måned og vikar

'vikarid =  request("vikar")
periode =  request("periode")
avdeling = request("dbxAvdeling")
avdelinger = request("dbxAvdeling")
if avdeling = "0" Then
avdeling = ""
	Set rsAvdeling = Conn.Execute("Select AvdelingID from Avdeling")
    Do Until rsAvdeling.EOF
    avdeling = avdeling & rsAvdeling("AvdelingID")&","
	rsAvdeling.MoveNext
	loop
	rsAvdeling.close
	set rsAvdeling=Nothing
	strLengde = len(avdeling)
	avdeling = left(avdeling,(strLengde-1))
end if


mnd =  mid(periode, 5)
aar = left(periode, 4)

if NOT mnd="" Then
vikar=true

SELECT	CASE mnd
		CASE 1 
			mnd2 = "Januar" 
		CASE 2  
			mnd2 = "Februar" 
		CASE 3  
			mnd2 = "Mars"	
		CASE 4  
			mnd2 = "April" 
		CASE 5 
			mnd2 = "Mai"
		CASE 6 
			mnd2 = "Juni" 
		CASE 7  
			mnd2 = "Juli" 
		CASE 8  
			mnd2 = "August"
		CASE 9 
			mnd2 = "September"
		CASE 10  
			mnd2 = "Oktober" 
		CASE 11  
			mnd2 = "November" 
		CASE 12  
			mnd2 = "Desember"
	END SELECT

'********************************************************************************************


'********************************************************************************************
'Henter opplysninger om vikar og oppdrag i perioden

Set rsFirma = Conn.Execute("Select distinct d.firmaid, f.firma "&_
							" from firma f, dagsliste_vikar D, oppdrag O "&_
							" where datepart(month, D.dato)="& mnd &_  
							" And datepart(year, D.dato)= "& aar &_
							" And d.firmaid=f.firmaid "&_							
							" And O.oppdragID = d.oppdragid "&_
							" AND O.avdelingID in("& avdeling &")")
'Response.Write(rsFirma.Source)	

End if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<script type="text/javascript">
function sjekk(){

var countBx = document.forms[1].elements.length;

for (var i=0; i < countBx; i++){
if (document.forms[1].elements[i].type == "checkbox"){
	if (document.forms[1].elements[i].checked == true)
		document.forms[1].elements[i].checked = false;
	else
		document.forms[1].elements[i].checked = true;
}//end if
}//end for
}//end function
</script>
<body>
	<div class="pageContainer" id="pageContainer">

<h1>Lønn-/ fakturagrunnlag for <% =mnd2&"  "&aar %></h1>

<form action="aggrFaktgrlagkundeStart.asp" method="post">
<%
YYYY = year(date)
MM   = month(date) - 9


if MM < 1 THEN
	MM = MM + 12
	YYYY = YYYY-1
End If


i=1
%>
	<select name="periode">
<% do until i = 13

	SELECT	CASE MM

		CASE 1 
			mnd = "januar" 
		CASE 2  
			mnd = "februar" 
		CASE 3  
			mnd = "mars"	
		CASE 4  
			mnd = "april" 
		CASE 5 
			mnd = "mai"
		CASE 6 
			mnd = "juni" 
		CASE 7  
			mnd = "juli" 
		CASE 8  
			mnd = "august"
		CASE 9 
			mnd = "september"
		CASE 10  
			mnd = "oktober" 
		CASE 11  
			mnd = "november" 
		CASE 12  
			mnd = "desember"
	END SELECT

%>
		<option value="<%=YYYY&MM%>"> <%=mnd&" "&YYYY%>
<%
MM = MM+1
If MM > 12 Then
	MM = 1
	YYYY = YYYY + 1
End If
i = i+1
loop
%>
	</select>
Avdeling:
	<select name="dbxAvdeling">
	    <option value="0">Alle
<% 
   ' Get avdeling
   Set rsAvdeling = Conn.Execute("Select AvdelingID, Avdeling from Avdeling order by avdeling")
	
      Do Until rsAvdeling.EOF 
      if rsAvdeling("AvdelingID")=cint(request("dbxAvdeling")) THEN
      	strSelected =  " SELECTED"
      else
      	strSelected = ""
      end if
      %>      
    	<option value="<% =rsAvdeling("AvdelingID") %>" <%=strSelected%>><% =rsAvdeling("Avdeling") %>
<%   rsAvdeling.MoveNext
 Loop

   ' Close and release recordset
   rsAvdeling.Close
   Set rsAvdeling = Nothing
 %>
	</select>

	<input type="submit" value=" Velg periode og avdeling  ">
</form>

<form action="aggrFaktgrlagkunde.asp">
	<input type="hidden" name="periode" value="<%=periode%>">
	<input type="hidden" name="avdelinger" value="<%=avdeling%>">
<% if vikar=true Then %>
<% if NOT rsFirma.EOF  Then %>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<th colspan="3">
			<input type=submit value="Hent rapport">
		</th>
	</tr>
	<tr>
		<th colspan="3">
			<input type=button onClick="sjekk();" value="Merk alle">
		</th>
	</tr>
	<tr>
		<th>KontaktID</th>
		<th>Kontakt</th>
		<th>&nbsp;</th>
	</tr>


<%

do until rsFirma.eof 
%>
	<tr>
		<td class="right"><%=rsFirma("firmaid")%></td>
		<td><%=rsFirma("firma")%></td>
		<td><input type="checkbox" name="valg" value="<%=rsFirma("firmaid")%>"></td>
<%

rsFirma.movenext
loop
%>
	</tr>
</table>
<input type=submit value="Hent rapport">
</form>
<% 
end if
rsFirma.close
set rsFirma = nothing
end if
%>
    </div>
</body>
</html>

