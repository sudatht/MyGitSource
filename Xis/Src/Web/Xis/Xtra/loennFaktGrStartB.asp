<%@ Language=VBScript %>
<%
vikar=false

'********************************************************************************************
'oppretter databaseforbindelse

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")


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

Set rsVikar = Conn.Execute("Select distinct d.vikarid, Navn=(v.Fornavn+' '+v.Etternavn) "&_
							" from vikar v, dagsliste_vikar D, oppdrag O "&_
							" where datepart(month, D.dato)="& mnd &_  
							" And datepart(year, D.dato)= "& aar &_
							" And d.vikarid=v.vikarid "&_
							" And v.typeID<>3 "&_
							" And O.oppdragID = d.oppdragid "&_
							" AND O.avdelingID in("& avdeling &")")
'Response.Write(rsVikar.Source)	

'"& mnd &_
'"& aar )
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
	<meta name="generator" content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<script language="javaScript" type="text/javascript">
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
<h4>Lønn-/ fakturagrunnlag for <% =mnd2&"  "&aar %> </h4>

<form action="loennFaktGrStartB.asp" method="post">
<%
YYYY = year(date)
MM   = month(date) - 9


if MM < 1 THEN
	MM = MM + 12
	YYYY = YYYY-1
End If


i=1
%>
<SELECT NAME="periode">
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

	if (trim(YYYY&MM)= trim(request("periode"))) THEN
		sel = " SELECTED"
	Else
		sel = ""
	end if
%>

	<OPTION  VALUE="<%=YYYY&MM%>" <%=sel%>> <%=mnd&" "&YYYY%>
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
<SELECT NAME="dbxAvdeling">
    <OPTION VALUE=0>Alle
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
    <OPTION VALUE="<% =rsAvdeling("AvdelingID") %>" <%=strSelected%>><% =rsAvdeling("Avdeling") %>
<%   rsAvdeling.MoveNext
 Loop

   ' Close and release recordset
   rsAvdeling.Close
   Set rsAvdeling = Nothing
 %>
  </select>

<% 'VikarNr:<INPUT TYPE=text name="vikar" size=10>%>
<INPUT TYPE=submit value=" Velg periode og avdeling  ">
</form>
<FORM action="loennFaktGrunnlagB.asp">
<input type="hidden" name="periode" value="<%=periode%>">
<input type="hidden" name="avdelinger" value="<%=avdeling%>">
<% if vikar=true Then %>
<% if NOT rsVikar.EOF  Then %>
<table cellpadding='0' cellspacing='0'>
	<tr>
		<th colspan="3"><input type=submit value="Hent rapport"></th>
	</tr>
	<tr>
		<th colspan="3"><input type=button onClick="sjekk();" value="Merk alle"></th>
	</tr>
	<tr>
		<th>VikarID</th>
		<th>Navn</th>
		<th></th>
	</tr>


<%

do until rsvikar.eof 
%>
	<tr>
		<td class="right"><%=rsvikar("vikarid")%></td>
		<td><%=rsvikar("navn")%></td>
		<td><INPUT TYPE="CHECKBOX" NAME="valg" VALUE="<%=rsVikar("vikarid")%>"</td>
<%

rsvikar.movenext
loop
%>
	</tr>
	<tr>
		<th colspan="3"><input type=submit value="Hent rapport"></th>
	</tr>
</table>
</form>
<% 
end if
rsVikar.close
set rsVikar = nothing
end if
%>
    </div>
</body>
</html>

