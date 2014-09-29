<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = 0%>
<% 

' Connect to database
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

' Get adressetype
Set rsAdrType = Conn.Execute("select AdrtypeID, Adressetype from H_ADRESSE_TYPE where AdrTypeID > 1 order by AdrtypeID")

' Get suspect info
Set rsFirm  = Conn.Execute("Select firmaid, firma from SUSPECT where FirmaID = " & Request.QueryString("ID"))

strID       = rsFirm("FirmaID")
strName     = rsFirm("Firma")
lAdrRelasjon = 1

' Create page heading 
strHeading = "Adresse for " & strName

' Existing adress ?
If Request.QueryString("AdrID") <> "" Then
   Set rsAdresse= Conn.Execute("Select A.Adresse, A.Postnr, A.Poststed, T.AdrTypeID, T.AdresseType from SUSPECT_ADRESSE A, H_ADRESSE_TYPE T " & _
				"where A.AdrID = " & Request.QueryString("AdrID") & _
				" and A.Adressetype = T.AdrtypeID")

   lAdrtype    = rsAdresse("AdrTypeID")
   strAdress   = rsAdresse("Adresse")
   strPostNr   = rsAdresse("Postnr")
   strPoststed = rsAdresse("Poststed")
	    
End If

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
	<title>Adresse</title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">

<form name="formEn" ACTION="suspectadressedb.asp" METHOD="POST">
<input TYPE="HIDDEN" NAME="tbxRelasjon" VALUE="<%=Request.QueryString("Relasjon") %>">
<input TYPE="HIDDEN" NAME="tbxID" VALUE="<%=Request.QueryString("ID") %>">
<input TYPE="HIDDEN" NAME="tbxAdrID" VALUE="<%=Request.QueryString("AdrID")%>">

<table cellpadding='0' cellspacing='0'>
<h3><%=strHeading%></h3>
<tr>
<td>Type:</td>
<td COLSPAN="2"><select NAME="dbxAdrType">
<% 
Do Until rsAdrtype.EOF 

   If rsAdrtype("adrtypeID")  = lAdrtype  Then
      strSelected = rsAdrtype("AdrtypeID") & " " & "SELECTED"
   Else
      strSelected = rsAdrtype("AdrtypeID")
   End If
%>
    <option VALUE="<%=strSelected%>"><%=rsAdrtype("Adressetype")%>
<% 
rsAdrtype.MoveNext
Loop
%>
</select></td>

<tr>
<td>Adresse:</td>
<td COLSPAN="3"><input NAME="tbxAdress" TYPE="TEXT" SIZE="40" MAXLENGTH="50" Value="<%=strAdress%>"></td>
<tr>
<td>Postnr:</td>
<td><input NAME="tbxPostnr" TYPE="TEXT" SIZE="5" MAXLENGTH="5" Value="<%=strPostnr%>"></td>
<td>Poststed:</td>
<td><input NAME="tbxPoststed" TYPE="TEXT" SIZE="20" MAXLENGTH="50" Value="<%=strPoststed%>"></td>

</table>

<p>

</table>
<p>
<table cellpadding='0' cellspacing='0'>
   <td><INPUT NAME='pbnDataAction' TYPE='SUBMIT'  VALUE='Nullstill'>
   <td><INPUT NAME='pbnDataAction' TYPE='SUBMIT'  VALUE='Lagre'>
<% 
If Request.QueryString("AdrID")<>"" Then
   Response.write "<td><INPUT NAME='pbnDataAction' TYPE=SUBMIT  VALUE='Slette'>"
End If
%>
</table>

</form>
    </div>
</body>
</html>

