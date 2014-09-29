<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<% 

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

If Request.QueryString("FirmaID") <> "" Then
   Set rsFirm  = Conn.Execute("Select FirmaID, Firma from SUSPECT where FirmaID = " & Request.QueryString("FirmaID"))
   strFirmaID = rsFirm("FirmaID")
   strFirm   = rsFirm("Firma")
   
   If Request.QueryString("ContactID") <> "" Then
      Set rsContact  = Conn.Execute("Select * from SUSPECT_KONTAKT where SUSPECT_KONTAKT.KontaktID = " & Request.QueryString("ContactID") )

      strContactId = rsContact("KontaktID")
      strLastName  = rsContact("Etternavn")
      strFirstName = rsContact("Fornavn")
      strBirthDate = rsContact("Foedselsdato")
      strNameday   = rsContact("Navndag")
      lMedID       = rsContact("MedID")
      strMemo      = rsContact("Notat")
      lFagOmrID    = rsContact("fagomraadeID")
      strTelefon    = rsContact("Telefon")
      strFax        = rsContact("Fax") 
      strMobilTlf   = rsContact("MobilTlf") 
      strEPost      = rsContact("EPost") 
      strStilling   = rsContact("Stilling")

   Else
       strContactId = 0
       strLastName  = ""
       strFirstName = ""
       strBirthDate = ""
       strNameday   = ""
       lMedID       = ""
       strMemo      = ""
       lFagOmrID    = 0

  End If
    
   Set rsFagomrade   = Conn.Execute("select * from H_KONTAKT_FAGOMRADE")
      
Else
   Set rsFirm = Nothing
End If

strHeading = "Kontaktperson for " & strFirm 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<title><%=strHeading %></title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<form NAME="KONTAKTP" ACTION="Suspectkontaktpdb.asp" METHOD="POST">
<input TYPE="HIDDEN" NAME="tbxFirmaID" VALUE="<% =strFirmaID %>">
<input TYPE="HIDDEN" NAME="tbxContactID" VALUE="<% =Request.QueryString("ContactID") %>">

<table cellpadding='0' cellspacing='0'>
<h3 ALIGN="CENTER"><%=strHeading %></h3>
<tr>
<td>Fornavn:
<td><input NAME="tbxFirstName" Size="20" TYPE="TEXT" Value="<% =strFirstName %>">
<td>Etternavn:
<td><input NAME="tbxLastName" Size="20" TYPE="TEXT" Value="<% =strLastName %>">
<tr>
<td>Telefon:</td>
<td><input NAME="tbxTelefon" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<% =strTelefon %>"></td>
<td>Mobil:</td>
<td><input NAME="tbxMobilTlf" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<% =strMobilTlf %>"></td>
<tr>
<td>Fax:</td>
<td><input NAME="tbxFax" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<% =strFax %>"></td>
<td>E-Post:</td>
<td><input NAME="tbxEPost" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<% =strEPost %>"></td>
<tr>
<td>Fødselsdato:
<td><input NAME="tbxBirthDate" TYPE="TEXT" SIZE="8" MAXLENGTH="10" Value="<% =strBirthdate %>" ONBLUR="dateCheck(this.form, this.name)">
<td>Navnedag:
<td><input NAME="tbxNameDay" TYPE="TEXT" TYPE="TEXT" SIZE="8" MAXLENGTH="10" Value="<% =strNameDay %>" ONBLUR="dateCheck(this.form, this.name)">
<tr>
<td>Stilling:
<td><input NAME="tbxStilling" TYPE="TEXT" Value="<% =strStilling %>">
<td>Fagområde:
<td><select NAME="dbxFagOmrade">
    <option VALUE="0">
<% 
Do Until rsFagomrade.EOF
   If rsFagomrade("FagomrID") = lFagomrID Then
      strValueSelected = rsFagomrade("FagomrID") & " SELECTED"
   Else
      strValueSelected = rsFagomrade("FagomrID")
   End If  
%>
    <option VALUE="<% =strValueSelected %>"><% =rsFagomrade("Fagomerade") %>
<% 
rsFagomrade.MoveNext
Loop
%>
</select></td>
<tr>
<td>Notat:
<td COLSPAN="3"><textarea NAME="tbxMemo" COLS="50" ROWS="6">
<% =strMemo %>
</textarea>

</table>
<p>

<table cellpadding='0' cellspacing='0'>
<% If Request.QueryString("ContactID")="" Then
   Response.write "<td>"
   Response.write "<td><INPUT NAME=pbDataAction TYPE=SUBMIT  VALUE=Nullstill>"
   Response.write "<td><INPUT NAME=pbDataAction TYPE=SUBMIT  VALUE=Lagre>"
Else
   Response.write "<td><INPUT NAME=pbDataAction TYPE=SUBMIT  VALUE=Nullstill>"
   Response.write "<td><INPUT NAME=pbDataAction TYPE=SUBMIT  VALUE=Lagre>"
   Response.write "<td><INPUT NAME=pbDataAction TYPE=SUBMIT  VALUE=Slette>"
End If
%>
</table>


</form>
    </div>
</body>
</html>

