<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------------------------------------------
' Checking parameters
'--------------------------------------------------------------------------------------------------

If Request.QueryString("VikarID") = "" Then
   strVikarID = Request.Form("VikarID")
   strOppdragID = Request.Form("OppdragID")
   strFirmaID = Request.Form("FirmaID")
   kode = Request.Form("kode")
   tilgang = Request.Form("tilgang")
Else
   strVikarID = Request.QueryString("VikarID")
   strOppdragID = Request.QueryString("OppdragID")
   strFirmaID = Request.QueryString("FirmaID")
   kode = Request.QueryString("kode")
   tilgang = Request.QueryString("tilgang")
End If

'Response.write strVikarId & "<br>"
'Response.write strOppdragId & "<br>"
'Response.write strFirmaID & "<br>"
'Response.write kode & "<br>"
'Response.write tilgang & "<br>"

'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
' Get data
'--------------------------------------------------------------------------------------------------
conn.Execute("Set rowcount 0")

   Set rsPost        = Conn.Execute("select * from H_POSTNUMMER order by sted")


   Set rsVikar  = Conn.Execute("select VikarID, Etternavn, Fornavn, " &_
				"Foedselsdato, Personnummer , Bankkontonr, " &_
				"Kommunenr, Skattetabellnr, Skatteprosent, Ansattdato, " &_
				"Loenn1, TypeID, Notat " & _
                               	"from VIKAR " &_
				"where VikarID = " & Request.QueryString("VikarID") )


   Set rsAdress = Conn.Execute("select AdrId, AdresseType, Adresse, Postnr, Poststed " &_
				"from ADRESSE where " &_
				"AdresseRelID = " & Request.QueryString("VikarID") &_
				"and ADRESSE.AdresseType = 1 " &_
				"and ADRESSE.Adresserelasjon = 2 ")

'--------------------------------------------------------------------------------------------------
' Put into variables
'--------------------------------------------------------------------------------------------------
   strVikarID   = rsVikar("VikarID")
   strEtternavn = rsVikar("Etternavn")
   strFornavn   = rsVikar("Fornavn")
   strFoedselsdato = rsVikar("Foedselsdato")
   strPersonnummer = rsVikar("Personnummer")

   strAnsattdato = rsVikar("Ansattdato")
   strBankkontonr = rsVikar("Bankkontonr")
   strKommunenr = rsVikar("Kommunenr")
   strSkattetab = rsVikar("Skattetabellnr")
   strSkatteprosent = rsVikar("Skatteprosent")

   strNotat = rsVikar("Notat")
   strTimeloenn = rsVikar("Loenn1")

If Not rsAdress.EOF and Not rsAdress.BOF Then
   lAdrID = rsAdress("AdrID")

   ' Set adress values
   strAdress = rsAdress("Adresse")
Else
   lAdrID = 2
   strAdress = ""
End IF
'Response.write rsVikar("TypeID")
   If rsVikar("TypeID") = 3 Then
       strChecked = "CHECKED"
   End If

'   strHeading = "Endre opplysninger for vikar " & strFornavn & " " & strEtternavn


'--------------------------------------------------------------------------------------------------
' Headder
'--------------------------------------------------------------------------------------------------
%>

<html>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
	<link rel="stylesheet" href="/xtra/css/
main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/
print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<title><%=strHeading %></title>
	<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
	<script LANGUAGE="VBSCRIPT">
		Function DateCheck( iType )

		     If Cint(iType) = 1 Then
		        strDate = document.vikar.tbxfoedselsdato.value
		     ElseIf Cint(iType) = 2 Then
		        strDate = document.vikar.tbxIntervjudato.value
			 End If
		     If strDate <> "" Then
			    If Not IsDate( strDate ) Then
			       Alert("Ulovlig dato. Lovlig format er DD.MM.YY")
		          If Cint(iType) = 1 Then
		             strDate = document.vikar.tbxfoedselsdato.focus
		          ElseIf Cint(iType) = 2 Then
		             strDate = document.vikar.tbxIntervjudato.focus
			      End If
		    	End If
			End If

		End Function
</script>
</head>
<body>
	<div class="pageContainer" id="pageContainer">

<h1><%=strHeading %></h1>

<form Name="VIKAR" ACTION="Vikardb2.asp?VikarID=<% =strVikarID %>&amp;kode=<% =kode %>&amp;tilgang=<% =tilgang %>&amp;OppdragID=<% =strOppdragID %>&amp;FirmaID=<% =strFirmaID %>" METHOD="POST" TARGET="_parent">

<input name="tbxVikarID" TYPE="HIDDEN" VALUE="<%=strVikarID%>">
<input name="tbxStatus" TYPE="HIDDEN" VALUE="<%=strStatus%>">
<input name="tbxAdrID" TYPE="HIDDEN" Value="<%=lAdrID%>">
<input name="tbxAdrType" TYPE="HIDDEN" Value="1">


<table cellpadding='0' cellspacing='0'>
<tr>
  <td ALIGN="RIGHT">Fornavn:
  <td><input name="tbxFornavn" TYPE="TEXT" SIZE="30" MAXLENGTH="50" VALUE="<%=strFornavn%>"></td>
<tr>
  <td ALIGN="RIGHT">Etternavn:
  <td><input name="tbxEtternavn" TYPE="TEXT" SIZE="30" MAXLENGTH="50" VALUE="<%=strEtternavn %>"></td>
<tr>
  <td ALIGN="RIGHT">Fødselnr:
  <td><nobr><input name="tbxFoedselsdato" TYPE="TEXT" SIZE="8" MAXLENGTH="10" VALUE="<%=strFoedselsdato %>" ONBLUR="dateCheck(this.form, this.name)">
     <input name="tbxPersonnummer" TYPE="TEXT" SIZE="4" MAXLENGTH="5" VALUE="<%=strPersonnummer %>">
     </td>
<tr>
  <td ALIGN="RIGHT">Adresse:
  <td><input name="tbxAdresse" TYPE="TEXT" SIZE="30" MAXLENGTH="50" VALUE="<%=strAdress %>"></td>
<tr>
  <td ALIGN="RIGHT">Poststed:</td>

  <td><select NAME="dbxPostNo">
    	<option VALUE="0">
	<% Do Until rsPost.EOF
   		strPost = rsPost("PostNr") & " " & rsPost("Sted")
   		If rsAdress("Postnr") = rsPost("Postnr") Then
      			strSelected = rsPost("PostnrID") & " " & "SELECTED"
   		Else
      			strSelected = rsPost("PostnrID")
  	 	End If	%>

		<option VALUE="<%=strSelected %>"><%=strPost %>
	<%rsPost.MoveNext
	  Loop
	  rsPost.Close %>
    </select>

<tr>
  <td ALIGN="RIGHT">AvdelingID:
  <td><input name="tbxAvdelingid" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<%=strAvdelingId%>">
<tr>
  <td ALIGN="RIGHT">Ansattdato:
  <td><input name="tbxAnsattdato" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<%=strAnsattdato%>" ONBLUR="dateCheck(this.form, this.name)">
<tr>
  <td ALIGN="RIGHT">Bankkontonr:
  <td><input name="tbxBankkontonr" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<%=strBankkontonr%>">
<tr>
  <td ALIGN="RIGHT">Kommunenr:
<td><input name="tbxKommunenr" TYPE="TEXT" SIZE="7" VALUE="<% =strKommunenr %>">
<tr>
  <td ALIGN="RIGHT">Timelønn:
  <td><input name="tbxTimeloenn" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<%=strTimeloenn%>">
<tr>
  <td ALIGN="RIGHT">Skattetabellnr:
  <td><input name="tbxSkattetabellnr" TYPE="TEXT" SIZE="15" MAXLENGTH="50" VALUE="<%=strSkattetab%>">
<tr>
  <td ALIGN="RIGHT">Skatteprosent:
  <td><input name="tbxSkatteprosent" TYPE="TEXT" SIZE="15" MAXLENGTH="40" VALUE="<%=strSkatteprosent%>">
<tr>
  <td ALIGN="RIGHT">A/S:
  <td><input name="cbxAS" TYPE="CHECKBOX" <%=strChecked %>>
<tr>
<td ALIGN="Right">Notat:</td>
<td><textarea NAME="tbxNotat" ROWS="4" COLS="25"><%=strNotat%></textarea></td>
<tr>
  <td colspan="2"><input name="pbnDataAction" TYPE="SUBMIT" VALUE="                                Lagre                         "></td>
</form>

</table>

</body>
</html>