<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>

<% 
profil = Session("Profil")

' Continue on Error
On Error Resume Next

' Check data
' --------------------
' Check FirmaID

If Request.Querystring("FirmaID") <> "" Then
   lFirmaID = CLng( Request.Querystring("FirmaID") )
Else
   Response.write "Parameter missing"
   Response.End
End If

' Connect database
' --------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

' Get Firmadata
' ----------------------------

' Get firmadata
strSql = "Select FirmaID, Firma, OrgNr, Status, Kategori, Bransje, Etternavn, Fornavn, kreditgrense," &_
                        "telefon, fax, EPost, Hjemmeside, A.Adresse, A.Postnr, A.Poststed, kredittopplysning " &_
             "from SUSPECT F, H_FIRMA_BRANSJE B, H_FIRMA_KATEGORI K, H_FIRMA_STATUS S, MEDARBEIDER M, suspect_ADRESSE A " &_
                            "where FirmaID = " & lFirmaID &_
                            " and F.BransjeID *= B.BransjeID " &_
                            " and F.KategoriID *= K.KategoriID  " &_
                            " and F.StatusID *= S.StatusID " &_
                            " and F.AnsvMedID *= M.MedID " &_
                            " and F.FirmaID = A.adresseRelID " & _
                            " and A.AdresseRelasjon = 1 and A.AdresseType = 1 "

Set rsFirma  = Conn.Execute( strSQL )

' Create poststed
strPoststed = rsFirma("Postnr") & " " & rsFirma("postSted")

' Get all connected addresses
Set rsAdresse = Conn.Execute("Select A.AdrId, A.Adresse , T.AdrtypeID, T.AdresseType, A.Postnr, A.Poststed from SUSPECT_ADRESSE A, H_ADRESSE_TYPE T where A.adresseRelID = " & lFirmaID & " and A.AdresseRelasjon = 1 and A.adressetype > 1 and A.AdresseType = T.AdrTypeID" )

' Get all connected Kontaktpersoner
Set rsKontaktp= Conn.Execute("Select KontaktId, Fornavn, Etternavn, Stilling, Telefon, MobilTlf, Fax, EPost from SUSPECT_KONTAKT where SUSPECT_KONTAKT.FirmaID = " & lFirmaID)
 
strName = rsFirma("Fornavn") & " " & rsFirma("Etternavn")

' Set heading in page
strHeading="Suspect " & rsFirma("Firma") 

' set parametre for hotlist
strFirma = rsFirma("Firma") 
strFirmaID = Request.Querystring("FirmaID")
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
</head>
<body>
	<div class="pageContainer" id="pageContainer">

<h1><%=strHeading %></h1>
<table cellpadding='0' cellspacing='0'>
<tr>
   <th>Suspect nr:</th>
   <td COLSPAN="2"><%=rsFirma("FirmaID") %></td>
   <td></td>
</tr>
<tr>
   <th>Navn:</th>
   <td><%=rsFirma("Firma") %></td>
   <th>OrgNo:</th>
   <td><%=rsFirma("OrgNr") %></td>
</tr>
<tr>
   <th>Telefon:</th>
   <td><%=rsFirma("Telefon") %></td>
   <th>Fax:</th>
   <td><%=rsFirma("Fax") %></td>
</tr>
<tr>
   <th>E-Post:</th>
   <td><a HREF="mailto:<%=rsFirma("EPost") %>"><%=rsFirma("EPost") %></td>
   <th>Hjemmeside:</th>
   <td><a HREF="http://<%=rsFirma("Hjemmeside") %>" TARGET="_NEW"><%=rsFirma("Hjemmeside") %></td>
</tr>
<tr>
   <th>Besøksadresse:</th>
   <td><%=rsFirma("Adresse") %></td>
   <th>Poststed:</th>
   <td><%=strPoststed %></td>
   
</tr>
<tr>
   <th>Vikar</th>
   <td><%=strName %></td>
   <th>Bransje:</th>
   <td><%=rsFirma("Bransje") %></td>
</tr>
<tr>
   <th>Kategori:</th>
   <td><%=rsFirma("Kategori") %></td>
   <th>Status:</th>
   <td><%=rsFirma("Status") %></td>
</tr>   
<tr>
   <th>Kredittgrense:</th>
   <td COLSPAN="3"><%=rsFirma("Kreditgrense") %></td>
</tr>
<tr>
   <th>Kredittoppl:</th>
   <td COLSPAN="3"><%=rsFirma("KredittOpplysning") %></td>
</tr>
</table>
<table cellpadding='0' cellspacing='0'>
<tr>
<% If  Mid(profil,2,1) > 1 Then %>
<td>
<form ACTION="suspectny.asp?FirmaID=<%=lFirmaID %>" METHOD="POST">
   <input NAME="tbxFirmaID" TYPE="HIDDEN" VALUE="<%=lFirmaID %>">
   <input NAME="pbnDataAction" TYPE="SUBMIT" VALUE="Endre Suspect">
</form>
</td>
<td>
<form ACTION="suspect_prospect.asp" METHOD="POST">
   <input NAME="FirmaID" TYPE="HIDDEN" VALUE="<%=lFirmaID %>">
   <input NAME="TypeID" TYPE="HIDDEN" VALUE="1">
   <input NAME="pbnDataAction" TYPE="SUBMIT" VALUE="Overfør til prospect">
</form>
</td>

<% End If %>
<td>
</td>
</tr>
</table>
<form ACTION="suspectadresse.asp?Relasjon=1&amp;ID=<%=rsFirma("FirmaId") %>" METHOD="POST">
 <table BORDER="1">

   <% Do Until rsAdresse.EOF
   strFullAdresse = rsAdresse("Adresse") & " " & rsAdresse("PostNr") & " " & rsAdresse("PostSted") %>
 <tr>
   <th><%=rsAdresse("AdresseType") %> </th>
<td>
<% If  Mid(profil,2,1) > 1 Then %>
   <a Href="suspectadresse.asp?Relasjon=1&amp;ID=<%=rsFirma("FirmaId") %>&amp;AdrID=<%=rsAdresse("AdrId") %>"><%= strFullAdresse %> </a> 

<% Else  %>
 <%= strFullAdresse %>  
 <% End If %>  

</td>
<% rsAdresse.MoveNext

   Loop %>
 </table>
<% If  Mid(profil,2,1) > 1 Then %>
 <input NAME="pbDataAction" TYPE="SUBMIT" VALUE="Ny adresse">
<% End If %>
</form>

<form ACTION="suspectkontaktp.asp?FirmaID=<%=rsFirma("FirmaId") %>" METHOD="POST">
 <table BORDER="1">
   <tr>
   <th>Kontaktperson<th>Stilling<th>Telefon<th>Fax<th>MobilTlf<th>E-Post
<% 
Do Until rsKontaktP.EOF
   strName = rsKontaktP("Fornavn") & " " & rsKontaktP("Etternavn")
%>
<tr>
<td>
<% If  Mid(profil,2,1) > 1 Then %>
<a Href="suspectkontaktp.asp?FirmaID=<%=rsFirma("FirmaID") %>&amp;ContactID=<%=rsKontaktP("KontaktId") %>"><%=strName %> </a>

<% Else %>
<%=strName %>
<% End If %> 

</td>
<td><%=rsKontaktP("Stilling") %></td>
<td><%=rsKontaktP("Telefon") %></td>
<td><%=rsKontaktP("Fax") %></td>
<td><%=rsKontaktP("MobilTlf") %></td>
<td><a HREF="mailto:<%=rsKontaktP("EPost") %>"><%=rsKontaktP("EPost")%></td>
<%
   rsKontaktP.MoveNext
Loop
%>
 </table>
<% If  Mid(profil,2,1) > 1 Then %>
 <input NAME="pbDataAction" TYPE="SUBMIT" VALUE="Ny kontakt">
<% End If %>
</form>
    </div>
</body>
</html>

