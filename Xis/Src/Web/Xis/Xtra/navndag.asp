<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/Library.inc"--> 

<% profil = Session("Profil") 


' Is this first time to show this page
If Request.Form( "tbxPageNo") = "" Then
 
   ' Do we have KontaktID ?
   If Request.Querystring( "KontaktID" ) = "" Then
       Response.Write "Systemfeil: KontaktID mangler"
       Response.End
   Else
      KontaktID = Request.Querystring( "KontaktID" )
   End If

   ' Do we have FirmaID ?
   If Request.Querystring( "FirmaID" ) = "" Then
      Response.Write "Systemfeil: FirmaID mangler"
      Response.End
   Else
      FirmaID = Request.Querystring( "FirmaID" )
   End If

   SokNavn = Request.QueryString( "SokNavn" )

Else

   ' Add values from current page    
   KontaktID = Request.Form( "tbxKontaktID" )
   FirmaID    = Request.Form( "tbxFirmaID" )
   SokNavn  = Request.Form( "tbxSokNavn" )

End If

' First time page called and search value exist ?
If    SokNavn <> ""  Then 

   ' Open database connection
   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
   Conn.CommandTimeOut = Session("xtra_CommandTimeout")
   Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

   ' Get all navndager
   Set rsNavndag = Conn.Execute("Select NavndagID, Navn, Dato from H_NAVNDAG where Navn like " & Quote( SokNavn & "%") & " order by navn" )

   ' No records found ?
   If rsNavndag.BOF = True And rsNavndag.EOF = True Then 
      RecordsFound = 0
   Else
      RecordsFound = 1
   End If

Else

   ' No records found
   RecordsFound = 0

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
	<title>Søk navndag</title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<h1>Søk navndag</h1>

<form name="formEn" ACTION="navndag.asp" METHOD="POST">
	<input type="hidden" NAME="tbxPageNo" VALUE="1">
	<input type="hidden" NAME="tbxKontaktID" VALUE="<%=KontaktID%>">
	<input type="hidden" NAME="tbxFirmaID" VALUE="<%=FirmaID%>">
	
	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>Fornavn:</td>
			<td><input name="tbxSokNavn" MAXLENGTH="50" value="<%=SokNavn%>"></td>
			<td><input type="submit" name="pbnDataAction" value="Søk"></td>
		</tr>
	</table>
</form>

<%
' -----------------------------------------------
' Create table only when records found
' -----------------------------------------------

If  RecordsFound = 1  Then

   ' Create table
   Response.Write "<table cellpadding='0' cellspacing='0'>"

   ' Create table heading
   Response.Write "<tr>"
   Response.Write "<th>Velg navn</th>"
   Response.Write "<th>Dato</th>"
   Response.Write "</tr>"

   Do Until rsNavndag.EOF

      ' Create row
      Response.Write "<tr>"
      Response.Write "<td><A href='navndagdb.asp?KontaktID=" & KontaktID & "&FirmaID=" & FirmaID & "&Dato=" & rsNavndag( "Dato") & "'>" & rsNavndag( "Navn") & "</a></td>"
      Response.Write "<td>" & rsNavndag( "Dato") & "</td>"
      Response.Write "</tr>"

       ' Get next record
       rsNavndag.MoveNext
   Loop

   ' Close recordset
   rsNavndag.Close

   ' Clear recordset
   set rsNavndag = Nothing

   ' End table
   Response.Write "</table>"

End If
%>

    </div>
</body>
</html>

