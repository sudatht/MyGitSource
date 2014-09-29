<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/Library.inc"-->

<%
profil = Session("Profil")

' Is this first time to show this page
If Request.Form( "tbxPageNo") = "" Then

Else

   ' Add values from current page
   Fradato     = Request.Form( "tbxFradato" )
   Tildato      = Request.Form( "tbxTildato" )

End If

' First time page called and search value exist ?
If Fradato <> "" And Tildato <> ""  Then

	if (ValidDateInterval(ToDateFromDDMMYY(Fradato), ToDateFromDDMMYY(Tildato)) = false) then
		Response.write "<p class'warning'>Fradato kan ikke være senere enn tildato!</p>"
		Response.End
	end if

   ' Open database connection
   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
   Conn.CommandTimeOut = Session("xtra_CommandTimeout")
   Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

   ' Get all
   strSql = "Select F.FirmaID, F.Firma, Fornavn, Etternavn, Navndag " &_
                "From KONTAKT K, FIRMA F " &_
                "Where  DATEPART (month, navndag) >= " & Month( Fradato )&_
                " And   DATEPART (month, navndag) <= "  & Month( Tildato ) &_
                " And  K.FirmaID = F.FirmaID" &_
                " Order by  DATEPART (month, navndag), DATEPART (day, navndag)"

  ' Response.write strSQL
   Set rsRapport = Conn.Execute( strSql )

   ' No records found ?
   If rsRapport.BOF = True And rsRapport.EOF = True Then
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
	<script language="javascript" src="Js/javascript.js" type="text/javascript"></script>
	<title>Navndag-kontaktpersoner</title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<h1>Navndag-kontaktpersoner</h1>
<form name="formEn" ACTION="navnedager.asp" METHOD="POST">
	<input type="hidden" NAME="tbxPageNo" VALUE="1">

	<table cellpadding='0' cellspacing='0'>
		<tr>
			<td>Fra dato:</td>
			<td><input NAME="tbxFraDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Fradato%>" ONBLUR="dateCheck(this.form, this.name)"> </td>
			<td>Til dato:</td>
			<td><input NAME="tbxTilDato" TYPE="TEXT" SIZE="10" MAXLENGTH="10" Value="<%=Tildato%>" ONBLUR="dateCheck(this.form, this.name), dateInterval(this.form, this.name)"> </td>
			<td><input type="submit" name="pbnDataAction" value="     Søk    "></td>
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
   Response.Write "<th>Navnedag</th>"
   Response.Write "<th>Kontaktperson</th>"
   Response.Write "<th>Kontakt</th>"
   Response.Write "</tr>"

   ' Create new fromdate
   FromMonth = Month( Fradato )
   FromDay = Day( Fradato )

   If FromDay < 10 Then
      From = FromMonth & 0 & FromDay
   Else
      From = FromMonth & FromDay
   End If

   Do Until rsRapport.EOF

      If ( Month( rsRapport("navndag") ) = Month( fradato ) And Day( rsRapport("navndag") ) >= Day ( fradato ) ) or ( Month( rsRapport("navndag") ) = Month( tildato ) And Day( rsRapport("navndag") ) <= Day ( tildato )  )  or ( Month( rsRapport("navndag")) > Month( fradato )  and  Month( rsRapport("navndag")) < Month( Tildato ) ) Then

          ' Create row
         Response.Write "<tr>"
         Response.Write "<TD class=right>" & Mid(rsRapport("navndag"), 1,5) & "</td>"
         Response.Write "<td>" & rsRapport( "Fornavn") & " " & rsRapport( "Etternavn") & "</td>"
         Response.Write "<td><A href='kundevis.asp?FirmaID=" & rsRapport("FirmaID") & "'>" & rsRapport("Firma") & "</a></td>"

         Response.Write "</tr>"
      End If

       ' Get next record
       rsRapport.MoveNext
   Loop

   ' Close recordset
   rsRapport.Close

   ' Clear recordset
   set rsRapport = Nothing

   ' End table
   Response.Write "</table>"

End If
%>

    </div>
</body>
</html>

