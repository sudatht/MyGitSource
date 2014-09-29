<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Library.inc"-->
<% 

' Do not stop when error occurs
On Error Resume Next

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

If Request("pbDataAction") = "Nullstill" Then 
    Response.redirect "suspectkontaktp.asp?FirmaID=" & Request.form("tbxFirmaID")

ElseIf Request("pbDataAction") = "Lagre" and Request.form("tbxContactID") = "" Then 

  If Request.form("tbxLastName") = "" Then
     Response.Write "Etternavn ikke utfyllt" 
     Response.End
  End If 
     
  strSQL = "Insert into SUSPECT_KONTAKT( FirmaID, EtterNavn, Fornavn, Foedselsdato, Navndag, Notat, fagomraadeID, Stilling, Telefon, Fax, MobilTlf, EPost ) " & _
           "Values( " & Request.form("tbxFirmaID") & "," & _
		   "'" & Request.form("tbxLastName") & "'," & _
		   "'" & Request.form("tbxFirstName") & "'," & _
		   DbDate( Request.form("tbxBirthDate") ) & "," & _
		   DbDate( Request.form("tbxNameDay") ) & "," & _
		   "'" & PadQuotes(Request.form("tbxMemo")) & "'," & _
		   Request.form("dbxfagomrade") & "," &_
		   "'" & Request.form("tbxStilling") & "'," & _
		   "'" & Request.form("tbxTelefon") & "'," & _
		   "'" & Request.form("tbxMobilTlf") & "'," & _
		   "'" & Request.form("tbxFax") & "'," & _
		   "'" & Request.form("tbxEPost") & "')"

   Conn.Execute( strSQL )

   ' Get new key value
   Set rsContactID  = Conn.Execute("Select NewContactID=max(KontaktID) from SUSPECT_KONTAKT where FirmaID = " & Request.form("tbxFirmaID") )

   Response.redirect "suspectvis.asp?FirmaID=" & Request.form("tbxFirmaID") & "&ContactID="&rsContactID("NewContactID")   

ElseIf Request("pbDataAction") = "Lagre" and Request.form("tbxContactID") <> "" Then
  strSQL = "Update SUSPECT_KONTAKT set Etternavn = " & "'" & Request.form("tbxLastName") & "'" & _
                                      ", Fornavn      =" & "'" & Request.form("tbxFirstName") & "'" & _ 
		  ", Foedselsdato =" & DbDate( Request.form("tbxBirthDate") ) & _
		  ", Navndag      =" & DbDate( Request.form("tbxNameDay") ) & _
		  ", Notat        =" & "'" & PadQuotes(Request.form("tbxMemo")) & "'" & _
		  ", fagomraadeID  = " & Request.form("dbxFagomrade") & _
		  ", Stilling     = " & "'" & Request.form("tbxStilling") & "'" & _
		  ", Telefon      = " & "'" & Request.form("tbxTelefon") & "'" & _
		  ", Fax          = " & "'" & Request.form("tbxFax") & "'" & _
		  ", MobilTlf     = " & "'" & Request.form("tbxMobilTlf") & "'" & _
		  ", EPost        = " & "'" & Request.form("tbxEPost") & "'" & _
		  " where kontaktID = " & Request.form("tbxContactID")

  'Response.Write strSql
  Conn.Execute( strSQL )

   Response.redirect "suspectvis.asp?FirmaID=" & Request.form("tbxFirmaID") & "&ContactID=" & Request.form("tbxContactID") 

ElseIf Request("pbDataAction") = "Slette" Then

  strSQL = "Delete SUSPECT_KONTAKT where KontaktID = " & Request.form("tbxContactID")
  Conn.Execute( strSQL )

  Response.redirect "suspectvis.asp?FirmaID=" & Request.form("tbxFirmaID") 

Else
  strSQL = "No action added for this test" 
End If

Response.redirect "suspectvis.asp?FirmaID = " & Request.form("tbxFirmaID")
%>