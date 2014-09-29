<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\Library.inc"-->

<%

' Continue on error
On Error Resume Next

' Open database
' ------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' Clear all fields ?
' -------------------
If Request("pbDataAction") = "Nullstill" Then
	Response.redirect "suspectny.asp"
' Save new  ?
ElseIf Request("pbDataAction") = "Lagre" and Request.form("tbxFirmaID") = "" Then
   ' Adresse must exist
   If Trim( Request.Form("tbxAdresse") ) = "" Then
       Response.Write "Adresse mangler!"
       Response.End
   End If

   ' Start transaction
   '   Conn.BeginTrans

   ' Insert subject
   ' -----------------------

   ' Build sql-string
   strSQL = "Insert into suspect( Firma, StatusID, BransjeID, OrgNr, Telefon, Fax, Hjemmeside, LagtinnNaar ) " & _
           "Values( " & Quote( Request.form("tbxFirm") ) & "," &_
		   "1" & "," &_
		   Request.form("dbxBransje") & "," &_
		   Quote( Request.form("tbxOrgNo") ) &"," &_
		   Quote( Request.form("tbxTelefon") ) & "," & _
		   Quote( Request.form("tbxFax") ) & "," & _
		   Quote( Request.form("tbxHjemmeside") ) & "," &_
		   "Getdate()" &  ")"

   ' Insert in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans

      Call SqlError()
   End if

   ' Get new FirmaID
   Set rsFirma  = Conn.Execute("Select NewFirmaID=max(FirmaID) from suspect")

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
'      Conn.RollBackTrans

      Call SqlError()
   End if

   ' Insert ADRESSE
   ' -----------------------

   ' Build SQL_statement
   strSQL = "Insert into suspect_ADRESSE(AdresseRelID, Adresse, Postnr, Poststed, AdresseType, AdresseRelasjon ) " & _
                  "Values( " & rsFirma("NewFirmaID") & "," & _
		   Quote( Request.form("tbxadresse") ) & "," & _
		   Quote( Request.form("tbxPostNr") )& "," & _
		   Quote( Request.form("tbxPoststed") )& "," & _
		   "1" & "," &_
		   "1" & " )"

   ' Insert in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
 '     Conn.RollBackTrans

      Call SqlError()
   End if

   ' Insert KONTAKTPERSON
   ' ------------------------------

   ' Build SQL-statement
   strSQL = "Insert into SUSPECT_KONTAKT( FirmaID, EtterNavn, Fornavn, Notat, Telefon, Fax, EPost ) " & _
                   "Values( " & rsFirma("NewFirmaID") & "," & _
		   "'" & Request.form("tbxEtternavn") & "'," & _
		   "'" & Request.form("tbxFornavn") & "'," & _
		   "'" & Request.form("tbxMemo") & "'," & _
		   "'" & Request.form("tbxKTelefon") & "'," & _
		   "'" & Request.form("tbxKFax") & "'," & _
		   "'" & Request.form("tbxKEPost") & "')"

   ' Insert in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      '  Conn.RollBackTrans

      Call SqlError()
   End if

   ' Commit transaction
   ' Conn.CommitTrans

   ' Redirect page
   Response.redirect "suspectvis.asp?FirmaID=" & rsFirma("NewFirmaID")

' Save  ?
' -------------------
ElseIf Request("pbDataAction") = "Lagre" and Request.form("tbxFirmaID") <> "" Then

   ' Update SUSPECT
   ' -----------------------------

   ' Build sql -statement
   strSQL = "Update suspect set firma = " & "'" & Request.form("tbxFirm") & "'" & _
                                      ", OrgNr = " & Quote( Request.form("tbxOrgNo") ) & _
		  ", Endring = 1"  & _
		  ", BransjeID =" & Request.form("dbxBransje") & _
		  ", Telefon      = " & Quote( Request.form("tbxTelefon") ) & _
		  ", Fax          = " & Quote( Request.form("tbxFax") ) & _
		  ", Hjemmeside   = " & Quote( Request.form("tbxHjemmeside") )& _
                                      ", EndretNaar   = " & "GetDate()" &_
		  " where firmaid = " & Request.form("tbxFirmaID")

   ' Update in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then
		'Rollback transaction
		'Conn.RollBackTrans
		Call SqlError()
   End if

   ' Update ADRESSE
   strSQL = "Update suspect_Adresse set " &_
                                      " Adresse      = " & Quote( Request.form("tbxAdresse") ) & "," & _
                                      " Postnr = " & Quote( Request.form("tbxPostNr") )& "," & _
                                      " Poststed = " & Quote( Request.form("tbxPoststed") ) & _
                                      " where Adrid =" & Request.form("tbxAdrID")

   ' Update in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
'      Conn.RollBackTrans

      Call SqlError()
   End if

   ' Update KONTAKTPERSON ??

   ' Build SQL statement
   strSQL = "Update SUSPECT_KONTAKT set Etternavn = " & "'" & Request.form("tbxEtternavn") & "'" & _
                                      ", Fornavn      =" & "'" & Request.form("tbxFornavn") & "'" & _
		  ", Notat        =" & "'" & Request.form("tbxMemo") & "'" & _
		  ", Telefon      = " & "'" & Request.form("tbxKTelefon") & "'" & _
		  ", Fax          = " & "'" & Request.form("tbxKFax") & "'" & _
		  ", EPost        = " & "'" & Request.form("tbxKEPost") & "'" & _
		  " where KontaktID = " & Request.form("tbxKontaktID")

   ' Update in database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
'      Conn.RollBackTrans

      Call SqlError()
   End if

   Response.redirect "suspectvis.asp?FirmaID=" &  Request.form("tbxFirmaID")

' Delete  ?
ElseIf Request("pbDataAction") = "Slette" Then

   ' Delete from suspect
  strSQL = "Delete suspect where firmaid = " & Request.form("tbxFirmaID")
  Conn.Execute( strSQL )

   ' Delete from suspect_adresse
  strSQL = "Delete suspect_adresse where adresserelid = " & Request.form("tbxFirmaID") & " and adresserelasjon=1"
  Conn.Execute( strSQL )

   ' Delete from suspect_Kontakt
  strSQL = "Delete suspect_kontakt where firmaid = " & Request.form("tbxFirmaID")
  Conn.Execute( strSQL )

  lFirmaID = ""

Else
  strSQL = "No action added for this test"
End If
Response.redirect "suspectvis.asp?FirmaID=" &  Request.form("tbxFirmaID")

%>
