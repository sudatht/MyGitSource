<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>

<!--#INCLUDE FILE="includes/Library.inc"--> 
<% 
' Continue on error
On Error Resume Next

' Check input values
' --------------------------

' TypeID of move  (SuspectID = 1or Prospect = 2) ?
If Request.Form( "TypeID" ) <> "" And ( Request.Form( "TypeID" ) = "1" Or Request.Form( "TypeID" ) = "2" ) Then
   TypeID = Request.Form( "TypeID" )
Else
   Response.Write "Systemfeil: TypeID mangler"
   Response.End
End If

' ID of move
If Request.Form( "FirmaID" ) = "" Then
   Response.Write "Systemfeil: FirmaID mangler"
   Response.End
Else
   ID = Request.Form( "FirmaID" )
End If

' Connect to database 
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' Is this suspect to prospect ?

If TypeID = 1 Then

   strFromFirmaTable = "Suspect"
   strToFirmaTable = "Firma"
   FirmaStatus = 1

   strFromAdresseTable = "Suspect_adresse"
   strToAdresseTable = "Adresse"

   strFromKontaktTable = "Suspect_kontakt"
   strToKontaktTable = "Kontakt"

Else

   strFromFirmaTable = "Firma"
   strToFirmaTable = "Suspect"
   FirmaStatus = 0

   strFromAdresseTable = "Adresse"
   strToAdresseTable = "Suspect_adresse"

   strFromKontaktTable = "Kontakt"
   strToKontaktTable = "Suspect_kontakt"

End If


   ' ---------------------------------
   ' Copy from suspect to firma
   ' ---------------------------------

   ' Build sql-statement
   strSql = "Insert into " & strToFirmaTable & "( Firma, StatusID, BransjeID, KategoriID, AnsvMedID, OrgNr, Kreditgrense" &_
                   ", Telefon, Fax, EPost, Hjemmeside, Kredittopplysning, LagtinnNaar, LagtInnMedID ) " & _
                  "Select Firma, " & FirmaStatus &", BransjeID, KategoriID, AnsvMedID, OrgNr, Kreditgrense" &_
                   ", Telefon, Fax, EPost, Hjemmeside, Kredittopplysning, LagtinnNaar, LagtInnMedID  " & _
                   " from " & strFromFirmaTable & " where FirmaID = " & ID

    ' Start transaction
    '  Conn.BeginTrans

   ' Insert into database
   Conn.Execute( strSql )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' ---------------------
   ' Get new FirmaID
   ' --------------------
   Set rsFirma  = Conn.Execute("Select NewFirmaID=max(FirmaID) from " & strToFirmaTable )

   ' Database error ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
'      Conn.RollBackTrans 

      Call SqlError()
   End if

   ' ----------------------------------------
   ' Copy from suspect_adresse to adresse
   ' ---------------------------------------
 
  ' Build Sql-statement   
   strSQL = "Insert into " & strToAdresseTable & "(AdresseRelID, Adresse, Postnr, Poststed, AdresseType, AdresseRelasjon ) " & _
                   " Select " & rsFirma("NewFirmaID") & ", Adresse, Postnr, Poststed, AdresseType, AdresseRelasjon " &_
                   " from " & strFromAdresseTable & " where AdresserelID = " & ID

   ' Insert into database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' ----------------------------------------
   ' Copy from suspect_kontakt to kontakt
   ' ---------------------------------------
 
  ' Build Sql-statement   
  strSQL = "Insert into " & strToKontaktTable & "( FirmaID, EtterNavn, Fornavn, Foedselsdato, Navndag, Notat, fagomraadeID, Stilling, Telefon, Fax, MobilTlf, EPost ) " &_
                  "select  " & rsFirma("NewFirmaID") & ", EtterNavn, Fornavn, Foedselsdato, Navndag, Notat, fagomraadeID, Stilling, Telefon, Fax, MobilTlf, EPost " &_
                  " from " & strFromKontaktTable & " where FirmaID = " & ID

   ' Insert into database
   Conn.Execute( strSQL )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' ---------------------------------
   ' Delete Firma
   ' ---------------------------------

   ' Build sql-statement
   strSql = "Delete from " & strFromFirmaTable & " where FirmaID = " & ID

    ' Start transaction
    '  Conn.BeginTrans

   ' Delete in database
   Conn.Execute( strSql )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' ---------------------------------
   ' Delete Adresse
   ' ---------------------------------

   ' Build sql-statement
   strSql = "delete from " & strFromAdresseTable & " where adresserelid = " & ID

    ' Start transaction
    '  Conn.BeginTrans

   ' delete in database
   Conn.Execute( strSql )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' ---------------------------------
   ' Delete Kontakt
   ' ---------------------------------

   ' Build sql-statement
   strSql = "delete from " & strFromKontaktTable & " where firmaid = " & ID

    ' Start transaction
    '  Conn.BeginTrans

   ' Delete in database
   Conn.Execute( strSql )

   ' Error from database ?
   If Conn.Errors.Count > 0 then

      ' Rollback transaction
      ' Conn.RollBackTrans 

      ' Error message
      Call SqlError()
   End if

   ' Commit transaction
'   Conn.CommitTrans

   Response.redirect "kundevis.asp?FirmaID=" & rsFirma("NewFirmaID")

%>
