<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/Library.inc"-->
<%
Function Modulus_11_Check( aNr )
' This function checks a string after MOD 11
'(If CheckSum = 10 erstattes 10 med et minustegn "-")
    
    lNrStr = Mid( aNr, 1, Len(aNr) - 1 )
	lLenNr = Len(lNrStr)

    lProdukt = 0
	lKontrol = 0

    For I = 0 To lLenNr - 1
        lProdukt = lProdukt + (Mid(lNrStr, lLenNr - I, 1) * ((I Mod 6) + 2))
    Next
    
    lKontrol = 11 - ( lProdukt Mod 11)
    
    If lKontrol = 10 Then
        Modulus_11_ControlNr = 0
    Else
        Modulus_11_ControlNr = lKontrol
    End If

	' check if correct
	If Modulus_11_ControlNr = CInt(Right(aNr,1)) Then
	   Modulus_11_Check = True
	Else 
	   Modulus_11_Check = False
    End If

End Function


Sub OnlyDigits( strString )

' Remove all non-nummeric signs from string

   For idx = 1 To Len( strString) Step 1
      Digit = Asc( Mid( strString, idx, 1 ) ) 
      If (( Digit > 47 ) And ( Digit < 58 )) Then 
         strNewstring = strNewString & Mid(strString, idx,1)
      End If
   Next
   strString = StrNewString
End Sub

' Do not stop when error occurs
'On Error Resume Next

' **********************************************
' Check datavalues 
' **********************************************

If Request.Form("tbxEtternavn") = "" Then
   Response.write "Etternavn ikke utfyllt"
   Response.End
End If

If Request.Form("tbxAdresse") = "" Then
   Response.write "Adresse ikke utfyllt"
   Response.End
End If

If Request.Form("tbxFoedselsdato") <> "" Then
   If Not IsDate( Request.Form("tbxFoedselsdato") ) Then
      Response.write "Fødselsdato: Ulovlig dato"
      Response.End
	Else
	  ' Response.Write CDate(Request.Form("tbxFoedselsdato"))
   End If
End If
 
' Get Fødselnr
strDate = Request.Form("tbxFoedselsdato")

If Request.Form("tbxPersonnummer") <> "" and Request.Form("tbxPersonnummer") <> "0" Then

   ' Remove seperator
   Call OnlyDigits( strDate )

   ' Create Fødselsnummer
   strNumber =  strDate & Request.Form("tbxPersonnummer")

   ' Check Fødselnummer
   'If Not Modulus_11_Check( strNumber ) Then
   '   Response.write "FødselsNummer: Ulovlig verdi1" & strNumber
   '   Response.End
   'End If
End If

' Timeloen must have value
If Request.Form("tbxTimeloenn") = "" Then
   lTimeloenn = 0
Else
   lTimeloenn = Request.Form("tbxTimeloenn")
End IF

' Personnummer must have value
If Request.Form("tbxPersonnummer") = "" Then
   lPersonnummer = 0
Else
   lPersonnummer = Request.Form("tbxPersonnummer")
End IF

' Check faktura
If Request.Form("cbxAS") = "on" Then
   lSelvstendig = 3
Else
   lSelvstendig = 1
End If

' must have value
If Request.Form("tbxKommunenr") = "" Then
   lKommunenr = "Null"
Else
   lKommunenr = "'" & Request.Form("tbxKommunenr") & "'"
End IF

If Request.Form("tbxBankkontonr") = "" Then
   lBankkontonr = "Null"
Else
   lBankkontonr = "'" & Request.Form("tbxBankkontonr") & "'"
End IF

If Request.Form("tbxSkatteprosent") = "" Then
   lSkatteprosent = "Null"
Else
   lSkatteprosent = Request.Form("tbxSkatteprosent")
End IF

If Request.Form("tbxSkattetabellnr") = "" Then
   lSkattetabellnr = "Null"
Else
   lSkattetabellnr = "'" & Request.Form("tbxSkattetabellnr") & "'"
End IF

If Request.Form("tbxAvdelingid") = "" Then
   lAvdelingid = "Null"
Else
   lAvdelingid = Request.Form("tbxAvdelingid")
End IF

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
'Response.write kode & "<br>"
'Response.write tilgang & "<br>"

'Response.write Request.form("tbxFornavn") & " fn<br>"
'Response.write Request.form("tbxEtternavn") & " en<br>"
'Response.write Request.form("tbxFoedselsdato") & " fd<br>"
'Response.write Request.form("tbxAnsattdato") & " ad<br>"
'Response.write Request.form("tbxKommunenr") & " knr<br>"
'Response.write Request.form("tbxSkattetabellnr") & " sk<br>"
'Response.write Request.form("tbxNotat") & " not<br>"
'Response.write Request.form("tbxFoedselsdato") & " fd<br>"
'Response.write Request.form("tbxPersonnummer") & " pnr<br>"
'Response.write Request.form("tbxSkatteprosent") & " sp<br>"
'Response.write Request.form("tbxBankkontonr") & " bknr<br>"
'Response.write Request.form("dbxPostno") & " pidnr<br>"
'Response.write kode


' ****************************************************
' Open database connection
' ****************************************************
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("Xtra_CommandTimeout")
Conn.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")

' ****************************************************
' Update  Vikar
' ****************************************************

   ' Update Vikar in database
   strSQL = "Update Vikar set " &_
          	" Fornavn      = " & "'" & Request.form("tbxFornavn") & "'" &_
		", Etternavn    = " & "'" & Request.form("tbxEtternavn") & "'" &_
          	", Foedselsdato = " & DbDate(Request.form("tbxFoedselsdato") ) &_
		", Ansattdato   = " & DbDate( Request.form("tbxAnsattdato") ) &_
		", Kommunenr    = " & lKommunenr &_
		", Bankkontonr  = " & lBankkontonr &_
		", Skattetabellnr = " & lSkattetabellnr &_
		", Skatteprosent = " & lSkatteprosent &_
		", Notat        = " & "'" & Request.form("tbxNotat") & "'" & _
		", PersonNummer = " & lPersonnummer & _
		", Loenn1    = " & lTimeloenn & _
		", TypeID  = " & lSelvstendig & _
		", Endring  = 1" & _
		" where Vikarid =" & Request.form("tbxVikarID")

'Response.write strSQL & "<br>"
'Response.write Request.Form("cbxAS")
'Response.end

   Conn.Execute( strSQL )
   IF Conn.Errors.Count > 0 then
      Call SqlError()
   End if

' ****************************************************
' Update  Adress
' ****************************************************

  strSQL = "select postnr, sted from H_POSTNUMMER where PostnrID = " & Request.Form("dbxPostNo")
  
'Response.write Request.Form("dbxPostNo") & "<br>"

Set rsADR = conn.Execute(strSQL)
'Response.write strSQL & "<br>"
'Response.write rsADR("Postnr")

if not rsADR.EOF then
	strSQL = "Update ADRESSE set" &_
       	" Adresse = " & "'" & Request.form("tbxAdresse") & "'" & _
	 	", Postnr = '" & rsADR("Postnr") &_
	 	"', Poststed = '" & rsADR("Sted") &_
	 	"', Adressetype = 1 " &_
 	 	"where AdrId  = " & Request.form("tbxAdrID")
else 
   strSQL = "Update ADRESSE set" &_
          	" Adresse = " & "'" & Request.form("tbxAdresse") & "'" & _
			", Adressetype = 1 " &_
 			"where AdrId  = " & Request.form("tbxAdrID")
end if


'Response.write strSQL & "<br>"
   Conn.Execute(strSQL)

'   IF Conn.Errors.Count > 0 then
'      Call SqlError()
'   End if

rsADR.Close
Set rsADR = Nothing

   ' Set return value
   strVikarID = Request.form("tbxVikarID")


redir = "Vikar_frames.asp?VikarID=" & strVikarID & "&kode=" & kode & "&tilgang=" & tilgang & "&OppdragID=" & strOppdragID & "&FirmaID=" & strFirmaID
'Response.write redir

Response.redirect redir

 %>
