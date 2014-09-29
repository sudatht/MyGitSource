<!--#INCLUDE FILE="../includes/Library.inc"-->
<html>
<head>
	<title>Overføring av personopplysninger</title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<h1>Eksport til Hult &amp; Lillevik.</h1>

<h2>Var overføringen vellykket?</h2>

<FORM ACTION=Vikar_overf_ok.asp?kode=<% =Request.QueryString("kode") %> METHOD=POST>
	<INPUT TYPE=SUBMIT VALUE="JA">
</form>

<FORM ACTION="Hult_Lill_03.asp?kode=<% =Request.QueryString("kode") %>" TARGET=_top METHOD=POST>
	<INPUT TYPE=SUBMIT VALUE="Nei">
</form>

<%
'--------------------------------------------------------------------------------------------------
' Modifications:
' Date - Who - What
' 23.01.2000 - Arne Leithe - GetAvdeling is called to calculate correct avdeling (why isn't it in the db, anyway?)
'--------------------------------------------------------------------------------------------------

Response.write Request.QueryString("kode")

'--------------------------------------------------------------------------------------------------
' Connect to database
'--------------------------------------------------------------------------------------------------

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
Conn.CommandTimeOut = Session("xtra_CommandTimeout")
Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

'--------------------------------------------------------------------------------------------------
'      Hjelpe funksjoner
'--------------------------------------------------------------------------------------------------

Function StrOfSign(aStrLength, aSign)
'Returns a string with aSign with length aStrLength
	lStr = ""
	For I=1 to aStrLength
	   lStr = lStr & aSign
	next
	StrOfSign = lStr
end Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ExpNum2Str(aNum, aFormat)
'Returns a number as a string using the format in aFormat (Left filled numbers with 0)

    If IsNull(aNum) Then
      aNum = 0
    End If

    If Len(aNum) = 0 Then
      aNum = 0
    End If

'First check format
    lPos = InStr(1, aFormat, ".")

    If (lPos > 0) Then  'It exist a fractional
        lLenFormatDecimals = Len(aFormat) - lPos
        lLenFormatNumber = Len(aFormat) - lLenFormatDecimals - 1    '1 is komma
    Else
        lLenFormatDecimals = 0
        lLenFormatNumber = Len(aFormat)
    End If



    'Check inputnumber. We first want to concern about positiv number
    lInput = Abs(aNum)

    If lLenFormatDecimals > 0 Then    'We have want a fractional number

        'Check input number
        lLen = Len(lInput)                'If there is a comma in the number, exchange this to dash
        lPos = InStr(1, lInput, ",")

        'Make fraction
        If lPos > 0 Then
            lInputFraction = Mid(lInput, lPos + 1, lLen - lPos)
            lInputFraction = lInputFraction & StrOfSign(lLenFormatDecimals - Len(lInputFraction), "0")
        Else    'The input umber does not contain a fraction
            lInputFraction = StrOfSign(lLenFormatDecimals, "0")
        End If

        'Make number
        If lPos > 0 Then
            lInputNumber = Left(lInput, lPos - 1)
        Else
            lInputNumber = lInput
        End If

        If Len(lInputNumber) > lLenFormatNumber Then
            MsgBox "Overflow !"
        Else
            lNumber = StrOfSign((lLenFormatNumber - Len(lInputNumber)), "0") & lInputNumber
        End If

        lReturnNumber = lNumber & "." & lInputFraction

    Else    'We have a whole number
        If Len(aNum) > lLenFormatNumber Then
            MsgBox "Overflow !"
            Exit Function
        Else
            lReturnNumber = StrOfSign((lLenFormatNumber - Len(lInput)), "0") & lInput
        End If
    End If


   'Bruk minustegn på første plass i formatet (hvis negativt tall)

    If Not IsNumeric(aNum) Then         'Eks. -00.12   (Med punktum er dette en string)
        If Left(aNum, 1) = "-" Then
            lReturnNumber = "-" & Right(lReturnNumber, Len(lReturnNumber) - 1)
        End If
    Else
        If aNum < 0 Then
          lReturnNumber = "-" & Right(lReturnNumber, Len(lReturnNumber) - 1)
        End If
    End If
'101
    ExpNum2Str = lReturnNumber

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RightPad(aStr, aLen)
'Fill out right side of aStr with spaces, such that total string length is aLen
    If IsNull(aStr) Then
        lLenStr = 0
        aStr = ""
    Else
        lLenStr = Len(aStr)
    End If

    If lLenStr > aLen Then
        ' Only take number of chars that are allowed
        lStr = Left(aStr, aLen)
    Else
        lStr = aStr & Space(aLen - lLenStr)
    End If
    RightPad = lStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function RightFill(aStr, aLen, asign)
	If IsNull(aStr) Then
	     aStr = ""
	End If
	If Len(aStr) > aLen Then
	     aStr = Left(aStr, aLen)
	End If
	ss = aStr

	For i=Len(aStr) To aLen - 1
	    If asign = "_" Then
		ss = ss & " "
	    Else
		ss = ss & asign
	    End If
	Next
	RightFill = ss
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function LeftFill(aStr, aLen, asign)
	If IsNull(aStr) Then
	     aStr = ""
	End If
	If Len(aStr) > aLen Then
	     aStr = Left(aStr, aLen)
	End If
	ss = aStr

	For i=Len(aStr) To aLen - 1
	    If asign = "_" Then
		ss = " " & ss
	    Else
		ss = asign & ss
	    End If
	Next
	LeftFill = ss
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function OnlyDigits( strString )
' Remove all non-nummeric signs from string
  If Not IsNull(strString) Then
   For idx = 1 To Len( strString) Step 1
      Digit = Asc( Mid( strString, idx, 1 ) )
      If (( Digit > 47 ) And ( Digit < 58 )) Then
         strNewstring = strNewString & Mid(strString, idx,1)
      End If
   Next
  End If
   Onlydigits = strNewString
End Function

Function Konverter(strInn)
	For i = 1 To Len(strInn)
		nr = Asc(Mid(strInn, i,1))
		Select Case nr
			Case 198
		strMidl = Left(strInn, i - 1) + Chr(146) + Mid(strInn, i + 1)
		strInn = strMidl
			Case 216
		strMidl = Left(strInn, i - 1) + Chr(157) + Mid(strInn, i + 1)
		strInn = strMidl
			Case 197
		strMidl = Left(strInn, i - 1) + Chr(143) + Mid(strInn, i + 1)
		strInn = strMidl
			Case 230
		strMidl = Left(strInn, i - 1) + Chr(145) + Mid(strInn, i + 1)
		strInn = strMidl
			Case 248
		strMidl = Left(strInn, i - 1) + Chr(155) + Mid(strInn, i + 1)
		strInn = strMidl
			Case 229
		strMidl = Left(strInn, i - 1) + Chr(134) + Mid(strInn, i + 1)
		strInn = strMidl
			Case Else
		End Select
	Next
	Konverter = strInn
End Function

%>
<H4>Eksport av personopplysninger til Hult & Lillevik.  </H4>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''
''''''''''''''''''''	PERSONOPPLYSNINGER   -  NYE OG ENDREDE VIKARER
''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rsVikar = Conn.Execute("SELECT " & _
							"VIKAR.VikarId, " & _
							"VIKAR.Etternavn, " & _
							"VIKAR.Fornavn, " & _
							"VIKAR.foedselsdato, " & _
							"VIKAR.Personnummer, " & _
							"VIKAR.Bankkontonr, " & _
							"VIKAR.Kommunenr, " & _
							"VIKAR.Skattetabellnr, " & _
							"VIKAR.Skatteprosent, " & _
							"VIKAR.AnsattDato, " & _
							"TimL = VIKAR.Loenn1 * 100, " & _
							"VIKAR_ANSATTNUMMER.ansattnummer, " & _
							"ADRESSE.Adresse, " & _
							"ADRESSE.Postnr, " & _
							"ADRESSE.Poststed " & _
							"FROM VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, ADRESSE " & _
							"WHERE VIKAR.VikarID = ADRESSE.AdresseRelID " & _
							"AND ADRESSE.AdresseRelasjon = '2' AND ADRESSE.AdresseType = '1' " & _
							"AND VIKAR.overfort = '1' " & _
							"ORDER BY VIKAR.Etternavn ")

' No records found
If rsVikar.BOF = True And rsVikar.EOF = True Then
   Response.write "<H4>Ingen rader i tabellen</H4>"
Else
%>


<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Ansattnummer</th>
		<th>Navn</th>
		<th>Adresse</th>
		<th>Post</th>
		<th>FDato</th>
		<th>Bank</th>
		<th>Kommune</th>
		<th>Tabell</th>
		<th>Prosent</th>
		<th>Beg.dato</th>
		<th>Timesats</th>
	</tr>
<%
Do Until rsVikar.EOF

   strFullName   = konverter(rsVikar("Etternavn").Value) & " " & konverter(rsVikar("Fornavn").Value)
   strPostAddress = rsVikar("Postnr").Value & " " & konverter(rsVikar("PostSted").Value)
	fNr = rsVikar("foedselsdato").Value  & " " & rsVikar("personnummer").Value

%>

	<tr>
		<td><%=rsVikar("ansattnummer").Value%></td>
		<td><%=strFullName%></td>
		<td><%=rsVikar("Adresse").Value%></td>
		<td><%=strPostAddress%></td>
		<td><%=fNr %></td>
		<td><%=rsVikar("Bankkontonr").Value%></td>
		<td><%=rsVikar("Kommunenr").Value%></td>
		<td><%=rsVikar("Skattetabellnr").Value%>4</td>
		<td><%=rsVikar("Skatteprosent").Value%></td>
		<td><%=rsVikar("Ansattdato").Value%></td>
		<td><%=rsVikar("TimL").Value%></td>
	</tr>

<%
   rsVikar.MoveNext
Loop
%>
</table>


<p>----------------------------- TEKSTFIL FOR VIKAROPPLYSNINGER -------------------------------------</p>
<%
   Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
   'TestFile = Server.MapPath ("/Cosmossystems/xtradev") & "\dataf\H002opl.txt"
   User = Request.ServerVariables("LOGON_USER")
   pos = Instr(User,"\")
   User = Mid(User, pos + 1)
   nn = OnlyDigits(Now)
   nn = Left(nn,10)
   name = "\H02_" & nn & "_" & User & ".txt"

	'"rubicon_test" is the test area for HL / rubicon files, "rubicon" is the production area
	BackupFile = Application("RubiconFileRoot") & Left(session("avdkontornavn"),3) & "\" & name
	TestFile = Application("RubiconFileRoot") & Left(session("avdkontornavn"),3) & "\H02.txt"

   Set OutStream= FileObject.CreateTextFile (TestFile, True, False)
   Set BkStream = FileObject.CreateTextFile (BackupFile, True, False)

rsVikar.MoveFirst
Do Until rsVikar.EOF

	strFullName   = konverter(rsVikar("Etternavn").Value) & " " & konverter(rsVikar("Fornavn").Value)
	strPostAddress = rsVikar("Postnr").Value & " " & konverter(rsVikar("PostSted").Value)
	fNr =  OnlyDigits(rsVikar("foedselsdato").Value)  & " " & rsVikar("personnummer").Value

	sql = "Select AvdelingID From VIKAR_AVDELING where VikarID = " & rsVikar("VikarID").Value
	set rsAvd = conn.execute(sql)
	If Not rsAvd.EOF Then
		'temporary method for finding correct avdeling, replaces the silly formula below with some silly if-statements.
		'Avdeling = (rsAvd("AvdelingID")+ 5) * 100
		Avdeling = GetAvdeling(Session("avdkontornavn"), rsAvd("AvdelingID").Value)
	Else
		Avdeling = ""
	End If
	rsAvd.Close: Set rsAvd = Nothing

'--------------------------------------------------------------------------------------------------
'       Varabel = innhold			    feltlengde     kommentar
'--------------------------------------------------------------------------------------------------
LoennsNr = ExpNum2Str(rsVikar("ansattnummer").Value,"0000")	'4
Navn = RightFill(strFullName,24,"_")			'24
Adresse = RightFill(rsVikar("Adresse").Value,24,"_")		'24
Sted = RightFill(strPostAddress,24,"_") 		'24
FNr = LeftFill(fNr,12,"0")				'(0)12 (med eller uten skilletegn)
BankkontoNr = LeftFill(rsVikar("Bankkontonr").Value,13,"_")	'(0)13 (. = blank eller ingen skilletegn)
KommuneNr = LeftFill(rsVikar("Kommunenr").Value,4,"_")	'(0)4
SkattetabellNr = LeftFill(rsVikar("Skattetabellnr").Value,4,"_")	'(0)4
Skatteprosent = LeftFill(OnlyDigits(rsVikar("Skatteprosent").Value),2,"_")	'(0)2
BegDato = LeftFill(OnlyDigits(rsVikar("Ansattdato").Value),6,"_")		'(0)6 (ddmmaaaa)
AvdelingsNr = LeftFill(Avdeling,6,"_")			'(0)6
If rsVikar("timL") > 0 Then
	Timesats = LeftFill(rsVikar("TimL").Value,5,"_")		'(0)5
Else
	Timesats = LeftFill("",5,"_")		'(0)5
End If
'Slutt = Chr(13) & Chr(10)	'CRLF

'--------------------------------------------------------------------------------------------------
'      lager record streng
'--------------------------------------------------------------------------------------------------

RecStr = ""
RecStr = RecStr & LoennsNr
'Response.write LoennsNr & " 1<br>"
RecStr = RecStr & Navn
'Response.write Navn & " 2<br>"
RecStr = RecStr & Adresse
'Response.write Adresse & " 3<br>"
RecStr = RecStr & Sted
'Response.write Sted & " 4<br>"
RecStr = RecStr & FNr
'Response.write FNr & " 5<br>"
RecStr = RecStr & BankkontoNr
'Response.write BankkontoNr & " 6<br>"
RecStr = RecStr & KommuneNr
'Response.write KommuneNr & " 7<br>"
RecStr = RecStr & SkattetabellNr
'Response.write SkattetabellNr & " 8<br>"
RecStr = RecStr & Skatteprosent
'Response.write Skatteprosent & " 9<br>"
RecStr = RecStr & BegDato
'Response.write Begdato & " 10<br>"
RecStr = RecStr & AvdelingsNr
'Response.write AvdelingsNr & " 11<br>"
RecStr = RecStr & Timesats
'Response.write Timesats & " 12<br>"
RecStr = RecStr & Slutt


   Response.Write RecStr
    Response.write "<br>"

   OutStream.WriteLine RecStr
   BkStream.WriteLine RecStr

   'OutStream.WriteLine Chr(13)

rsVikar.MoveNext
Loop

   Set OutStream = Nothing
   Set BkStream = Nothing
End If 'ingen rader i tabellen
   rsVikar.Close

Response.end
%>
<h2>Eksport av faste lønnsopplysninger til Hult &amp; Lillevik.</h2>
<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''
''''''''''''''''''''	FASTE LØNNSOPPLYSNINGER   -  NYE OG ENDREDE VIKARER
''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rsVikar = Conn.Execute("SELECT " & _
							"VIKAR.VikarId, " & _
							"VIKAR_ANSATTNUMMER.ansattnummer, " & _
							"VIKAR_LOENN_FASTE.Avdeling, " & _
							"VIKAR_LOENN_FASTE.Loennsart, " & _
							"Ant = VIKAR_LOENN_FASTE.Antall * 100, " & _
							"Sat = VIKAR_LOENN_FASTE.Sats * 100, " & _
							"Bel = VIKAR_LOENN_FASTE.Beloep * 100, " & _
							"Sal = VIKAR_LOENN_FASTE.Saldo * 100 " & _
							"FROM VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, VIKAR_LOENN_FASTE " & _
							"WHERE VIKAR.VikarID = VIKAR_LOENN_FASTE.VikarID " & _
							"AND VIKAR.overfort = '1' ")

Set rsVik = Conn.Execute("SELECT " & _
							"VIKAR.VikarId, " & _
							"VIKAR_ANSATTNUMMER.ansattnummer, " & _
							"VIKAR_LOENN_FASTE.Avdeling, " & _
							"VIKAR_LOENN_FASTE.Loennsart, " & _
							"Ant = VIKAR_LOENN_FASTE.Antall * 100, " & _
							"Sat = VIKAR_LOENN_FASTE.Sats * 100, " & _
							"Bel = VIKAR_LOENN_FASTE.Beloep * 100, " & _
							"Sal = VIKAR_LOENN_FASTE.Saldo * 100 " & _
							"FROM VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, VIKAR_LOENN_FASTE " & _
							"WHERE VIKAR.VikarID = VIKAR_LOENN_FASTE.VikarID " & _
							"AND VIKAR.overfort = '1' ")

' No records found
If rsVikar.BOF = True And rsVikar.EOF = True Then
   Response.write "<H4>Ingen rader i tabellen</H4>"
Else
%>


<table cellpadding='0' cellspacing='0'>
	<tr>
		<th>Ansattnummer</th>
		<th>Avd</th>
		<th>Lønnsart</th>
		<th>Antall</th>
		<th>Sats</th>
		<th>Beløp</th>
		<th>Saldo</th>
	</tr>
<%
Do Until rsVikar.EOF
%>
	<tr>
		<th><%=rsVikar("ansattnummer").Value%></th>
		<th><%=rsVikar("Avdeling").Value%></th>
		<th><%=rsVikar("Loennsart").Value%></th>
		<th><%=rsVikar("Ant").Value%></th>
		<th><%=rsVikar("Sat").Value%></th>
		<th><%=rsVikar("Bel").Value%></th>
		<th><%=rsVikar("Sal").Value%></th>
	</tr>
<%
   rsVikar.MoveNext
Loop
%>
</table>


<p>----------------------------- TEKSTFIL FOR FASTE LØNNSOPPLYSNINGER --------------------------------</p>
<%
   Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
   'TestFile = Server.MapPath ("/Cosmossystems/xtradev") & "\dataf\H003opl.txt"
   User = Request.ServerVariables("LOGON_USER")
   pos = Instr(User,"\")
   User = Mid(User, pos + 1)
   nn = OnlyDigits(Now)
   nn = Left(nn,10)
   name = "\H03_" & nn & "_" & User & ".txt"
'Response.write name & "<br>"
   'BackupFile = Server.MapPath ("/Cosmossystems/xtradev") & name

   BackupFile = "t:\loenn" & name
   TestFile = "t:\loenn\H03.txt"

   Set OutStream= FileObject.CreateTextFile (TestFile, True, False)
   Set BkStream = FileObject.CreateTextFile (BackupFile, True, False)

   rsVikar.MoveFirst
 Do Until rsVikar.EOF


'--------------------------------------------------------------------------------------------------
'       Varabel = innhold			    feltlengde     kommentar
'--------------------------------------------------------------------------------------------------
LoennsNr = ExpNum2Str(rsVikar("ansattnummer").Value,"0000")	'4
Avdeling = LeftFill(rsVikar("Avdeling").Value,8,"0")		'8
Loennsart = LeftFill(rsVikar("Loennsart").Value,3,"0")	'3
Ledig1 = "      "					'6
Ledig2 = "      "					'6
Antall = LeftFill(rsVikar("Ant").Value,10,"0")		'10 (100-deler)
Sats = LeftFill(rsVikar("Sat").Value,10,"0")			'10 (ører)

If rsVikar("Bel").Value < 0 Then
	FortegnBeloep = "-"
Else
	FortegnBeloep = "+"
End If							'1

Beloep = LeftFill(OnlyDigits(rsVikar("Bel").Value),9,"0")	'9 (ører)

If rsVikar("Sal").Value < 0 Then
	FortegnSaldo = "-"
Else
	FortegnSaldo = "+"
End If							'1

Saldo = LeftFill(rsVikar("Sal").Value,9,"0")			'9
Ledig3 = "                                                             " '61
'Slutt = Chr(13) & Chr(10)	'CRLF

'--------------------------------------------------------------------------------------------------
'      lager record streng
'--------------------------------------------------------------------------------------------------

RecStr = ""
RecStr = RecStr & LoennsNr
RecStr = RecStr & Avdeling
RecStr = RecStr & Loennsart
RecStr = RecStr & Ledig1
RecStr = RecStr & Ledig2
RecStr = RecStr & Antall
RecStr = RecStr & Sats
RecStr = RecStr & FortegnBeloep
RecStr = RecStr & Beloep
RecStr = RecStr & FortegnSaldo
RecStr = RecStr & Saldo
RecStr = RecStr & Ledig3
RecStr = RecStr & Slutt


   Response.Write RecStr
	Response.write "<br>"
   OutStream.WriteLine RecStr
   BkStream.WriteLine RecStr

   'OutStream.WriteLine Chr(10)

rsVikar.MoveNext
Loop

   Set OutStream = Nothing
   Set BkStream = Nothing

End If 'ingen rader i tabellen

   rsVikar.Close

%>
    </div>
</body>
</html>

