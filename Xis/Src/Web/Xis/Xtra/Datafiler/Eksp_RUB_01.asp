<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<%

	If (HasUserRight(ACCESS_ADMIN, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Conn
	dim strSQL
	dim valgt_avd
	Dim strAvdelingNr
	dim rsOrdre
	dim util

	Dim kundeFile		'As String
	Dim kundeStream		'As TextFile
	Dim strKundeSQL		'As String
	Dim rsKundeData		'As Recordset

	if Request.QueryString("avd") <> "" then
		valgt_avd = CInt(Request.QueryString("avd"))
	else
		valgt_avd = 0
	end if
		
	Function StrOfSign(aStrLength, aSign)
		lStr = ""
		For I=1 to aStrLength
		lStr = lStr & aSign
		next
		StrOfSign = lStr
	end Function

	Function ExpNum2Str(aNum, aFormat)
	'This function returns a number as a string using the format in aFormat (Left filled numbers with 0)

		If IsNull(aNum) Then
			aNum = ""
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

	Function RightPad(aStr, aLen)
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

	Function OnlyDigits( strString )
		' Remove all non-nummeric signs FROM string
		If Not IsNull(strString) Then
			For idx = 1 To Len( strString) Step 1
				Digit = Asc( Mid( strString, idx, 1 ) )
				If (( Digit > 47 ) AND ( Digit < 58 )) Then
					strNewstring = strNewString & Mid(strString, idx,1)
				End If
			Next
		End If
		Onlydigits = strNewString
	End Function

	Function Konverter(strInn)
		For i = 1 To Len(strInn)
			nr = Asc(Mid(strInn, i,1))
			SELECT Case nr
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

<html>
	<head>
		<title>Overføring til Rubicon</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript">
		function sjekkPeriode(form, sjekkfelt, url)
		{

			verdi = parseInt(document.forms(form).elements(sjekkfelt).value);
			if ((verdi >= 1) && (verdi <= 12))
			{
				if (janei = confirm("Bekreft importen for periode " + verdi + "?")){
					(url.indexOf("?") > 1)? delim = "&" : delim = "?";
					document.forms(form).action = url + delim + "periode=" + verdi;
					document.forms(form).submit();
				}
				else
				{
					document.forms(form).elements(sjekkfelt).focus();
				}
			}
			else
			{
				alert("Ikke riktig periode tall");
			}
		}
		</SCRIPT>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="content">
				<h1>Eksport av ordre til Rubicon</h1>
				<h2>Var overføringen vellykket?</h2>
				<FORM NAME="JA" id="JA" ACTION="Faktura_overf_ordre_ok.asp?viskode=2&avd=<%=valgt_avd%>" METHOD="POST">
					<p>Periode: <input type="text" NAME="periode" ID="Text1"></p>
					<input name="jaknapp" TYPE=BUTTON VALUE="JA" onClick="sjekkPeriode(form.name, form.elements(0).name, form.action)" ID="Button1">
				</form>
				<FORM ACTION="faktura_list.asp?viskode=2&dato1=<% =session("limitdato") %>&velgAvdeling=<%=valgt_avd%>" TARGET="_top" METHOD="POST" ID="Form2">
					<INPUT TYPE=SUBMIT VALUE="Nei" ID="Submit1" NAME="Submit1">
				</form>
				<%
				' Open database connection
				Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

				' Get orders (Head)
				if Request.QueryString("avd") <> "" then
					valgt_avd = CInt(Request.QueryString("avd"))
				else
					valgt_avd = 0
				end if

				if valgt_Avd > 0 then
					strSQL = "SELECT * FROM EKSPORT_RUB_ORDRE WHERE Status = 1 AND Avdeling = " & valgt_Avd
				else
					strSQL = "SELECT * FROM EKSPORT_RUB_ORDRE WHERE Status = 1"
				end if

				set rsOrdre = GetFirehoseRS(strSQL, conn)
				' No records found
				If not HasRows(rsOrdre) Then
					set rsOrdre = nothing
					AddErrorMessage("ingen forekomster funnet i rubikon eksport tabell!")
					call RenderErrorMessage()
				End If

				'----------------------------- GENERERING AV TEKSTFIL-------------------------------------

				'''''''' Tekstfilen består av 2 deler:
				''''''''	1. Alle ordrehoder
				''''''''	2. Alle ordrelinjer
				User = Request.ServerVariables("LOGON_USER")
				pos = Instr(User,"\")
				User = Mid(User, pos + 1)
				nn = OnlyDigits(Now)
				nn = Left(nn,10)
				name = "C201_" & nn & "_" & User & ".txt"
				filnavn = "C201ordu.csv"

				Set util = Server.CreateObject("XisSystem.Util")				
				call util.Logon()

				Set FileObject = Server.CreateObject("Scripting.FileSystemObject")

				TestFile = Application("RubiconXtraFileRoot") & filnavn
				BackupFile = Application("RubiconXtraFileRoot") & name

				Set OutStream = FileObject.CreateTextFile (TestFile, True, False)
				Set BkStream = FileObject.CreateTextFile (BackupFile, True, False)

				Do Until rsOrdre.EOF

					KundeNavn = Konverter(rsOrdre("Kundenavn"))
					AdresseI = ""
					AdresseII = ""
					DeresRef = Konverter(rsOrdre("DeresRef"))
					VaarRef = Konverter(rsOrdre("VaarRef"))

					strSQL = "SELECT [avdNR] FROM [avdeling] AS [a] WHERE [a].[avdelingID] = " & rsOrdre("avdeling")
					set rsAvdeling = GetFirehoseRS(strSQL, Conn)
					
					If HasRows(rsAvdeling) then
						strAvdelingNr = rsAvdeling("avdNR")
						rsAvdeling.close
					Else
						strAvdelingNr = 0
					End if
					Set rsAvdeling = nothing

					'Kommenterte ut henting av MVAkode for ordrehode -
					'skal alltid være "01" for innenlands fakturering
					'De individuelle linjene overstyrer ordrehode
   					'ExpNum2Str(rsOrdre("MVAKode"),"00") & ",""" &_

					OutputString = ExpNum2Str(rsOrdre("Formatnummer"),"0000") & "," &_
   		  					ExpNum2Str(rsOrdre("Klientnummer"),"000") & "," &_
							ExpNum2Str(rsOrdre("Registernummer"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("FrBruk1"),"000000") & "," &_
							ExpNum2Str(rsOrdre("FrBruk2"),"0000") & "," &_
							ExpNum2Str(rsOrdre("OrdreType"),"000") & ",""" &_
							RightPad(rsOrdre("Ordrenummer"),7) & """," &_
   		  					ExpNum2Str(rsOrdre("OrdreDato"),"000000") & "," &_
   		  					ExpNum2Str(rsOrdre("Kundenummer"),"000000") & ",""" &_
							RightPad(Kundenavn,30) & """,""" &_
							RightPad(rsOrdre("Alfasort"),15) & """,""" &_
							RightPad(AdresseI,30) & """,""" &_
							RightPad(AdresseII,30) & """,""" &_
   		  					RightPad(rsOrdre("PostNummer"),6) & """,""" &_
							RightPad(rsOrdre("PostSted"),25) & """,""" &_
   		  					RightPad(rsOrdre("LevNavn"),30) & """,""" &_
							RightPad(rsOrdre("LevAdresseI"),30) & """,""" &_
    						RightPad(rsOrdre("LevAdresseII"),30) & """,""" &_
							RightPad(rsOrdre("LevPostNummer"),6) & """,""" &_
							RightPad(rsOrdre("LevPostSted"),25) & """," &_
   		  					ExpNum2Str(rsOrdre("KjedeNummer"),"000000") & "," &_
   		  					ExpNum2Str(rsOrdre("KjedeType"),"00") & "," &_
   		  					ExpNum2Str(rsOrdre("Kundegruppe"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("Medarbeidernummer"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("Rabattgruppe"),"00") & "," &_
   		  					ExpNum2Str(rsOrdre("Betalingsbetingelse"),"00") & "," &_
   		  					ExpNum2Str(rsOrdre("Priskode"),"00") & "," &_
   		  					ExpNum2Str(rsOrdre("Kontantrabatt"),"00.00") & "," &_
   		  					ExpNum2Str(rsOrdre("Ordrerabatt"),"00.00") & ",""" &_
   		  					RightPad(DeresRef,30) & """,""" &_
							RightPad(VaarRef,30) & """,""" &_
   		  					RightPad(rsOrdre("Levbet"),30) & """,""" &_
							RightPad(rsOrdre("LevMaate"),30) & """," &_
   		  					ExpNum2Str(rsOrdre("Porto"),"0000000.00") & "," &_
   		  					ExpNum2Str(rsOrdre("EkspGebyr"),"0000000.00") & "," &_
   		  					ExpNum2Str(rsOrdre("ValPorto"),"0000000.00") & "," &_
   		  					ExpNum2Str(rsOrdre("ValEkspGebyr"),"0000000.00") & "," &_
   		  					"01,""" &_
							RightPad(rsOrdre("FaktDatoType"),1) & """," &_
   		  					ExpNum2Str(rsOrdre("BehProfil"),"00") & ",""" &_
							RightPad(rsOrdre("Factoring"),1) & """," &_
   		  					ExpNum2Str(rsOrdre("Distrikt"),"0000") & "," &_
   		  					ExpNum2Str(rsOrdre("SumOrdre"),"00000000.00") & "," &_
   		  					ExpNum2Str(rsOrdre("LandKode"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("Motbilag"),"000000") & "," &_
   		  					ExpNum2Str(rsOrdre("SprkKode"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("ValutaKode"),"000") & "," &_
   		  					ExpNum2Str(rsOrdre("ValutaKurs"),"00000000.00") & ",""" &_
							RightPad(rsOrdre("Autogiro"),1) & """," &_
   		  					ExpNum2Str(rsOrdre("LevDato"),"000000") & "," &_
   		  					ExpNum2Str(strAvdelingNr,"0000") & "," &_
   		  					ExpNum2Str(rsOrdre("Prosjekt"),"00000") & "," &_
   		  					ExpNum2Str(rsOrdre("Produkt"),"0000") & "," &_
							ExpNum2Str(rsOrdre("HovedbokKtoGr"),"00") & "," &_
							ExpNum2Str(rsOrdre("OppfolgDato"),"000000") & "," &_
							ExpNum2Str(rsOrdre("Innsalgsprosent"),"00.00") & "," &_
   		  					ExpNum2Str(rsOrdre("Lagerstedskode"),"0000") & "," &_
   		  					ExpNum2Str(rsOrdre("HovedbokKtoNr"),"000000") & "," &_
   		  					ExpNum2Str(rsOrdre("FrBruk3"),"0000") & "," &_
							ExpNum2Str(rsOrdre("FrBruk4"),"0000") & "," &_
   		  					ExpNum2Str(rsOrdre("FrBruk5"),"000000") & "," &_
							ExpNum2Str(rsOrdre("FrBruk6"),"000000")

					Response.Write "<p><strong>HODE Faktnr: </strong>" & rsOrdre("Fakturanr") & "</p>"
					Response.Write OutputString & "<br>"

					OutStream.WriteLine OutputString
					BkStream.WriteLine OutputString

					'Legg til ordrelinjer i eksportfil
					strSQL = "SELECT * FROM [EKSPORT_RUB_ORDRELINJE] WHERE [Status] = 1 AND [Fakturanr] = " & rsOrdre("Fakturanr") & " ORDER BY [ID]"

					set rsLinje = GetFirehoseRS(strSQL, Conn)
					If not HasRows(rsLinje) Then
						rsOrdre.close
						set rsOrdre = nothing
						set rsLinje = nothing
						CloseConnection(Conn)
						set Conn = nothing
						AddErrorMessage("ingen rubikon ordre linjer funnet!")
						call RenderErrorMessage()
					End If
					
					Response.Write "LINJER" & "<br>"

					Do Until rsLinje.EOF
						ArtikkelNavn = Konverter(rsLinje("ArtikkelNavn"))
						strSQL = "SELECT [avdnr] FROM [avdeling] AS [a] WHERE [a].[avdelingID] = " & rsOrdre("avdeling")
						Set rsAvdeling = GetFirehoseRS(strSQL, Conn)
						If HasRows(rsAvdeling) Then
							strAvdelingNr = rsAvdeling("avdNr")
							rsAvdeling.close
						Else
							strAvdelingNr = 0
						End if
						Set rsAvdeling = nothing

						StrArtikkelnr = rsLinje("ArtikkelNummer")
						if instr(StrArtikkelnr, "-") then
							StrArtikkelnr = mid(StrArtikkelnr, 1, instr(StrArtikkelnr, "-") - 1)
						end if
						'Sjekk om artikkelnr er av ny sort (301001 osv),  dersom tilfelle
						'slå opp tjenesteområde og hent momskoden, dersom gammel type
						'ikke moms på artikkelen.
						if mid(StrArtikkelnr, 1, 4)="3010" then
							intTOM = clng(clng(StrArtikkelnr) - 301000)
	 						strSQL = "SELECT [rub_Mvakode] FROM [tjenesteomrade] WHERE [tomid] = " & intTom
	 						Set rsTOMKode =  GetFirehoseRS(strSQL, Conn)
	 						strMva = rsTOMKode("rub_Mvakode")
	 						rsTOMKode.close
	 						set rsTOMKode = nothing
						else
							strMva = "00"
						end if

						OutputString = ExpNum2Str(rsLinje("Formatnummer"),"0000") & "," &_
   		  						ExpNum2Str(rsLinje("Klientnummer"),"000") & "," &_
								ExpNum2Str(rsLinje("Registernummer"),"000") & "," &_
   		  						ExpNum2Str(rsLinje("FrBruk1"),"000000") & "," &_
								ExpNum2Str(rsLinje("FrBruk2"),"0000") & ",""" &_
 		  						RightPad(rsLinje("ArtikkelNummer"),24) & """,""" &_
								RightPad(ArtikkelNavn,40) & """,""" &_
   		  						RightPad(rsLinje("AltArtNummer"),24) & """,""" &_
								RightPad(rsLinje("EANArtikkelNummer"),24) & """,""" &_
								RightPad(rsLinje("Plassering"),10) & """," &_
   		  						ExpNum2Str(rsLinje("EnhetsAntall"),"000000.000") & ",""" &_
								RightPad(rsLinje("Enhet"),5) & """," &_
   		  						ExpNum2Str(rsLinje("LevrerandorKto"),"000000") & "," &_
   		  						ExpNum2Str(rsLinje("HovedGruppe"),"000") & "," &_
   		  						ExpNum2Str(rsLinje("UnderGruppe"),"0000") & "," &_
								ExpNum2Str(rsLinje("HovedbokKtoNr"),"000000") & "," &_
								strMva & "," &_
								ExpNum2Str(rsLinje("Lagerstedskode"),"0000") & "," &_
								ExpNum2Str(rsLinje("LevDato"),"000000") & "," &_
								ExpNum2Str(strAvdelingNr,"0000") & "," &_
   		  						ExpNum2Str(rsLinje("Prosjekt"),"00000") & "," &_
   		  						ExpNum2Str(rsLinje("Produkt"),"0000") & "," &_
								ExpNum2Str(rsLinje("IOrdre"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("Levert"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("IRest"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("DelLevert"),"00000000.00") & "," &_
 		  						ExpNum2Str(rsLinje("LinjeRabatt1"),"00.00") & "," &_
   		  						ExpNum2Str(rsLinje("LinjeRabatt2"),"00.00") & "," &_
   		  						ExpNum2Str(rsLinje("InnPris"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("LevPris"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("UtPris1"),"00000000.00") & "," &_
  		  						ExpNum2Str(rsLinje("Utpris2"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("UtPris3"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("Utpris4"),"00000000.00") & "," &_
   		  						ExpNum2Str(rsLinje("UtPris5"),"00000000.00") & "," &_
  		  						ExpNum2Str(rsLinje("Faktor"),"0000000.0000") & ",""" &_
								RightPad(rsLinje("Serienummer"),15) & """," &_
   		  						ExpNum2Str(rsLinje("Avgiftsnummer"),"00") & "," &_
   		  						ExpNum2Str(rsLinje("FrBruk3"),"00000000.00") & "," &_
								ExpNum2Str(rsLinje("FrBruk4"),"000000") & "," &_
   		  						ExpNum2Str(rsLinje("FrBruk5"),"000000") & "," &_
								ExpNum2Str(rsLinje("FrBruk6"),"000000") & "," &_
   		  						ExpNum2Str(rsLinje("FrBruk7"),"0000") & "," &_
								ExpNum2Str(rsLinje("FrBruk8"),"0000")'

						Response.Write OutputString & "<br>"

						OutStream.WriteLine OutputString

						rsLinje.MoveNext
					Loop
					rsOrdre.MoveNext
				Loop
				'Creating extra basic customer data file, added 21.12.2001(E.L)

				strKundeSQL = "SELECT DISTINCT " & _
					"EKSPORT_RUB_ORDRE.Kundenummer, " & _
					"EKSPORT_RUB_ORDRE.Kjedenummer, " & _
					"EKSPORT_RUB_ORDRE.Kundenavn, " & _
					"EKSPORT_RUB_ORDRE.AdresseI, " & _
					"EKSPORT_RUB_ORDRE.AdresseII, " & _
					"EKSPORT_RUB_ORDRE.Postnummer, " & _
					"EKSPORT_RUB_ORDRE.Poststed, " & _
					"EKSPORT_RUB_ORDRE.Deresref, " & _
					"EKSPORT_RUB_ORDRE.Alfasort, " & _
					"FIRMA.Telefon, " & _
					"FIRMA.Fax, " & _
					"FIRMA.Orgnr, " & _
					"FIRMA.Kreditgrense, " & _
					"EKSPORT_RUB_ORDRE.Levnavn, " & _
					"EKSPORT_RUB_ORDRE.LevadresseI, " & _
					"EKSPORT_RUB_ORDRE.LevadresseII, " & _
					"EKSPORT_RUB_ORDRE.Levpostnummer, " & _
					"EKSPORT_RUB_ORDRE.Levpoststed, " & _
					"EKSPORT_RUB_ORDRE.Medarbeidernummer, " & _
					"EKSPORT_RUB_ORDRE.Distrikt, " & _
					"EKSPORT_RUB_ORDRE.Betalingsbetingelse, " & _
					"EKSPORT_RUB_ORDRE.Kundegruppe, " & _
					"EKSPORT_RUB_ORDRE.Rabattgruppe, " & _
					"EKSPORT_RUB_ORDRE.Behprofil, " & _
					"EKSPORT_RUB_ORDRE.Priskode " & _
					"FROM EKSPORT_RUB_ORDRE " & _
					"LEFT OUTER JOIN FIRMA ON EKSPORT_RUB_ORDRE.Kundenummer = FIRMA.Firmaid " & _
					"WHERE EKSPORT_RUB_ORDRE.Status = '1' "

					set rsKundeData = GetFirehoseRS(strKundeSQL, Conn)

					If HasRows(rsKundeData) Then

						kundeFile = Application("RubiconXtraFileRoot") & "C205KUND.csv"
						'kundeFile = "c:\temp\C205KUND.csv"

						Set kundeStream = FileObject.CreateTextFile(kundeFile, True, False)

						Do Until rsKundeData.EOF

							OutputString = ExpNum2Str(rsKundeData("kundenummer").Value,"000000") & "," &_
								ExpNum2Str(rsKundeData("kjedenummer").Value,"000000") & ",""" &_
								RightPad(rsKundeData("kundenavn").Value, 30) & """,""" &_
	  							RightPad(rsKundeData("AdresseI").Value, 30) & """,""" &_
								RightPad(rsKundeData("AdresseII").Value, 30) & """,""" &_
	  							RightPad(rsKundeData("postnummer").Value, 6) & """,""" &_
								RightPad(rsKundeData("poststed").Value, 25) & """,""" &_
								RightPad(rsKundeData("deresref").Value, 30) & """,""" &_
								RightPad(rsKundeData("telefon").Value, 15) & """,""" &_
								RightPad(rsKundeData("fax").Value, 15) & """,""" &_
								RightPad(rsKundeData("alfasort").Value, 15) & """,""" &_
								RightPad("", 15) & """,""" &_
								RightPad("", 15) & """,""" &_
								RightPad(rsKundeData("orgnr").Value, 15) & """," &_
	  							ExpNum2Str(rsKundeData("kreditgrense").Value,"00000000.00") & ",""" &_
	  							RightPad(rsKundeData("levnavn").Value,30) & """,""" &_
								RightPad(rsKundeData("levadresseI").Value,30) & """,""" &_
   	  							RightPad(rsKundeData("levadresseII").Value,30) & """,""" &_
								RightPad(rsKundeData("levpostnummer").Value,6) & """,""" &_
								RightPad(rsKundeData("levpoststed").Value,25) & """," &_
   	  							ExpNum2Str(rsKundeData("medarbeidernummer").Value,"000") & "," &_
   	  							ExpNum2Str(rsKundeData("distrikt").Value,"0000") & "," &_
   	  							ExpNum2Str(rsKundeData("betalingsbetingelse").Value,"00") & "," &_
   	  							ExpNum2Str("","000000") & "," &_
   	  							ExpNum2Str(rsKundeData("kundegruppe").Value,"000") & "," &_
   	  							ExpNum2Str(rsKundeData("rabattgruppe").Value,"00") & "," &_
   	  							ExpNum2Str(rsKundeData("behprofil").Value,"00") & "," &_
   	  							ExpNum2Str(rsKundeData("priskode").Value,"00")

							kundeStream.WriteLine OutputString
							rsKundeData.MoveNext
						Loop
						Set kundeStream = Nothing
						rsKundeData.Close
						set rsKundeData = nothing
						Response.Write "<br>" & "Fil med kontaktdata for Rubicon ble opprettet under: " & kundeFile & "<br>"
					End If
					'Cleaning up..
					Set OutStream = Nothing
					Set BkStream = Nothing

					rsOrdre.Close
					set rsOrdre = nothing
					rsLinje.Close
					set rsLinje = nothing
					call util.Logoff()		
					set util = nothing
					
					Response.Write "<p>Filen(e) " & TestFile & "og "& backUpFile & " er sist oppdatert: " & Date & "</p>"
				%>
				</div>
			</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>