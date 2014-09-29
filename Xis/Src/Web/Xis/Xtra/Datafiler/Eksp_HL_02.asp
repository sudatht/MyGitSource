<!--#INCLUDE FILE="../includes/Library.inc"-->

<%
	dim valgt_avd
	dim strOverfort
	dim blnOKtoTransfer
	dim sFortegn

	If Request.QueryString("avd") <> "" Then
		valgt_avd = CInt(Request.QueryString("avd"))
	else
		valgt_avd = 0
	End If

	function NOFormatNumber(number, NofDecimalDigits)
			NOFormatNumber = mid(number,1, len(number) - NofDecimalDigits) & "," & mid(number, len(number) - NofDecimalDigits + 1)
	end function
%>
<html>
	<head>
		<title>Overføring av lønn</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript">
			function sjekkPeriode(form, sjekkfelt, url)
			{
				verdi = parseInt(document.forms(form).elements(sjekkfelt).value);
				if ((verdi >= 1) && (verdi <= 12)){
					if (janei = confirm("Bekreft importen for periode " + verdi + "?")){
						(url.indexOf("?") > 1)? delim = "&" : delim = "?";
						document.forms(form).action = url + delim + "periode=" + verdi;
						document.forms(form).submit();
					}else{
						document.forms(form).elements(sjekkfelt).focus();
					}
				}else{
					alert("Ikke riktig periode tall");
				}
			}
		</SCRIPT>
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Eksport av Lønnsdata til H&amp;L</h1>
			</div>
			<div class="content">
				<%
				'--------------------------------------------------------------------------------------------------
				' Modifications
				' Date - Who - What
				' 23.01.2000 - Arne Leithe - Included and used GetAvdeling to get the correct department name
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

				Function AvdNr(strAvdelingID)
					Dim strSQl
					strSQL = "select a.avdnr from avdeling a where a.avdelingID = " & strAvdelingID
					Set rsAvdeling = conn.Execute(strSQL)
					If not rsAvdeling.EOF then
						AvdNr = rsAvdeling("avdnr")
					Else
						AvdNr = "0000"
					End if
					rsAvdeling.close
					set rsAvdeling = nothing
				End Function


				%>
				<h2>Eksport av variable lønnsopplysninger til Hult &amp; Lillevik.</h2>
				<%
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				''''''''''''''''''''
				''''''''''''''''''''	VARIABLE LØNNSOPPLYSNINGER   -  NYE OG ENDREDE VIKARER
				''''''''''''''''''''
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

				Set rsVikar = Conn.Execute("SELECT " & _
					"VIKAR.VikarId, " & _
					"VIKAR.[H_L_Status], " & _
					"VIKAR_ANSATTNUMMER.ansattnummer, " & _
					"VIKAR_LOEN_VARIABLE.Avdeling, " & _
					"VIKAR_LOEN_VARIABLE.Prosjektnr, " & _
					"VIKAR_LOEN_VARIABLE.Loennsartnr, " & _
					"VIKAR_LOEN_VARIABLE.Dato, " & _
					"Ant = VIKAR_LOEN_VARIABLE.Antall * 100, " & _
					"Sat = VIKAR_LOEN_VARIABLE.Sats * 100, " & _
					"Bel = VIKAR_LOEN_VARIABLE.Beloep * 1000, " & _					
					"ISNULL((SELECT  FrameworkAgreement.FaCode FROM " & _         
					"OPPDRAG LEFT OUTER JOIN FrameworkAgreement ON " & _ 
					"OPPDRAG.FaID = FrameworkAgreement.FaID " & _ 
					"Where OPPDRAG.OPPDRAGID = VIKAR_LOEN_VARIABLE.OppdragId),0)As FaCode " & _
					"FROM VIKAR LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON " & _ 
					"VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid, " & _
					"VIKAR_LOEN_VARIABLE WHERE VIKAR.VikarID = VIKAR_LOEN_VARIABLE.VikarID " & _ 
					"AND Avdeling = '" & valgt_avd & "' " & _
					"AND VIKAR_LOEN_VARIABLE.Overfor_Loenn_status = '2' ")
				' No records found
				If (rsVikar.EOF) Then
				Response.write "<H4>Ingen rader i tabellen!</H4>"
				Else

					set rsAvdnavn = Conn.Execute ("select avdeling from Avdeling where avdelingID = " & valgt_avd)
					%>

					<table class="listing">
						<tr>
							<th>Ansattnummer</th>
							<th>Dato</th>
							<th>Avdnr</th>
							<th>Avdnavn</th>
							<th>Prosjektnr</th>
							<th>Lønnsart</th>
							<th>Antall</th>
							<th>Sats</th>
							<th>Beløp</th>
						</tr>
					<%
					'Initialize to OK!
					blnOKtoTransfer = true
					Do Until (rsVikar.EOF)
						strAvdeling = AvdNr(rsVikar("Avdeling"))
						if (rsVikar("H_L_Status").Value=0) then
							strOverfort = "Nei"
						else
							strOverfort = "Ja"
							blnOKtoTransfer = false
						end if
						%>
							<tr>
								<td><%=rsVikar("ansattnummer").Value%></td>
								<td class="right">&nbsp;<%=rsVikar("Dato").Value%></td>
								<td>&nbsp;<%=strAvdeling%></td>
								<td><%=rsAvdnavn("avdeling").Value%></td>
								<td class="right">&nbsp;<%=rsVikar("Prosjektnr").Value%></td>
								<td class="right">&nbsp;<%=rsVikar("Loennsartnr").Value%></td>
								<td class="right">&nbsp;<% if rsVikar("Ant").Value > 0 then %><%=FormatNumber(rsVikar("Ant").Value / 100,2)%><% end if%></td>
								<td class="right">&nbsp;<% if rsVikar("Sat").Value > 0 then %><%=FormatNumber(rsVikar("Sat").Value / 100,2)%><% end if%></td>
								<td class="right">&nbsp;<% if rsVikar("Bel").Value > 0 then %><%=FormatNumber(rsVikar("Bel").Value/ 1000,2)%><%end if%></td>
							</tr>
						<%
						rsVikar.MoveNext
					Loop
					%>
					</table>
					<%
					if (blnOKtoTransfer = true) then
						set  util = Server.CreateObject("XisSystem.Util")				
						call util.Logon()

						Set FileObject = Server.CreateObject("Scripting.FileSystemObject")

						User = Request.ServerVariables("LOGON_USER")
						pos = Instr(User,"\")
						User = Mid(User, pos + 1)
						nn = OnlyDigits(Now)
						nn = Left(nn,10)
						name = "H001_" & nn & "_" & User & ".txt"

						'"rubicon_test" is the test area for HL / rubicon files, "rubicon" is the production area
						BackupFile = Application("RubiconFileRoot") & name
						TestFile = Application("RubiconFileRoot") & "H01.txt"

						Response.write "HovedFil:<a target='new' href='" & TestFile & "'>" & TestFile & "</a><br>"
						Response.write "Sikkerhetskopi:<a target='new' href='" & BackupFile & "'>" &  BackupFile & "</a><br>"

						Set OutStream = FileObject.CreateTextFile (TestFile, True, False)
						Set BkStream = FileObject.CreateTextFile (BackupFile, True, False)
						set FileObject = nothing

						rsVikar.MoveFirst
						Do Until rsVikar.EOF
							'02.11.00 Endret uthenting av avdelingsnummer til å hente fra avdelingstabellen
							'strAvdeling = GetAvdeling(Session("avdkontornavn"), rsVikar("Avdeling"))

							strAvdeling = AvdNr(rsVikar("Avdeling").Value)
							'--------------------------------------------------------------------------------------------------
							'       Varabel = innhold			    feltlengde     kommentar
							'--------------------------------------------------------------------------------------------------
							LoennsNr		= ExpNum2Str(rsVikar("ansattnummer").Value,"000000")	'6
							Loennsart		= LeftFill(rsVikar("Loennsartnr").Value,5,"0")		'5
							Avdeling		= LeftFill(strAvdeling,12,"0")		'12
							Prosjekt		= LeftFill(rsVikar("Prosjektnr").Value,12,"0")	'12
							Element01		= LeftFill(rsVikar("FaCode").Value,12,"0")	'12
							Element			= leftFill("",12*4,"0")
							If Not IsNull(rsVikar("Dato").Value) Then
								Dato = OnlyDigits(rsVikar("Dato").Value)		'6
							Else
								Dato = "000000"
							End If

							If rsVikar("Ant") < 0 Then
								sFortegn = "-"
							Else
								sFortegn = " "
							End If							'1
							Antall			= sFortegn & LeftFill(rsVikar("Ant").Value,9,"0")		'10 (100-deler)

							If rsVikar("Sat") < 0 Then
								sFortegn = "-"
							Else
								sFortegn = " "
							End If							'1
							Sats			= sFortegn & LeftFill(rsVikar("Sat").Value,9,"0")		'10 (ører)

							If rsVikar("Bel") < 0 Then
								sFortegn = "-"
							Else
								sFortegn = " "
							End If							'1
							Beloep			= sFortegn & LeftFill(OnlyDigits(rsVikar("Bel").Value),13,"0")	'13 (ører)
							Filler1			= "000000000000000000000000000000"			'30

							'--------------------------------------------------------------------------------------------------
							'      lager record streng
							'--------------------------------------------------------------------------------------------------

							RecStr = ""
							RecStr = LoennsNr & _
							Loennsart & _
							Avdeling & _
							Prosjekt & _
							Element01 & _  
							Element	 & _
							Dato  & _
							Antall & _
							Sats & _
							Beloep & _
							Filler1

							OutStream.WriteLine RecStr
							BkStream.WriteLine RecStr

							'OutStream.WriteLine Chr(13)

							rsVikar.MoveNext
						Loop

						Set OutStream = Nothing
						Set BkStream = Nothing
						
						call util.Logoff()		
						set util = nothing						
						%>
						<p>
							<h2>Var overføringen vellykket?</h2>
							<FORM NAME="Ja" ACTION="Vikar_overf_loenn_ok.asp?viskode=<%=Request("viskode")%>&avd=<%=valgt_avd%>" METHOD="post">
								Periode: <input type="text" NAME="periode" ID="Text1"><br>
								<input name="jaknapp" TYPE="BUTTON" VALUE="Ja" onClick="sjekkPeriode(form.name, form.elements(0).name, form.action )">
							</form>
							<FORM ACTION="Vikar_timeliste_list3.asp?viskode=<%=Request.QueryString("viskode")%>&avd=<%=valgt_avd%>" TARGET=_top METHOD="POST" ID="Form1">
								<INPUT name="neiknapp" TYPE="SUBMIT" VALUE="Nei">
							</form>
						</p>
						<%
					end if
				End If  'ingen rader i tabellen

				rsVikar.Close
				set rsVikar = nothing
				%>
			</div>
		</div>
	</body>
</html>