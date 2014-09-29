<%
Sub SqlError()
   ' Print error messages
   Response.write "<p class='warning'>Database Errors Occured!" & "</p>"
   Response.write "<p><strong>Description:</strong><br>" & conn.errors(0).description & "</p>"
   Response.write "<p><strong>Statement:</strong><br>" & strSQl & "</p>"
   Response.End
End Sub

Sub SqlError2(objCon, strSQL)
	dim NofErrors
	dim iCurrentError
	' Print error messages
	if (objCon.errors.count>0) then
		Response.write "<p class='warning'>Database Errors Occured!</p>"
		Response.write "<strong>Statement:</strong><br>" & strSQL & "</p>"
		NofErrors = objCon.errors.count
		iCurrentError = 0
		while (iCurrentError < NofErrors)
			Response.write "<p><strong>Description:</strong><br>" & objCon.errors(iCurrentError).description
			iCurrentError = iCurrentError + 1
		wend
	end if
	Response.End
End Sub


function ValidDateInterval(fromDate, toDate)
	if (fromDate = DateSerial(100, 1, 1 ) or toDate = DateSerial(100, 1, 1 )) then
		ValidDateInterval = false
		exit function
	end if
		
	if (fromDate > toDate) then
		ValidDateInterval = false		
	else
		ValidDateInterval = true		
	end if
	
end function

function ToDateFromDDMMYY(stringDate)
	dim strDay
	dim strMonth
	dim strYear

	if (len(stringDate) = 8) then
		strDay = mid(stringDate, 1, 2)
		strMonth = mid(stringDate, 4, 2)
		strYear	= mid(stringDate, 7, 2)
		ToDateFromDDMMYY = DateSerial(strYear, strMonth, strDay )
		
	else
		ToDateFromDDMMYY = DateSerial(100, 1, 1 ) 'invalid format, return lowest date
	end if
end function


function fixID(instring)
	if (isnull(instring)) OR len(trim(instring))=0 OR trim(instring)="0" then
		fixID = "null"
	else
		fixID = "'" & instring & "'"
	end if
end function

function FixString(instring)
	if (isnull(instring)) OR len(trim(instring))=0 then
		FixString = "null"
	else
		FixString = "'" & Replace(instring, "'", "''") & "'"
	end if
end function

function ToNorDate(instring)
	dim strDay
	dim strMonth

	if (isnull(instring)) OR (not isDate(instring)) then
		ToNorDate = ""
	else
		if day(instring) < 10 then
			strDay = "0" & day(instring)
		else
			strDay = day(instring)
		end if
		if Month(instring) < 10 then
			strMonth = "0" & Month(instring)
		else
			strMonth = Month(instring)
		end if
		ToNorDate = strDay & "." & strMonth & "." & Year(instring)
	end if
end function






function PadQuotes(instring)
		PadQuotes = Replace(instring, "'", "''")
end function



Sub se(txt)
	Response.write txt & "<BR>"
End Sub

Sub se0(txt)
	Response.write txt
End Sub

Sub seVar()
	for each x in request.querystring
		se x & "=" & Request(x)
	next
	for each x in request.form
		se x & "=" & Request(x)
	next
End Sub

Sub fjernKomma(strString )
	dim pos
    pos = InStr(strString, ",")
    If pos <> 0 Then
        strString = Left(strString, pos - 1) & "." & Left(Mid(strString, pos + 1), 2)
    End If
End Sub

Sub settKomma(strString)
    pos = InStr(strString, ".")
    If pos <> 0 Then
        strString = Left(strString, pos - 1) & "," & Mid(strString, pos + 1)
    End If
End Sub

'*******************************************************
'korrigerer for første uke i året og første dag i uka  *
'*******************************************************
sub datoKorreksjon(del, datoen)                       
 	datoen=DatePart(del,datoen,2,2) 		              
end sub                                               
'*******************************************************
'---------------------------------------------------------------------------------------------
'KorrigerUke
'Parameters: date
'Returns: Correct week number
'---------------------------------------------------------------------------------------------
sub KorrigerUke(ad_dato1)                       
    ad_dato1 = DatePart("ww", ad_dato1, 2, 2)
    If ad_dato1 < 10 Then
        ad_dato1 = "0" & ad_dato1
    End If
end sub                           

'---------------------------------------------------------------------------------------------
'KorrigerAar
'Parameters: date
'Returns: Correct year a certain week
'---------------------------------------------------------------------------------------------
sub KorrigerAar(ad_dato2)                       
 	
 	maaned= DatePart("m", ad_dato2, 2, 2)
 	uke= DatePart("ww", ad_dato2,2, 2)
 	ad_dato2 = DatePart("YYYY",ad_dato2,2,2)
 	If maaned = 1 and uke > 6 then
 		ad_dato2 = ad_dato2 - 1
 	End if		
end sub  
  
'----------------------------------------------------------------------------------------
' GetAvdeling: 
' parameters: department name ("OSLO", "DRAMMEN"..), departmentID (1,2,3,4)
' returns: correct department number
'----------------------------------------------------------------------------------------
function GetAvdeling(akn, avdid)
		if strcomp(akn,"OSLO") = 0 then
			if avdid = 1 then 
				Avdeling = 2203 'Kurs
			elseif avdid = 2 then
				Avdeling = 2201 'Data
			elseif avdid = 3 then
				Avdeling = 2202 'Dokument
			elseif avdid = 5 then
				Avdeling = 2204 'Kontor
			else
				Avdeling = ""
			end if
		elseif strcomp(akn,"DRAMMEN") = 0 then
			Avdeling = 2250
		elseif strcomp(akn, "PORSGRUNN") = 0 then
			Avdeling = 2260
		elseif strcomp(akn, "STAVANGER") = 0 then
			Avdeling = 2550
		elseif strcomp(akn, "BERGEN") = 0 then
			Avdeling = 2650
		elseif strcomp(akn, "TRONDHEIM") = 0 then
			Avdeling = 2750
		else
			Avdeling = ""
		end if
		GetAvdeling = Avdeling
end function


'----------------------------------------------------------------------------------------
' HentKontor: 
' parameters: none
' returns: navn på kontor brukeren er logget på som
'----------------------------------------------------------------------------------------
Function HentKontor()
	Dim AvdID
	Dim Avdkontor

	AvdID = Session("avdkontor")
	Select Case avdID
		Case 0
			Avdkontor = "(OSLO)"
		Case 1
			Avdkontor = "(DRAMMEN)"
		Case 2
			Avdkontor = "(PORSGRUNN)"
		Case 3
			Avdkontor = "(STAVANGER)"
		Case 4
			Avdkontor = "(BERGEN)"
		Case 5
			Avdkontor = "(REKRUTTERING)"	
		Case 6
			Avdkontor = "(TRONDHEIM)"	
	End select				
		
HentKontor = Avdkontor
End function

'--------------------------------------------------------------------------------------------
'HentAvdNavn
'Parameters: department ID
'Returns: Correct department name if found, "ukjent" if not found.
'---------------------------------------------------------------------------------------------
Function HentAvdnavn(AvdelingID)
	Dim rsAvd

	Set rsAvd = conn.Execute ("select a.avdeling from avdeling a where a.avdelingID = " &AvdelingID)
	If not rsAvd.EOF then
		HentAvdnavn = rsAvd("avdeling")
	Else
		HentAvdnavn = "Ukjent"
	End if
	rsAvd.close
	Set rsAvd = Nothing		
End function


'--------------------------------------------------------------------------------------------
'GmlOppdragID
'Parameters: OppdragID
'Returns: Old OppdragID from before converting to one central DB
'---------------------------------------------------------------------------------------------
Function GmlOppdragID(aOppdragID)
	Dim rsGmlOppdragID
	Dim strGmlOppdragID
	
	strSQL = "select gmlOppdragID from oppdrag where oppdragID = " & aOppdragID
	Set rsGmlOppdrag = conn.Execute(strSQL)
	strGmlOppdragID = rsGmlOppdrag("gmlOppdragid")
	rsGmlOppdrag.close
	set rsGmlOppdrag = nothing 
	if isnull(strGmlOppdragID) then
		GmlOppdragID = 0
	Else
		GmlOppdragID = strGmlOppdragID	
	End if
	
End function


Function GetSetting(strSettingName)
	Dim rsSetting
		
	strSQL = "SELECT SettingValue FROM Intra_settings WHERE settingName = '" & trim(strSettingName) & "'"

	Set rsSetting = conn.Execute(strSQL)
	if (rsSetting.EOF = false) then
		GetSetting = rsSetting.fields("SettingValue").value
	else
		Response.Write "Følgende innstilling/tekst mangler i Intrasettings:" & strSettingName &  "<br>"
		Response.end
	end if
	rsSetting.close
	set rsSetting = nothing	
End function

public function GeneratePassword(numberOfChars)
	'Array of legal letters for password
	dim legalLetters
	'Array of legal digits for password
	dim legalDigits
	dim returnValue
	dim letterLength, digitsLength
	dim currentPos
	dim charChance
	dim roll
	dim nofDigits, nofLetters
	dim exitLoop

	'Initialize
	charChance = 70 '% Chance of being a character
	currentPos = 1
	legalLetters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
	legalDigits = Array("2", "3", "4", "5", "6", "7", "8", "9")
	letterLength = UBound(legalLetters)
	digitsLength = UBound(legalDigits)

	call randomize()
	while (currentPos <= numberOfChars)
		'Use letter or digit?
		roll  = Int((100 - 1) * Rnd())
		if (roll <= charChance) then
		'add letter
			index = Int((letterLength - 1) * Rnd())
			nofLetters = nofLetters + 1
			returnValue = returnValue + legalLetters(index)
		else
		'add digit
			index = Int((digitsLength - 1) * Rnd())
			nofDigits = nofDigits + 1
			returnValue = returnValue + legalDigits(index)
		end if
		currentPos = currentPos + 1
	wend

	GeneratePassword = returnValue
end function

%>