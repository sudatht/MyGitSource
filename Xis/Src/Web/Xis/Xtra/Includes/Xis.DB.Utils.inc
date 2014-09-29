<%
	'This function return datetime now() in string format 2000-01-16 9:56:33
	Function GetDateNowString()
		GetDateNowString = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
	End Function
	
	Function DbDate( strDate )
	' This function convert date from date setting to Databaseformat 
		dim ddate 'as date
		dim lYear  'as integer
		dim lMonth 'as integer
		dim lDay   'as integer
		dim strDay 'as string
		dim strMonth 'as string
		dim strYear 'as string
	   
	If IsDate (strDate) Then
		DDate   = CDate( strDate ) 
		lYear  = Year(Datevalue( DDate ) )
		lMonth = Month( Datevalue( DDate ) )
		lDay   = Day( Datevalue( DDate ) )

		If lDay < 10 Then
			strDay = "0" & Cstr( lDay )  
		Else 
			strDay = Cstr( lDay )  
		End If

		If lMonth < 10 Then
			strMonth = "0" & Cstr( lMonth )  
		Else 
			strMonth = Cstr( lMonth )  
		End If

		strYear = Cstr( lYear )
	      
		DbDate = strMonth & "." & strDay & "." & strYear
	 
		DbDate = "'" & DbDate & "'" 
		Else
		DbDate = "null"
		End If

	End Function

	Function DbTime( strTime )
	If IsDate (strTime) and strTime <> "" Then
		DbTime = "'" & strTime & "'"
		Else
		DbTime = "null"
		End If

	End Function

	Function Quote( strString )
		Quote = "'" & strString & "'"
	End Function
	
	function PadQuotes(instring)
		PadQuotes = Replace(instring, "'", "''")
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

	Sub fjernKomma(strString )
		dim pos
		pos = InStr(strString, ",")
		If pos <> 0 Then
			strString = Left(strString, pos - 1) & "." & Left(Mid(strString, pos + 1), 2)
		End If
	End Sub

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
	
	'gjør om "limitDato" til siste dag i uken
	Sub newLimitDate(strDato) 	
		limitDato = Session("limitDato")
		if limitDato <> "" then 
			strStartDato = Request("StartDato")
			strDato = DBdate(limitdato)
			nyStrDato = left(strDato, len(strDato) - 1)
			nyStrDato = right(nyStrDato, len(nyStrDato) - 1)
			nyStrDato = mid(nyStrDato,4,3) & left(nyStrDato,3) & right(nyStrDato,4)
			korrDato = nyStrDato
			if isdate(korrDato) then 
				call datoKorreksjon("w", korrDato)
			end if
			ukedag = korrDato 
			sisteUkedagDato = DBdate(DateAdd("d", 7 - ukedag, nyStrDato) )
			nyStrDato = left(sisteUkedagDato, len(sisteUkedagDato) - 1)
			nyStrDato = right(nyStrDato, len(nyStrDato) - 1)
			nyStrDato = mid(nyStrDato,4,3) & left(nyStrDato,3) & right(nyStrDato,2)
			limitDato = nyStrDato
			Session("limitDato") = limitDato
		end if 'limidDato = ""
	End Sub	
	
	'korrigerer for første uke i året og første dag i uka 
	sub datoKorreksjon(del, datoen)                       
 		datoen = DatePart(del, datoen, 2, 2)    
	end sub
	
	Public Function WeekFix(DateToFix) 
		dim DateYear
		dim DateMonth
		dim DateDay
		dim DateWeek
		dim firstDayJanuaryDate
		dim firstDayJanuaryStartsAtWeekday
		
		DateYear = Datepart("yyyy",DateToFix, 2, 2)
		DateMonth = Datepart("m",DateToFix, 2, 2)
		DateDay = Datepart("d",DateToFix, 2, 2)
		DateWeek = Datepart("ww",DateToFix, 2, 2)
		
		if ((DateMonth = 12 OR DateMonth = 1) AND (DateWeek = 53 OR DateWeek = 1)) then
			if (DateWeek = 53) then
				firstDayJanuaryDate = DateSerial(DateYear + 1, 1, 1)
			elseif (DateWeek = 1) then
				firstDayJanuaryDate = DateSerial(DateYear, 1, 1)
			end if
			
			firstDayJanuaryStartsAtWeekday = Datepart("w", firstDayJanuaryDate, 2, 2)

			if (firstDayJanuaryStartsAtWeekday >= 5 AND firstDayJanuaryDate >= DateToFix) then
				WeekFix = 53
			else
				WeekFix = 1
			end if
		else
			WeekFix = DateWeek
		end if
	End Function	
%>