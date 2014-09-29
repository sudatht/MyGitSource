<%
	'KorrigerUke
	'Parameters: date
	'Returns: Correct week number
	sub KorrigerUke(ad_dato1)                       
		ad_dato1 = DatePart("ww", ad_dato1, 2, 2)
		If ad_dato1 < 10 Then
			ad_dato1 = "0" & ad_dato1
		End If
	end sub
	
	'KorrigerAar
	'Parameters: date
	'Returns: Correct year a certain week
	sub KorrigerAar(ad_dato2)                       
	 	
 		maaned = DatePart("m", ad_dato2, 2, 2)
 		uke= DatePart("ww", ad_dato2,2, 2)
 		ad_dato2 = DatePart("YYYY",ad_dato2,2,2)
 		If maaned = 1 and uke > 6 then
 			ad_dato2 = ad_dato2 - 1
 		End if		
	end sub  


	'HentAvdNavn
	'Parameters: department ID
	'Returns: Correct department name if found, "ukjent" if not found.

	Function HentAvdnavn(AvdelingID)
		Dim rsAvd
		dim strSQL

		strSQL = "SELECT [a].[avdeling] FROM [Avdeling] AS [a] WHERE [a].[avdelingID] = " &AvdelingID
		Set rsAvd = GetFireHoseRS(strSQL, Conn)
		if(HasRows(rsAvd)) then
			HentAvdnavn = rsAvd("Avdeling")
			rsAvd.close
		Else
			HentAvdnavn = "Ukjent"
		End if
		Set rsAvd = Nothing		
	End function

	Function HentAlleAktiveAvdelinger()
		dim strSQL
		dim rsAvdeling
		
		strSQL = "SELECT [AvdelingID], [Avdeling] FROM [avdeling] ORDER BY [Avdeling]"
		Set rsAvdeling = GetFireHoseRS(strSQL, Conn)
		set HentAlleAktiveAvdelinger = rsAvdeling
	End function
	

	sub datoKorreksjon(del, datoen)                       
 		datoen = DatePart(del, datoen, 2, 2)
	end sub                                               

	'gjør om "limitDato" til siste dag i uken
	Sub newLimitDate(strDato) 	
		limitDato = Session("limitDato")

		if limitDato<> "" then 
			strStartDato = Request("StartDato")
			strDato = DBdate(limitdato)
			nyStrDato= left(strDato, len(strDato)-1)
			nyStrDato= right(nyStrDato, len(nyStrDato)-1)
			nyStrDato= mid(nyStrDato,4,3)&left(nyStrDato,3)&right(nyStrDato,4)
			korrDato=nyStrDato
			if isdate(korrDato) then 
				call datoKorreksjon("w", korrDato)
			end if
			ukedag = korrDato 
			sisteUkedagDato = DBdate(DateAdd("d",7-ukedag, nyStrDato) )
			nyStrDato= left(sisteUkedagDato, len(sisteUkedagDato)-1)	
			nyStrDato= right(nyStrDato, len(nyStrDato)-1)
			nyStrDato= mid(nyStrDato,4,3)& left(nyStrDato,3)&right(nyStrDato,2)
			limitDato=nyStrDato
			Session("limitDato")= limitDato
		end if 'limidDato=""
	End Sub

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