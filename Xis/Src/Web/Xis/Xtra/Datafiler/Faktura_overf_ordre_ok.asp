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

	dim valgt_avd
	Dim strWhereAvd 'Where-clause for utvalg på avdeling
	dim Conn
	dim strSQL

	' Open database connection
	SET Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	strDato = dbDate(Date)
	periode = Request("periode")

	if Request.QueryString("avd") <> "" then
		valgt_avd = CInt(Request.QueryString("avd"))
	else 
		valgt_avd = 0
	end if

	'Lag where-clause for avdeling avh. av om det er valgt en avdeling eller ikke
	if valgt_avd > 0 then
		strWhereAvd = " AND avdeling = " & valgt_Avd & " "
	else 
		strWhereAvd = ""
	end if


	'Finner riktig årstall for periode slik at dette også blir riktig i årsskiftet
	'Henter den seneste datoen fra timeliste-tabellen og bruker denne som år i periode
	strSQL = "SELECT senest=MAX([dv].[dato]) FROM [dagsliste_vikar] AS [dv]" &_
		" WHERE [dv].[Fakturastatus] = 2" 
		
	SET rsStoersteDato = GetFirehoseRS(strSQL, conn)
	If hasRows(rsStoersteDato) then
		periode = (DatePart("yyyy", rsStoersteDato("senest")) * 100) + periode
		response.write(periode)
		rsStoersteDato.close
	Else
		periode = (DatePart("yyyy", Date)*100) + periode
	End if	
	SET rsStoersteDato = nothing

	' UPDATE FakturaNr, fakturadato og status
	strSQL = "SELECT DISTINCT Kontakt, SOKontakt, FakturaNr, Split FROM FAKTURAGRUNNLAG" &_
		" WHERE Status = 2" & strWhereAvd &_
		" ORDER BY Kontakt, SOKontakt "


	SET rsFnr = GetFirehoseRS(strSQL, conn)
	If hasRows(rsFnr) then
		Conn.BeginTrans
		DO WHILE NOT rsFnr.EOF
			FNR = rsFnr("Fakturanr")
			KONT = rsFnr("Kontakt")
			if(IsNull(rsFnr("Kontakt")) = false) then
				kontaktSQL = " [BestilltAv] = " & rsFnr("Kontakt")
				kontaktSql2 = " [BestilltAv] = " & rsFnr("Kontakt")
			else
				kontaktSQL = " [SOBestilltAv] = " & rsFnr("SOKontakt")
				kontaktSql2 = " [SoPeiD] = " & rsFnr("SOKontakt")
			end if
			splitt = rsFnr("split")

			' UPDATE status in VIKAR_UKELISTE
			strSQL = "UPDATE [VIKAR_UKELISTE] SET" &_
			" [Fakturanr] = " & FNR &_
			", [Overfort_fakt_status] = 3" &_
			", [FakturaDato] = " & strDato &_
			", [Faktperiode] = " & periode &_
			" WHERE " &_
			kontaktSQL &_
			" AND [Overfort_fakt_status] = 2"
			response.write(strSql)

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollbackTrans
				CloseConnection(Conn)
				SET Conn = nothing
				AddErrorMessage("Feil ved oppdatering av faktura status på ukeliste!")
				call RenderErrorMessage()
			End if

			' UPDATE status in DAGSLISTE_VIKAR (timelisten)
			strSQL = "UPDATE DAGSLISTE_VIKAR SET" &_
			" Fakturanr = " & FNR &_
			", Fakturastatus = 3" &_
			", FakturaDato = " & strDato &_
			", Faktperiode = " & periode &_
			" WHERE " &_
			kontaktSQL &_			
			" AND Fakturastatus = 2" 
			
			

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollbackTrans
				CloseConnection(Conn)
				SET Conn = nothing
				AddErrorMessage("Feil ved oppdatering av faktura status på dagsliste!")
				call RenderErrorMessage()
			End if
			
			'Update addtion table
			
			strSQL = "UPDATE ADDITION SET InvNo =" & FNR &_
			", InvStatus= 3 " &_
			",InvDate =" & strDato &_
			", InvPeriod =  " & periode &_
	                " FROM ADDITION INNER JOIN OPPDRAG ON Addition.Oppdragid = OPPDRAG.Oppdragid WHERE " &_
	                kontaktSql2 &_ 
	                " AND InvStatus = 2 "
	            response.write(strSql)
		If ExecuteCRUDSQL(strSQL, Conn) = false then
				Conn.RollbackTrans
				CloseConnection(Conn)
				SET Conn = nothing
				AddErrorMessage("Feil ved oppdatering av faktura status på Addtions!")
				call RenderErrorMessage()
			End if

			rsFnr.MoveNext
		loop

		rsFnr.Close
		SET rsFnr = Nothing

		' UPDATE status in FAKTURALINJER
		strSQL = "UPDATE FAKTURAGRUNNLAG SET " &_
			" FakturaDato = " & strDato &_
			", Status = 3" &_
			", Faktperiode = " & periode &_
			" WHERE Status = 2 " & strWhereAvd 

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollbackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("Feil ved oppdatering av fakturadato og status på fakturagrunnlag!")
			call RenderErrorMessage()
		End if

		' UPDATE status in EKSPORT_RUB_ORDRE
		strSQL = "UPDATE EKSPORT_RUB_ORDRE SET " &_
			"status = 2, " &_
			"Eksportert_Dato = " & strDato &_
			" WHERE status = 1" & strWhereAvd 

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollbackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("Feil ved oppdatering av rubikon ordrehode!")
			call RenderErrorMessage()
		End if

		' UPDATE status in EKSPORT_RUB_ORDRELINJE
		strSQL = "UPDATE EKSPORT_RUB_ORDRELINJE SET " &_
			"status = 2, " &_
			"Eksportert_Dato = " & strDato &_
			" WHERE status = 1" & strWhereAvd 
			
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollbackTrans
			CloseConnection(Conn)
			SET Conn = nothing
			AddErrorMessage("Feil ved oppdatering av rubikon ordrelinjer!")
			call RenderErrorMessage()
		End if

		
	End If 'ingen rader
	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing
	
	'Push another page
	Response.Redirect "../WebUI/Admin/Invoice/InvoiceList.aspx?selectedDept=" & valgt_avd
%>