<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Sub fjernKomma(strString )
		pos = Instr(strString, ",")
		If pos <> 0 Then
			mellom = Left(strString, pos - 1) & "." & Mid(strString, pos + 1)
 			strString = mellom
		End If
	End Sub

	dim Conn
	dim strSQL
	dim strBestilltAv
	dim strSOBestilltAv
	dim strVikarID
	dim strOppdragID
	dim strFirmaID
	dim strStartdato
	dim strDato
	dim strLinje
	dim strStarttid
	dim strSluttid
	dim strTimeLonn

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Check parameters AND put into variables
	strVikarID = Request.QueryString("VikarID")
	strOppdragID = Request.QueryString("OppdragID")

	strFirmaID = Request.Form("FirmaID")
	strStartdato = Request.Form("StartDato")

	Dato = Request.form("Dato")
	strDato = DbDate(Request.form("Dato"))

	strLinje = Request.QueryString("Linje")

	strStarttid = Request.Form("tbxFraKl")
	strSluttid = Request.Form("tbxTilKl")

	strTimeLonn = Request.Form("TimeLonn")

	If strTimeLonn = "" Then
		strTimeLonn = "NULL"
	End If
	Call fjernKomma(strTimeLonn)

	strSlett = Request.QueryString("Slett")

	strEndre = Request.Form("Endre")

	strFakturapris = Request.Form("Fakturapris")
	Call fjernKomma(strFakturapris)

	strNotat =  Request.Form("Notat")

	strAnttimer = Request.Form("tbxTimerPrDag")
	Call fjernKomma(strAnttimer)

	strFakturaTimer = Request.Form("tbxFaktTimerPrDag")
	Call fjernKomma(strFakturaTimer)

	strLunch = Request.Form("tbxLunsj")

	strLoennsArt = Request.Form("SYK")
	If strLoennsArt = "" Then
		strLoennsArt = "NULL"
	Else
		strLoennsArt = "'" & strLoennsArt & "'"
	End If

	If Request("splittuke") = "splitt" Then splittuke = 1 Else splittuke = "NULL"
	If Request("splittuke") = "splitt" AND NOT Request("Endret") = "1" Then session("splittopt") = splittuke

	if Request("Endret") <> "" then session("splittopt") = Request("Endret")

	'Fjernet lønns- og fakturastatus ved endring av timelister, så dette settes til 1
	strLoennstatus = 1
	strFakturastatus = 1

	If Dato = strStartDato Then
		strStartDato = ""
	End If

	strBestilltAv = Request.Form("BestilltAv")
	If len(strBestilltAv) = 0 Then
		strBestilltAv = "NULL"
	end if
	
	strSOBestilltAv = Request.Form("SOBestilltAv")
	If len(strSOBestilltAv) = 0 Then
		strSOBestilltAv = "NULL"
	end if
	
	if Request("Dato") <> "" AND strStartDato <> "" then
		if dateValue(strStartDato)> dateValue(Request.form("Dato")) then
			strStartDato = dateValue(Request.form("Dato"))
			session("startDato") = dateValue(Request.form("Dato"))
		end If
	end if

	' Delete row in DAGSLISTE_VIKAR
	If strSlett = "Ja" Then
		
		strStartDato = session("startDato")
		'under korrigeres fradato hvis første dag i uke/ukedel slettes
		if trim(request("sletteDato")) = trim(session("startDato")) AND not session("dagteller")=1 then
			
			startdato = session("startdato")
			
			strNyStartDato = "SELECT DISTINCT dato FROM DAGSLISTE_VIKAR " &_
				" WHERE dato > " & DbDate(startdato)  &_
				" AND vikarID IN (SELECT vikarID FROM DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ") "&_
				" AND oppdragID IN (SELECT oppdragID FROM DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ") order by dato"
			
			set rsNyDato = GetFirehoseRS(strNyStartDato, Conn)
			strNyDato = rsNyDato("dato")
			rsNyDato.close
			set rsNyDato = nothing
			strStartDato = strNyDato
		end if

		strSQL = "DELETE FROM DAGSLISTE_VIKAR WHERE TimelisteVikarID = "  & strLinje

		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under sletting av dagsliste.")
			call RenderErrorMessage()		
		end if

		if session("dagteller") = 1 then
			reDir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID &_
			"&viskode=1&OppdragID=" & strOppdragID &_
			"&limitdato=" & session("limitdato")&"&frakode=2"
		else
			reDir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID &_
				"&OppdragID=" & strOppdragID &_
				"&Startdato=" & strStartdato &_
				"&splittopt2=" & Request("splittopt2") &_
				"&splittuke=" & session("splittuke") &_
				"&splittopt=" & Request("splittopt2")
		end if

		Response.redirect reDir

	End IF 'deleting

If strDato <> "" AND strStarttid <> "" AND strSluttid <> ""  AND strLunch <> "" AND strTimelonn <> "" AND strFakturatimer <> "" AND strFakturapris <> "" Then

	' Update row in DAGSLISTE
	If strEndre = "Ja" AND strSlett = "" Then

		nyStrDato = left(strDato, len(strDato)-1)
		nyStrDato = right(nyStrDato, len(nyStrDato)-1)
		nyStrDato = mid(nyStrDato,4,3)&left(nyStrDato,3)&right(nyStrDato,4)
		korrDato = nyStrDato
		if trim(korrDato)<>"" then 
			call datoKorreksjon("w", korrDato)
		end if
		ukedag = korrDato
		forsteUkedagDato = DBdate(DateAdd("d",-ukedag+1, nyStrDato) )
		sisteUkedagDato = DBdate(DateAdd("d",7-ukedag, nyStrDato) )

		strSQL = "UPDATE DAGSLISTE_VIKAR SET" &_
			" Dato = " & strDato &_
			", Starttid = '" & strStarttid & "'" &_
			", Sluttid = '" & strSluttid & "'" &_
			", AntTimer = " & strAnttimer &_
			", Timelonn = " & strTimelonn &_
			", Fakturapris = " & strFakturapris &_
			", Notat = '" & PadQuotes(strNotat) & "'" &_
			", Lunch = '" & strLunch & "'" &_
			", Fakturatimer = " & strFakturatimer &_
			", LoennsArt = " & strLoennsArt &_
			", splittuke = " & session("splittopt") &_
			", fakturastatus = " & strFakturastatus &_
			", loennstatus = " & strLoennstatus &_
			" WHERE TimelisteVikarID = " & strLinje

		if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under oppdatering av dagsliste.")
			call RenderErrorMessage()		
		end if

		If Request("splittuke") = "splitt" AND NOT (Request("Endret") = "1" or Request("Endret") = "2") then

			StrSQL1 =  "Update DAGSLISTE_VIKAR set" &_
				" splittuke = 2 " &_
				" WHERE vikarID in(SELECT vikarID from DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ")" &_
				" AND oppdragID in(SELECT oppdragID from DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ")" &_
				" AND Dato > " & strDato &_
				" AND Dato<= " & sisteUkedagDato

			if (ExecuteCRUDSQL(StrSQL1, Conn) = false) then
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av splittuke på timeliste.")
				call RenderErrorMessage()		
			end if
			
			StrSQL2 =  "Update DAGSLISTE_VIKAR set" &_
				" splittuke = 1 "&_
				" WHERE vikarID in(SELECT vikarID from DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ")" &_
				" AND oppdragID in(SELECT oppdragID from DAGSLISTE_VIKAR WHERE TimelisteVikarID = " & strLinje & ")" &_
				" AND Dato <= " & strDato &_
				" AND Dato >= " & forsteUkedagDato

			if (ExecuteCRUDSQL(StrSQL2, Conn) = false) then
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av splittuke på timeliste.")
				call RenderErrorMessage()		
			end if

			session("splittopt") ="1"
			Session("splittopt2")=".1"
			session("splittuke") = "Ja"
			

		else
			Session("splittopt2") = "."&request("splittopt2")
		end if

	Else 'insert (not delete or update)
	'Star Rashid: 18.06.01 OBS! denne rutinen tar ikke hensuyn til splittuke
	' Insert into database DAGSLISTE_VIKAR

		splittverdi = request("splittopt2")
		session("splittopt") = splittverdi
		if lcase(splittverdi) <> "null" then
			splittverdi = cInt(splittverdi)
			session("splittuke") = "Ja"
		end if

		'Get oppdragvikarid
		strSQL = "SELECT [OppdragVikarId] " & _
				" FROM [OPPDRAG_VIKAR] " & _
				" WHERE [VikarID] = "  & strVikarID  & _
				" AND [OppdragID] = "  & strOppdragID  & _
				" AND " & strDato & " >= [FraDato] AND " & strDato & " <= [TilDato] "
		dim rs
		dim strOppdragVikarId

		set rs	= GetFirehoseRS(strSQL, Conn)
		if (hasRows(rs) = false) then
			set rs = nothing
			AddErrorMessage("Du forsøker å legge til en dag som faller utenfor en vikar-tilknytning.<br> utvid den foregående vikar-tilknytningen istedetfor.")
			call RenderErrorMessage()					
		else
			strOppdragVikarId = RS.fields("OppdragVikarId").value
			rs.close
			set rs = nothing
		end if

		strSQL = "INSERT INTO DAGSLISTE_VIKAR (TimelisteVikarStatus, Dato, Starttid, Sluttid, OppdragID, VikarID,  AntTimer, " &_
			"TimeLonn, FirmaID, Fakturapris, Fakturastatus, BestilltAv, SOBestilltAv, Notat, Fakturatimer, LoennsArt, Lunch, Loennstatus, splittuke) " &_
			"values (" &_
			"1," &_
			strDato & "," &_
			"'" & strStarttid & "'," &_
			"'" & strSluttid & "'," &_
			strOppdragID & "," &_
			strVikarID  & "," &_
			strAntTimer & "," &_
			strTimeLonn & "," &_
			strFirmaID & "," &_
			strFakturapris & "," &_
			strFakturastatus & "," &_
			strBestilltAV & "," &_
			strSOBestilltAv & ",'" &_
			PadQuotes(strNotat) & "'," &_
			strFakturatimer & "," &_
			strLoennsArt & ",'" &_
			strLunch & "'," &_
			strloennstatus &"," &_
			splittverdi & ")"

	   if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under oppretting av timeliste.")
			call RenderErrorMessage()		
	   end if

 End If 'update or insert

	' Update status in DAGSLISTE_VIKAR    (timelisten)
	strSQL = "UPDATE DAGSLISTE_VIKAR SET " &_
     "TimelisteVikarStatus = 1" &_
	" WHERE VikarID = " & strVikarID &_
	" AND OppdragID = " & strOppdragID &_
	" AND Dato >= " & DbDate(session("startdato")) &_
	" AND Dato <= " & Dbdate(session("sluttdato"))


 End If 'fields have content


	'Push another page
	strStartdato = session("Startdato")


	If Request("kode") = 1 Then
		session("splittopt2")=session("splittopt")

		reDir = "Vikar_timeliste_vis3.asp?VikarID=" & strVikarID &_
			"&OppdragID=" & strOppdragID &_
			"&Startdato=" & strStartdato &_
			"&splittopt2="& Session("splittopt2") &_
			"&splittuke=" & session("splittuke") &_
			"&splittopt=" & session("splittopt")
	End If
'Response.End
	Response.Redirect reDir

%>