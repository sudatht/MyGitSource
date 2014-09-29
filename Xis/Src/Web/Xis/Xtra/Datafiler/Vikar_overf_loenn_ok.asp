<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="../includes/Library.inc"-->
<%

dim valgt_avd
dim conn
dim strSQL

	If Request.QueryString("avd") <> "" Then
		valgt_avd = CInt(Request.QueryString("avd"))
	else
		valgt_avd = 0
	End If

	'oppdatering av alle tabeller etter overføring til Hult & Lillevik
	' Open database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	strDato = dbDate(Date)
	periode = Request("periode")

	'	finne neste lønnsnummer
	strSQL = "SELECT maks = MAX(LoenNr) FROM Vikar_loen_variable"
	set rsMax = GetFirehoseRS(strSQL, conn)
	LNR = rsMax("maks")
	rsMax.Close
	set rsMax = nothing

	If LNR = "" or IsNull(LNR) Then LNR = 0

	'Finne riktig årstall for periode slik at dette også blir riktig i årsskiftet
	'Henter den seneste datoen fra timeliste-tabellen og bruker denne som år i periode
	strSQL = "SELECT senest = MAX(dv.dato) FROM dagsliste_vikar dv " &_
	" WHERE dv.loennstatus = 2" 

	set rsStoersteDato = GetFirehoseRS(strSQL, conn)
	periode = (DatePart("yyyy", rsStoersteDato("senest")) * 100) + periode
	rsStoersteDato.close
	set rsStoersteDato = nothing

	'Finne aktuelle vikarer(lønn)
	strSQL = "SELECT DISTINCT VikarID FROM Vikar_loen_variable" &_
		" WHERE overfor_loenn_status = 2" &_
		" AND Avdeling = " & valgt_avd &_
		" ORDER BY VikarID"
		
	Set rsLoenn = GetFirehoseRS(strSQL, conn)

	If hasROWS(rsLoenn) Then
		'Start transaction
		Conn.Begintrans
		'	loop
		do while not rsLoenn.EOF

			'sette inn lønnsnummer
			LNR = LNR + 1
			' VIKAR_LOEN_VARIABLE
			strSQL = "UPDATE VIKAR_LOEN_VARIABLE SET LoenNr = " & LNR &_
			", Overfor_Loenn_status = 3" &_
			", Loenndato = " & strDato &_
			", Loennperiode = " & periode &_
			" WHERE VikarID = " & rsLoenn("VikarID") &_
			" AND Overfor_Loenn_status = 2" &_
			" AND Avdeling = " & valgt_avd 
			
			if (ExecuteCRUDSQL(strSQL, Conn) = false) then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av lønnvariable.")
				call RenderErrorMessage()		
			end if
			
			' VIKAR_UKELISTE
			strSQL = "UPDATE VIKAR_UKELISTE SET LoenNr = " & LNR &_
			", Overfort_loenn_status = 3" &_
			", Loenndato = " & strDato &_
			", Loennperiode = " & periode &_
			" WHERE VikarID = " & rsLoenn("VikarID") &_
			" AND Overfort_loenn_status = 2" &_
			" AND oppdragID in (SELECT distinct o.oppdragID FROM oppdrag o, oppdrag_vikar ov" &_
				" WHERE o.oppdragID = ov.oppdragID" &_
				" AND o.avdelingID = " & valgt_avd &_
				" AND ov.vikarid = " & rsLoenn("VikarID") & ")"
			
			if (ExecuteCRUDSQL(strSQL, Conn) = false) then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av lønnvariable.")
				call RenderErrorMessage()		
			end if			
			
			' DAGSLISTE_VIKAR    (timelisten)

			strSQL = "UPDATE DAGSLISTE_VIKAR SET LoenNr = " & LNR &_
			", Loennstatus = 3" &_
			", Loenndato = " & strDato &_
			", Loennperiode = " & periode &_
			" WHERE VikarID = " & rsLoenn("VikarID") &_
			" AND Loennstatus = 2" &_
			" AND oppdragID IN (SELECT distinct o.oppdragID FROM oppdrag o, oppdrag_vikar ov" &_
				" WHERE o.oppdragID = ov.oppdragID" &_
				" AND o.avdelingID = " & valgt_avd &_
				" AND ov.vikarid = " & rsLoenn("VikarID") & ")"
			
			if (ExecuteCRUDSQL(strSQL, Conn) = false) then
				Conn.RollBackTrans
				CloseConnection(Conn)
				set Conn = nothing	   
				AddErrorMessage("Feil under oppdatering av lønnvariable.")
				call RenderErrorMessage()		
			end if			
			
			rsLoenn.MoveNext
		loop
		rsLoenn.Close
		Conn.CommitTrans
		Set rsLoenn = Nothing
	End If 'ingen rader
	CloseConnection(Conn)
	set Conn = nothing

' Push another page
Response.Redirect "Vikar_timeliste_list3.asp?viskode=2&avd=" & valgt_avd 
%>