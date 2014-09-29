<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim Timeloenn 
	dim Timepris
	dim Frakl
	dim Tilkl
	dim Lunsj
	dim TimerPrDag 
	dim Faktor 
	dim StatusID
	dim VikarID
	dim OppdragID
	dim FirmaID
	dim strSQL
	dim Conn

	Timeloenn 	= request("tbxTimeloenn")
	Timepris 	= request("tbxTimepris")
	Frakl		= request("tbxFraKl")
	Tilkl 		= request("tbxTilKl")
	Lunsj 		= request("tbxLunsj")
	TimerPrDag 	= request("tbxTimerPrDag")
	Faktor 		= request("tbxFaktor")
	StatusID	= request("dbxStatus")
	VikarID 	= request.queryString("vikarID")
	OppdragID 	= request.queryString("oppdragID")
	FirmaID		= request("tbxFirmaID")
	
	If (Faktor = "") Then
		AddErrorMessage("Faktor er ikke utfylt.")
		call RenderErrorMessage()	
	End If

	If TimerPrDag = "" Then
      TimerPrDag = 0
   End If
   
   ' Convert from , to . in antTimer
   If Instr( TimerPrDag, "," ) > 0 Then
     ' Convert from , to . in antTimer
     TimerPrDag = Left( TimerPrDag, Instr( TimerPrDag , "," )-1) & "." & Mid( TimerPrDag, Instr( TimerPrDag , "," )+1  )
   End If

	Call fjernKomma(Faktor)

	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	
	
	Conn.BeginTrans
	for each x in Request.form("chkBx")

		FraDato = x
		TilDato = x

		strSQL = "INSERT INTO Oppdrag_Vikar( StatusID, Fradato, Tildato, frakl, tilkl, OppdragID, VikarID, Lunch, AntTimer, Timeloenn, Timepris, Faktor, Timeliste, FirmaId) " & _
			"VALUES(" &_
			StatusID & "," & _
			DbDate(FraDato)  & "," & _
			DbDate(TilDato)  & "," & _
			DbTime(Frakl)  & "," & _
			DbTime(Tilkl)  & "," & _
			OppdragID & "," & _
			VikarID & "," & _
			DbTime(Lunsj) & "," & _
			TimerPrDag & "," & _
			Timeloenn& "," & _
			Timepris & "," & _
			Faktor & "," & _
			"0" & "," & _
			FirmaID & ")"

		' Insert into database
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollbackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil oppstod under oppretting av timeliste.")
			call RenderErrorMessage()
		End if
	Next

	strSQL = "update oppdrag set tildato = " &_
		" (select max(tildato) from oppdrag_vikar where oppdragid ="&OppdragID &_
		" and statusId = 4 ) " &_
		" where oppdragid = "& OppdragID
		
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		Conn.RollbackTrans
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil oppstod under oppretting av timeliste.")
		call RenderErrorMessage()
	End if

	Conn.CommitTrans
	CloseConnection(Conn)
	set Conn = nothing
	Response.redirect "oppdragVis.asp?oppdragID="&OppdragID
%>