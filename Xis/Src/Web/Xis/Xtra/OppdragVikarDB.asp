<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<%


	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	Sub fjernKomma(strString )
		pos = Instr(strString, ",")
		If pos <> 0 Then
			mellom = Left(strString, pos-1) & "." & Mid(strString, pos+1)
			strString = mellom
		End If
	End Sub

	dim Conn
	dim ConnTrans	
	dim strSQL

	' Action against database depends on button pressed
	If Trim( Request("pbnDataAction") ) = "Nullstill" Then
	   ' Restart main function without key value to clear all fields
	   Response.redirect "OppdragVikarNy.asp"
	End If
	
	' Check datavalues
	If Request.Form("tbxOppdragID") = "" Then
		AddErrorMessage("Feil: Mangler oppdragid.")
		call RenderErrorMessage()	
	End If

	If Request.Form("tbxVikarID") = "" Then
		AddErrorMessage("Feil: Mangler vikarid.")
		call RenderErrorMessage()	
	End If

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	



	'Check oppdrag status and comment change to log activity
	Dim rsStatus
	Dim bChangedStatus
	Dim bCommentChange
	strSQL = "SELECT StatusID,ISNULL(Notat,'') AS Notat FROM [oppdrag_Vikar] WHERE [oppdragVikarid] = " & Request.form("tbxOppdragVikarID")
	set rsStatus = GetFirehoseRS(strSQL, Conn)
        
	' Any records found ?
	If HasRows(rsStatus) = false Then     
		Set rsStatus = Nothing		
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Failed getting Oppdragvikar previous status")
		call RenderErrorMessage()		
	End If
	
	bChangedStatus = false
	If ( cint(rsStatus("StatusID")) <> cint(Request.Form( "dbxStatus" )) ) then
	 	bChangedStatus = true
	end if
	'if comment is changed
	bCommentChange = true
	if ( 0=StrComp(rsStatus("Notat"), Request.form("tbxNotat"), 1 ) ) then
		bCommentChange = false
	end if 
	rsStatus.Close
	Set rsStatus = Nothing
	
	
	' Delete
	If Trim( Request("pbnDataAction") ) = "Slette"  AND Request.form("tbxOppdragVikarID") <> ""  Then
	   ' Create Sql-Statement
	   strSQL = "DELETE [oppdrag_Vikar] WHERE [oppdragVikarid] = " & Request.form("tbxOppdragVikarID")
	   ' Delete oppdrag vikar information
	   if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under sletting av oppdragvikar informasjon.")
			call RenderErrorMessage()		
	   end if

	   strSQL = "UPDATE [Oppdrag] SET [Tildato] = (SELECT MAX(Tildato) FROM [Oppdrag_Vikar] WHERE [OppdragID] = " & Request("tbxOppdragID") & ") WHERE [OppdragID] = " & Request("tbxOppdragID")
	   if (ExecuteCRUDSQL(strSQL, Conn) = false) then
			CloseConnection(Conn)
			set Conn = nothing	   
			AddErrorMessage("Feil under oppdatering av oppdrag tildato.")
			call RenderErrorMessage()		
	   end if
	   Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & Request.form("tbxOppdragID")
	End If

	' Is this update for FORKORT ?
	If Trim( Request("pbnDataAction")) = "Lagre" AND Request.form("tbxOppdragVikarID") <> "" AND Request.form("tbxAksjon") = "FORKORT" Then

		' Check on all input values
		If len(Request.Form("tbxFradato")) = 0 Then
			AddErrorMessage("Fradato mangler.")
		End If

		If len(Request.Form("tbxTildato")) = 0 Then
			AddErrorMessage("Tildato mangler.")	
		End If

		if(HasError() = true) then
			CloseConnection(Conn)
			set Conn = nothing	
			call RenderErrorMessage()
		end if
		
		' Check if new date is correct
		strSQL = "SELECT Fradato, Tildato from OPPDRAG_VIKAR OV " &_
			" WHERE OV.OppdragVikarID = " & Request.form( "tbxOppdragVikarID" ) &_
			" AND OV.fradato <= " & DbDate( Request.form( "tbxFraDato" ) ) &_
			" AND OV.tildato >= " & DbDate( Request.form( "tbxFraDato" ) ) &_
			" AND OV.fradato <= " & DbDate( Request.form( "tbxTilDato" ) ) &_
			" AND OV.tildato >= " & DbDate( Request.form( "tbxTilDato" ) )
		    
		set rsOppdragVikar = GetFirehoseRS(strSQL, Conn)
		

		' Any records found ?
		If HasRows(rsOppdragVikar) = false Then     
			Set rsOppdragVikar = Nothing		
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Forkortet dato må være mindre enn eksisterende tildato. Forkortelse ikke utført!")
			call RenderErrorMessage()		
		End If

		rsOppdragVikar.Close
		Set rsOppdragVikar = Nothing

		' Check if existing TIMELISTE allready have been treated
		strSQL = "SELECT DISTINCT DV.OppdragvikarID  from DAGSLISTE_VIKAR DV " &_
			" WHERE DV.OppdragVikarID = " & Request.form("tbxOppdragVikarID") &_
			" AND DV.TimelistevikarStatus > 1 " &_
			" AND ( DV.Dato < " & DbDate( Request.form("tbxFraDato") ) &_
			" OR DV.Dato > " & DbDate( Request.form("tbxTilDato") ) & " ) "

		set rsDagslisteVikar = GetFirehoseRS(strSQL, Conn)
	

		' Any records found ?
		If  HasRows(rsDagslisteVikar) = true Then
			rsDagslisteVikar.Close
			Set rsDagslisteVikar = Nothing		
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Det eksisterer timelister i denne perioden som allerede er behandlet. Forkortelse ikke utført!")
			call RenderErrorMessage()		
		End If

		rsDagslisteVikar.Close
		Set rsDagslisteVikar = Nothing
		
		'if transfertemps is on ge the required data before transaction begin
		
		dim   rsVikarDate
  if Request.form("TransferTemps")= "on"  then 
		    
		    
		    	strSQL = "Select OppdragID,  vikarid,dato,splittuke,Loennstatus,TimelisteVikarStatus,Fakturastatus, SoBestilltAv,Bestilltav FROM dagsliste_vikar " &_
			" WHERE OppdragVikarID = " & Request.form("tbxOppdragVikarID") &_
			" AND Dato = " & DbDate( Request.form("tbxTilDato") ) 
            
	      	set  rsVikarDate =  GetFirehoseRS(strSQL, Conn)
	      	
	       
		' Any records found ?
		If HasRows(rsVikarDate) =false Then     
			Set rsVikarDate = Nothing		
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Ingen registreringer funnet for overføring dato 1!")
			call RenderErrorMessage()		
		End If
		
		if (rsVikarDate("TimelisteVikarStatus")<> 1) then
		  rsVikarDate.Close
		  Set rsVikarDate = Nothing
		  AddErrorMessage("Timelistene er allerede godkjent til lønn/faktura. Opprett et eget rekrutteringsoppdrag.")
			call RenderErrorMessage()	
		End If
		

 
		 
	 
End IF

		Call fjernKomma(oppdrFaktor)
		Call fjernKomma(timepris)
		Call fjernKomma(timeloenn)

		' Start transaction
		Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))	
		ConnTrans.BeginTrans

		' Update oppdrag in database
		strSQL = "UPDATE oppdrag_vikar SET " &_
			"FraDato = " & DbDate( Request.form("tbxFraDato") ) & "," & _
			"TilDato = " & DbDate( Request.form("tbxTilDato") ) & _
			" WHERE oppdragVikarid =" & Request.form("tbxOppdragVikarID")

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under oppdatering av oppdrag vikar informasjon.")
			call RenderErrorMessage()
		End if

		' Update oppdrag fradato from oppdrag_vikar when status is aksept
		strSQL = "UPDATE oppdrag SET fradato = " &_
            "(SELECT MIN(fradato) FROM oppdrag_vikar WHERE oppdragid = " &_
            Request.form("tbxOppdragID") & " AND statusID = 4 ) " &_
            " WHERE oppdragid = " & Request.form("tbxOppdragID")
      
		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under oppdatering av fradato for oppdrag.")
			call RenderErrorMessage()
		End if

       strSQL = "UPDATE oppdrag SET tildato = " &_
			" (SELECT MAX(tildato) FROM oppdrag_vikar WHERE oppdragid =" &_
			Request.form("tbxOppdragID") & " AND statusId = 4 ) " &_
			" WHERE oppdragid = " & Request.form("tbxOppdragID") 

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under oppdatering av tildato for oppdrag.")
			call RenderErrorMessage()
		End if

		' Remove timelister on FORKORT
		strSQL = "DELETE FROM dagsliste_vikar " &_
			" WHERE OppdragVikarID = " & Request.form("tbxOppdragVikarID") &_
			" AND TimelistevikarStatus <= 1 " &_
			" AND ( Dato < " & DbDate( Request.form("tbxFraDato") ) &_
			" OR Dato > " & DbDate( Request.form("tbxTilDato") ) & ")"

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under sletting av timelister.")
			call RenderErrorMessage()
		End if
		
		
		if Request.form("TransferTemps")= "on"  then 
		  
		
		 
		 
		dim splitNumber
		dim transferAmount
        splitNumber = rsVikarDate("splittuke")
 
          if IsNull(splitNumber) then
               splitNumber = 0
          end if 
          
          dim weekNumber
          dim contact
          dim soContact
          
          If(ISNULL(rsVikarDate("Bestilltav")) OR rsVikarDate("Bestilltav") = 0 ) then
		contact = "NULL"
	  Else
		contact = rsVikarDate("Bestilltav")
	  End If

	  If(ISNULL(rsVikarDate("SoBestilltAv")) OR rsVikarDate("SoBestilltAv") = 0 ) then
		soContact = "NULL"
	  Else
		soContact = rsVikarDate("SoBestilltAv")
	  End If
          
          weekNumber =WeekFix(rsVikarDate("dato"))  
          if(weekNumber <10 ) then
               weekNumber =    Year(rsVikarDate("dato")) & 0 & weekNumber  
           
          else
                weekNumber =   Year(rsVikarDate("dato")) & weekNumber 
             
          end if
          
          
         transferAmount = Request.Form("tbxAmount")

		If Instr( transferAmount , "," ) > 0 Then
			transferAmount = Left( transferAmount, Instr( transferAmount , "," )-1) & "." & Mid( transferAmount, Instr( transferAmount , "," ) + 1  )
		End If
		
		
		 dim strValueAdditionID
		
		 strValueAdditionID = Request.Form("tbxAdditionID")

          
			strSQL = "EXECUTE [spECAddAddition] " & rsVikarDate("OppdragID") &_
				", " & rsVikarDate("vikarid") &_
				", " & weekNumber	&_
				", " & splitNumber &_
				", " & 6 &_
				", " & 0 &_
				", " & 0 &_
				", " & 0 & _
				", " & 1 &_
				", " & transferAmount  &_
				", " & transferAmount &_
				", " & rsVikarDate( "Loennstatus" ) &_
				", " & rsVikarDate( "Fakturastatus" ) &_
				", " & rsVikarDate( "TimelisteVikarStatus" )  &_
				", " & Quote( PadQuotes( Request.form("tbxComment"))) &_
				", " & contact  &_
				", " & soContact &_
				", " & 0 
				
           
 			If strValueAdditionID = 0 Then
				If ExecuteCRUDSQL(strSQL, ConnTrans) = false then		
					ConnTrans.RollbackTrans
					CloseConnection(ConnTrans)
					set ConnTrans = nothing
					AddErrorMessage("Ikke lagre filer.")
					call RenderErrorMessage()
				End if
			End if
			
		  	rsVikarDate.Close
		Set rsVikarDate = Nothing
		
	

		
		end if 
		
		
   
   		'Register activity for commision shortened - FORKORT
   		Dim  strActivity
		Dim rsActivityType
		Dim strComment
		Dim nActivityTypeID
		
		strActivity = "Oppdrag forkort"
		
		set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", ConnTrans)
		nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
		' Close and release recordset
	      	rsActivityType.Close
	      	Set rsActivityType = Nothing
				
		strComment = "Oppdraget forkortet. Dato: " & Request.form("tbxFraDato") & " - " & Request.form("tbxTilDato")
	      		      	
	      	sDate  = GetDateNowString()	      		      	
		strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
			"Values(" & nActivityTypeID & ",'" & sDate & "'," & Request.Form("tbxVikarID") & "," & Request.form("tbxFirmaID") & "," & Request.Form("tbxOppdragID") & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
		
		If ExecuteCRUDSQL(strSql, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Aktivitetsregistrering for forkorting av oppdrag feilet.")
			call RenderErrorMessage()
		End if
					
		' Commit transaction
		ConnTrans.CommitTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing

		Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & Request.Form("tbxOppdragID")
	End If

	' Check on all input values
	If Request.Form("tbxFradato") = "" Then
		AddErrorMessage("Fradato mangler.")
	End If

	If Request.Form("tbxTildato") = "" Then
		AddErrorMessage("Tildato mangler.")	
	End If

'  Check on antalltimer
	If Request.Form("tbxTimerPrDag") = "" Then
		AddErrorMessage("Timer pr dag mangler.")	
	End If

	if(HasError() = true) then
		CloseConnection(Conn)
		set Conn = nothing	
		call RenderErrorMessage()
	end if

	
	'  Check on TIMEPRIS
	If Request.Form("tbxTimePris") = "" Or  Request.Form("tbxTimePris") = "0" Then
		TimePris = 0
	Else
		TimePris = Request.Form("tbxTimePris")
		' Convert from ',' to '.' in antTimer
		If Instr( TimePris , "," ) > 0 Then
			TimePris = Left( TimePris, Instr( TimePris , "," )-1) & "." & Mid( TimePris, Instr( TimePris , "," ) + 1  )
		End If
	End If
		
	If Len(TimePris) = 8 Then
		TimePris = Left(TimePris, 1) & Mid(TimePris,3)
	End If	
	


	'  Check on timeloenn
	If Request.Form("tbxTimeloenn") = ""  Or  Request.Form("tbxTimeLoenn") = "0" Then
		Timeloenn = 0
	Else
		timeloenn = Request.Form("tbxTimeloenn")
		' Convert from ',' to '.' in timeloenn
		If Instr( timeloenn , "," ) > 0 Then
			timeloenn = Left( timeloenn, Instr( timeloenn , "," )-1) & "." & Mid( timeloenn, Instr( timeloenn , "," ) + 1  )
		End If
	End If
	
	If Len(timeloenn) = 8 Then
		timeloenn = Left(timeloenn, 1) & Mid(timeloenn,3)
	End If	


	'  Check on faktor
	If Request.Form("tbxFaktor") = ""  Then
		oppdrFaktor = 0
	else
		oppdrFaktor = Request.Form("tbxFaktor")
	End If

	'set kurskode
	' Commented by TPH - Enhancement No - 51455 
	'KursKode = Request.Form("tbxKurskode")

	' Is status Reservert or Aksept ?
	
	'If Request.Form( "dbxStatus" ) = 4 Or Request.Form( "dbxStatus" ) =  5  Then
'
'		' Check if VIKAR is available
'		' Add selection on status (reservert/akseptert)
'		strSQL = "SELECT OV.OppdragID FROM OPPDRAG_VIKAR OV, OPPDRAG O WHERE OV.VikarId = " & Request.form("tbxVikarID") & " AND " &_
'				"OV.OppdragID <> " & Request.form("tbxOppdragID") & " AND  " &_
'				"OV.OppdragID = O. OPPDRAGID AND " &_
'				"OV.StatusID IN ( 4, 5 ) AND " &_
'				"OV.tildato >= " & DbDate( Request.form("tbxFradato") ) & " AND " &_
'				"OV.fradato <=" & DbDate( Request.form("tbxTildato") )
'
'		' If kurskode has value check if free
'		If Clng( Kurskode ) = 1 Then
'			strSQL = strSQL & " AND O.Kurskode In (1,3) "
'		ElseIf Clng( Kurskode ) = 2 Then
'			strSQL = strSQL & " AND O.Kurskode in (2,3) "
'		End If
'
'		set rsOppdragVikar = GetFirehoseRS(strSQL, Conn)
'
'		' Any records found ?
'		If  HasRows(rsOppdragVikar) = true Then
'			AddErrorMessage("Vikaren er ikke tilgjengelig i denne perioden. Se oppdragnr:  " & rsOppdragVikar("OppdragID"))	
'			rsOppdragVikar.Close
'			set rsOppdragVikar = Nothing
'			CloseConnection(Conn)
'			set Conn = nothing
'			call RenderErrorMessage()
'		End If
'
'		rsOppdragVikar.Close
'		set rsOppdragVikar = Nothing
'
'	End If

	' Save New ?
	If Trim( Request("pbnDataAction")) = "Lagre" AND ( Request.form("tbxAksjon") = "UTVID" Or Request.form("tbxOppdragVikarID") = "" ) Then

		TimerPrDag = Request.Form("tbxTimerPrDag")

		If TimerPrDag = "" Then
			TimerPrDag = 0
		End If

		' Convert from , to . in antTimer
		If Instr( TimerPrDag, "," ) > 0 Then
			' Convert from , to . in antTimer
			TimerPrDag = Left( TimerPrDag, Instr( TimerPrDag , "," )-1) & "." & Mid( TimerPrDag, Instr( TimerPrDag , "," )+1  )
		End If

		Call fjernKomma(oppdrFaktor)

		' set flag on utvid
		If  Request.form("tbxAksjon") = "UTVID" Then
			Utvid = 1
		Else
			Utvid = 0
		End If

		' test for å se om nye linjer kan opprettes for oppdraget. hvis timelistestatus > 4 kan ikke dette tillates, pga at endringer ikke kan godtas i godkjente uker
		' i slike tilfeller må timelisten først nedgraderes på foreskrevet måte.
		strSQL = " DECLARE @teller INT " &_
			" EXEC @teller = utvidOppdrOk  " &_
			  DbDate(Request.form("tbxFraDato")) & "," &_
			  Request.form("tbxOppdragID") & "," &_
			  Request.form("tbxVikarID") &_
			" SELECT teller = @teller "

		set rs = GetFirehoseRS(strSQL, Conn)

		if (HasRows(rs)) then
			if rs("teller") > 0 THEN
				rs.close
				set rs = nothing			
				CloseConnection(Conn)
				set Conn = nothing				
				AddErrorMessage("Du kan ikke utvide oppdraget da det ligger godkjente timelister i uken med fradato.")	
				call RenderErrorMessage()			
			end if
		end if

		rs.close
		set rs = nothing
		
		
		
		

		' Start transaction
		Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))	
		ConnTrans.BeginTrans

		' create new oppdrag_vikar in database
		strSQL = "INSERT INTO Oppdrag_Vikar( StatusID, Fradato, Tildato, frakl, tilkl, OppdragID, VikarID, Lunch, AntTimer, Timeloenn, Timepris, Faktor, Timeliste, FirmaId, Notat, direkteTelefon, jobbEpost, utvid ,CategoryId) " & _
			"Values(" &_
			Request.form("dbxStatus") & "," & _
			DbDate( Request.form("tbxFraDato") ) & ", " & _
			DbDate( Request.form("tbxTilDato") ) & ", " & _
			DbTime( Request.form("tbxFraKl") ) & ", " & _
			DbTime( Request.form("tbxTilKl") ) & ", " & _
			Request.form("tbxOppdragID") & ", " & _
			Request.form("tbxVikarID") & ", " & _
			DbTime( Request.form("tbxLunsj") ) & ", " & _
			TimerPrDag & ", " & _
			Replace(Timeloenn,",",".") & ", " & _
			Replace(Timepris,",",".") & ", " & _
			Replace(oppdrFaktor,",",".") & ", " & _
			"0" & ", " & _
			Request.form("tbxFirmaID") & ", " & _
			Quote( PadQuotes(Request.form("tbxNotat")) )  & ", " &_
			Quote( Request.form("tbxDirTlf") ) & ", " &_
			Quote( Request.form("tbxWorkEmail") ) & ", " &_
			Utvid & ","
			'If (Request.form("tbxFaId") <> "0") Then
			If ( Request.Form("dbxCategory") = 0 ) Then				
				strSQL = strSQL + " Null )"
			Else
				strSQL = strSQL + Request.form("dbxCategory") & ")"		
			End If

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under utvidelse av oppdrag vikar.")
			call RenderErrorMessage()
		End if

		' Update oppdrag fradato from oppdrag_vikar when status is aksept
		If  Request.form("dbxStatus") = 4 Then
			
			strSQL = "UPDATE oppdrag SET fradato = " &_
				" (SELECT MIN(fradato) FROM oppdrag_vikar WHERE oppdragid = " &_
				Request.form("tbxOppdragID") & " AND statusID = 4 ) " &_
				" WHERE oppdragid = " & Request.form("tbxOppdragID")
			
			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av fradato på oppdrag.")
				call RenderErrorMessage()
			End if

			strSQL = "UPDATE oppdrag SET tildato = " &_
			" (SELECT MAX(tildato) FROM oppdrag_vikar WHERE oppdragid = " &_
			Request.form("tbxOppdragID") & " AND statusId = 4 ) " &_
			" WHERE oppdragid = " & Request.form("tbxOppdragID")
			
			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av tildato på oppdrag.")
				call RenderErrorMessage()
			End if
		End If
		
		If Request.Form( "dbxSOKontaktP" ) > 0 Then
			strSQL = "UPDATE oppdrag SET SOPeID = " & Request.Form( "dbxSOKontaktP" )  & " WHERE oppdragid = " & Request.form("tbxOppdragID")
			
			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av SOPeID på oppdrag.")
				call RenderErrorMessage()
			End if
			
		End If
		
		If utvid = 1 Then
			'Register activity for commision extend - utvid
	   					
			strActivity = "Oppdrag utvid"
			
			set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", ConnTrans)
			nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
			' Close and release recordset
		      	rsActivityType.Close
		      	Set rsActivityType = Nothing
					
		      	Timeloenn = Replace(Timeloenn,".",",")
					
			strComment = "Oppdraget forlenget. Dato: " & Request.form("tbxFraDato") & " - " & Request.form("tbxTilDato") & " Lønn: " & FormatNumber(Timeloenn,2)
		      		      	
		      	sDate  = GetDateNowString()	      		      	
			strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
				"Values(" & nActivityTypeID & ",'" & sDate & "'," & Request.Form("tbxVikarID") & "," & Request.form("tbxFirmaID") & "," & Request.Form("tbxOppdragID") & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
			
			If ExecuteCRUDSQL(strSql, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Aktivitetsregistrering for forkorte/forlenge oppdrag feilet.")
				call RenderErrorMessage()
			End if
		End If

		' Commit transaction
		ConnTrans.CommitTrans

		If utvid = 0 Then
			Response.redirect "Kalender.asp?VikarID=" & Request.form("tbxVikarID") & "&OppdragID=" & Request.Form("tbxOppdragID")
		End If

	' Is this normal update ?
	ElseIf Trim( Request("pbnDataAction")) = "Lagre" AND Request.form("tbxOppdragVikarID") <> "" AND Request.form("tbxAksjon") <> "FORKORT" Then

		' Check on Timerprdag/anttimer
		TimerPrDag = Request.Form("tbxTimerPrDag")
		' Convert from , to . in antTimer
		If Instr( TimerPrDag, "," ) > 0 Then
			' Convert from , to . in antTimer
			TimerPrDag = Left( TimerPrDag, Instr( TimerPrDag , "," )-1) & "." & Mid( TimerPrDag, Instr( TimerPrDag , "," )+1  )
		End If

		Call fjernKomma(oppdrFaktor)
		Call fjernKomma(timepris)
		Call fjernKomma(timeloenn)

		' Start transaction
		Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))	
		ConnTrans.BeginTrans

		' Update oppdrag in database
		strSQL = "UPDATE oppdrag_vikar SET " &_
		"StatusID = " & Request.form("dbxStatus") & "," & _
		"FraDato = " & DbDate( Request.form("tbxFraDato") ) & "," & _
		"TilDato = " & DbDate( Request.form("tbxTilDato") ) & "," & _
		"FraKl = " & DbTime( Request.form("tbxFrakl") ) & "," & _
		"TilKl = " & DbTime( Request.form("tbxTilkl") ) & "," & _
		"Lunch = " & DbTime( Request.form("tbxLunsj") ) & "," & _
		"AntTimer = " & TimerPrDag & "," & _
		"Timeloenn = " & Replace(Timeloenn,",",".") & "," & _
		"Timepris = " & Replace(Timepris,",",".") & "," & _
		"Faktor = " &  Replace(oppdrFaktor,",",".") & "," & _
		"OppdragID = " & Request.form("tbxOppdragID") & "," & _
		"VikarID = " & Request.form("tbxVikarID") & "," & _
		"Notat = " & Quote( PadQuotes(Request.form("tbxNotat")) )& ","  &_
		"jobbEpost = " & Quote( Request.form("tbxWorkEmail") ) & "," &_
		"direkteTelefon = " & Quote( Request.form("tbxDirTlf") )& " , "
		if Request.Form("dbxCategory") = 0 then
			strSQL = strSQL +   "CategoryId = Null WHERE oppdragVikarid = " & Request.form("tbxOppdragVikarID")
		else
			strSQL = strSQL +  "CategoryId = "& Request.Form("dbxCategory") &" WHERE oppdragVikarid = " & Request.form("tbxOppdragVikarID")
		End If
		 

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under oppdatering av oppdrag vikar informasjon.")
			call RenderErrorMessage()
		End if
			
		'set activity type
		strActivity = "Tilkn. vikar status"
		strSQL = "SELECT Status FROM [H_OPPDRAG_VIKAR_STATUS] WHERE [OppdragVikarStatusID] = " & Request.form("dbxStatus")
		set rsStatus = GetFirehoseRS(strSQL, Conn)
		strComment = "Tilknyttet vikar status endret til " & rsStatus("Status")
		rsStatus.Close
		Set rsStatus = Nothing
				
		' Update oppdrag fradato from oppdrag_vikar when status is aksept
		If  Request.form("dbxStatus") = 4 Then

			strSQL = "UPDATE oppdrag SET fradato = " &_
			" (SELECT MIN(fradato) FROM oppdrag_vikar WHERE oppdragid =" &_
			Request.form("tbxOppdragID") & " AND statusID = 4 ) " &_
			" WHERE oppdragid = " & Request.form("tbxOppdragID") 
			
			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av oppdragsinformasjon.")
				call RenderErrorMessage()
			End if

			strSQL = "UPDATE oppdrag SET tildato = " &_
			" (SELECT MAX(tildato) FROM oppdrag_vikar WHERE oppdragid =" &_
			Request.form("tbxOppdragID") & " AND statusId = 4 ) " &_
			" WHERE oppdragid = " & Request.form("tbxOppdragID")
			
			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av oppdragsinformasjon.")
				call RenderErrorMessage()
			End if
			
			'set activity type - this is on hold.
			strActivity = "Vikar aksept"
			strComment = "Vikaren fikk status aksept på oppdraget"
			'strComment = "Vikar akseptert. Dato: " & Request.form("tbxFraDato") & " - " & Request.form("tbxTilDato") & " Lønn: " & FormatNumber(Timeloenn,2)
		End If
		
		'if only status and comment both changed
		if (bCommentChange) then
			strComment = strComment & " Notat: " & Request.form("tbxNotat")
		end if
		

		
		if bChangedStatus then
			'Register activity for Temp accept		
			
			set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", ConnTrans)
			nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
			' Close and release recordset
		      	rsActivityType.Close
		      	Set rsActivityType = Nothing
			      	
		      	sDate  = GetDateNowString()	      		      	
			strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
				"Values(" & nActivityTypeID & ",'" & sDate & "'," & Request.Form("tbxVikarID") & "," & Request.form("tbxFirmaID") & "," & Request.Form("tbxOppdragID") & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
			
			If ExecuteCRUDSQL(strSql, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Aktivitetsregistrering for bytte av status feilet.")
				call RenderErrorMessage()
			End if
		end if

		'if status changed also we dont log comment change activity
		if (bCommentChange and (not bChangedStatus)) then
			'set activity type
			strActivity = "Nytt/endret notat"
			strComment = Request.form("tbxNotat")
			
			set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", ConnTrans)
			nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
			' Close and release recordset
		      	rsActivityType.Close
		      	Set rsActivityType = Nothing
			      	
		      	sDate  = GetDateNowString()	      		      	
			strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
				"Values(" & nActivityTypeID & ",'" & sDate & "'," & Request.Form("tbxVikarID") & "," & Request.form("tbxFirmaID") & "," & Request.Form("tbxOppdragID") & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
			
			If ExecuteCRUDSQL(strSql, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Aktivitetsregistrering for endring av kommentar feilet.")
				call RenderErrorMessage()
			End if
		end if
		
		' Commit transaction
		ConnTrans.CommitTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing			

		' Set return value
		strOppdragVikarID = Request.form("tbxOppdragVikarID")
	Else
		AddErrorMessage("Ingen gyldige parametere.")
		call RenderErrorMessage()
	End If
	
	CloseConnection(Conn)
	set Conn = nothing		
	Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & Request.Form("tbxOppdragID")
%>