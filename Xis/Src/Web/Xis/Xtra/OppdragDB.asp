<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.Settings.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

    Dim customerApprovalDefaultValue
    Dim commentstring 
	Dim strpart			
	Dim linenumber
	Dim length    
	Dim position
	dim Conn
	dim strSQL
	dim SObestilltAv
	dim bestilltAv
	
	'Variables used by IM Publish (IMP) user routines.
	dim ConIMP 'Connection
	dim rsUser 'Recordset
	dim iIMPId 'integer, IMP userid
	dim strUsername 'string
	dim strSubstitute 'string
	dim LVikarID 'as long
	dim LConsID 'as long
	Dim LOppdragID 'as long
	dim faId
	dim bQuestback
	dim ReportingContactID
	dim postedOppId	
	dim amount
	dim noofpersons
	dim noofopportunities
	dim parentAssignment
	dim isChildAssignment
	dim hourlyRate
	dim ComTerms

	'default value
	customerApprovalDefaultValue = Application("CustomerApprovalDefaultValue")

	'Constants from table "web_rettigheter"
	'Consultant
	const C_Cons_Right_oppdrag  = "6"
	const C_Cons_Right_Kurs = "7"
	const C_Cons_Right_timelister  = "1"
	const C_Cons_Right_faktura = "2"

	'Customer
	const C_Cust_Right_oppdrag  = "11"
	const C_Cust_Right_Kurs = "10"
	const C_Cust_Right_timelister  = "9"
	const C_Cust_Right_faktura = "8"

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))


	if len(trim(Request("OppdragID"))) > 0 then
		postedOppId = trim(Request("OppdragID"))
	elseif Request.QueryString("OppdragId") <> "" then
		postedOppId = Request.QueryString("OppdragId")
	else
		postedOppId = 0
	end if
	LOppdragID = postedOppId
	
	'response.write postedOppId
	'response.write "@@@@"
	
	' Copy the assignment id from which the copy was invoked as the parentassignment value
	parentAssignment = postedOppId
	
	dim rsCheckIsChildOppdrag
	strSQL = "select isnull(parentassignment,0) as parentassignment from oppdrag " & _			
		 "WHERE [Oppdrag].[oppdragid] = " & postedOppId			

	set rsCheckIsChildOppdrag = GetFirehoseRS(strSQL, conn)
	If (HasRows(rsCheckIsChildOppdrag) = false) Then
			
	Else		
		' If it return null keep the parentassignment value as it is
		If isnull(rsCheckIsChildOppdrag.fields("parentassignment").value) Then
			parentAssignment = parentAssignment
		Else
			assvalue = rsCheckIsChildOppdrag.fields("parentassignment").value
			' If it return 0 keep the parentassignment value as it is
			If (assvalue = 0) Then
				parentAssignment = parentAssignment
				isChildAssignment = "no"
			' If it return a value (meaning this is a copy from a copy) replace with the new id which is the original parentassignment id
			Else
				parentAssignment = assvalue
				isChildAssignment = "yes"
			End If			
			
		End If
		
		rsCheckIsChildOppdrag.close
	End if

	Set rsCheckIsChildOppdrag = Nothing
	
	'response.write parentAssignment
	'response.write "@@@@"
	

	if trim(Request.form("hdnNotRecuit"))="DISABLED" then
		ComTerms = trim(Request.form("hdnTerm"))
	else
		ComTerms = trim(Request.Form("TermsCombo"))
	end if
	

	
' Create timeliste

	If Trim( Request.Form("hdnJobAction")) = "Lag Timeliste" or Trim(Request.QueryString("HdnJobAction")) = "Lag Timeliste" Then
		dim rsExistsTimesheets
		dim blnTimeSheetExists
		dim strExistingConsultantComment
		dim strExistingCustomerComment

		strSQL = "SELECT max([oppdrag_vikar].[tildato]) as Maxval, [BekreftelseKonsulentTekst],[BekreftelseKundeTekst]  FROM [oppdrag_vikar] " & _
			"INNER JOIN [Oppdrag] ON [Oppdrag].[OppdragID] = [oppdrag_vikar].[OppdragID] " & _
			"WHERE [Oppdrag].[oppdragid] = " & LOppdragID & " " & _
			"AND [Timeliste] > 0 " & _
			"GROUP BY [BekreftelseKonsulentTekst],[BekreftelseKundeTekst] Order by [maxval] desc"

		set rsExistsTimesheets = GetFirehoseRS(strSQL, conn)

		if (HasRows(rsExistsTimesheets) = false) then
			blnTimeSheetExists = false
		else
			blnTimeSheetExists = true
			if isnull(rsExistsTimesheets.fields("BekreftelseKonsulentTekst").value) then
				strExistingConsultantComment = ""
			else
				strExistingConsultantComment = rsExistsTimesheets.fields("BekreftelseKonsulentTekst").value
			end if
			if isnull(rsExistsTimesheets.fields("BekreftelseKundeTekst").value) then
				strExistingCustomerComment = ""
			else
				strExistingCustomerComment = rsExistsTimesheets.fields("BekreftelseKundeTekst").value
			end if
			rsExistsTimesheets.close
		end if

		set rsExistsTimesheets = nothing

		' Get consultants with task that does not have timesheets.
		strSQL = "SELECT [OV].[OppdragVikarId], [OV].[OppdragID], [OV].[VikarID], [V].[TypeID], [O].[FirmaID], [OV].[Fradato], [OV].[Tildato], " &_
				"[OV].[Frakl], [OV].[TilKl], [OV].[AntTimer], [OV].[Timeloenn], [OV].[Timepris], [OV].[Lunch], [O].[bestilltAv], [O].[SOPeID] " &_
				" FROM [OPPDRAG_VIKAR] AS [OV], [VIKAR] AS [V], [Oppdrag] AS [O] " &_
				" WHERE " &_
				" [OV].[VikarID] = [V].[VikarID]" &_
				" AND [OV].[OppdragID] = [O].[OppdragID]" &_
				" AND [OV].[OppdragID] = " & LOppdragID &_
				" AND [OV].[StatusID] = 4 " &_
				" AND [OV].[Timeliste] = 0 "
				
		set rsOppdragVikar = GetFirehoseRS(strSQL, conn)

		' No records found ?
		If (HasRows(rsOppdragVikar) = false) Then
			set rsOppdragVikar = nothing
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Ingen vikarer med status Aksept som ikke har timelister funnet!")
			call RenderErrorMessage()
		End If

		Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))
		ConnTrans.BeginTrans

		' Create timesheets for all accepted consultants
		Do Until rsOppdragVikar.EOF

			' Check value in AntTimer
			TimerPrDag = rsOppdragVikar("AntTimer")

			If TimerPrDag = "" Then
				TimerPrDag = 0
			End If

			' Convert from ',' to '.' in antTimer
			If Instr( TimerPrDag, "," ) > 0 Then
				TimerPrDag = Left( TimerPrDag, Instr( TimerPrDag , "," )-1) & "." & Mid( TimerPrDag, Instr( TimerPrDag , "," )+1  )
			End If

			' Set correct VIKARID
			If rsOppdragVikar("TypeID") = 1 Then
				Vikartype = 1
			Else
				Vikartype = 0
			End If
			
			if(isnull(rsOppdragVikar("bestilltAv"))) then
				bestilltAv = "NULL"
			else
				bestilltAv = rsOppdragVikar("bestilltAv")
			end if

			if(isnull(rsOppdragVikar("SOPeID"))) then
				SObestilltAv = "NULL"
			else
				SObestilltAv = rsOppdragVikar("SOPeID")
			end if

			' Format time value
			Lunch = FormatDateTime( rsOppdragVikar( "Lunch" ), 3 )

			TimelisteVikarStatus = 1
			' Create sql-statement for Procedure Lag_timeliste
			strSQL = "EXECUTE [Lag_Timeliste] " & rsOppdragVikar("OppdragVikarID") &_
				", " & rsOppdragVikar("OppdragID") &_
				", " & rsOppdragVikar("VikarID")	&_
				", " & rsOppdragVikar("FirmaID") &_
				", " & DbDate( rsOppdragVikar("Fradato") ) &_
				", " & DbDate( rsOppdragVikar("TilDato") ) &_
				", " & DbTime( rsOppdragVikar("Frakl") ) &_
				", " & DbTime( rsOppdragVikar("Tilkl") ) & _
				", " & TimerPrDag  &_
				", " & DbTime( Lunch ) &_
				", " & rsOppdragVikar( "Timeloenn" ) &_
				", " & rsOppdragVikar( "Timepris" ) &_
				", " & VikarType &_
				", " & bestilltAv  &_
				", " & SObestilltAv  &_
				", " & TimelisteVikarStatus

			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppretting av timeliste.")
				call RenderErrorMessage()
			End if

			' Update oppdrag_vikar med Timeliste created
	   		if (blnTimeSheetExists = true) then
	   			strSQL = "Update [oppdrag_vikar] set [Timeliste] = 1, [BekreftelseKonsulentTekst]=" & Quote( PadQuotes(strExistingConsultantComment)) & ", [BekreftelseKundeTekst]=" & Quote(PadQuotes(strExistingCustomerComment)) & " WHERE oppdragvikarid =" & rsOppdragVikar("OppdragVikarID")
	   		elseif (blnTimeSheetExists = false) then
				strSQL = "Update [oppdrag_vikar] set [Timeliste] = 1, [BekreftelseKonsulentTekst]=" & Quote( PadQuotes(getSetting("TXT_oppdragBekreftelse_vikar"))) & ", [BekreftelseKundeTekst]=" & Quote( PadQuotes(getSetting("TXT_oppdragBekreftelse_Kunde"))) & " WHERE oppdragvikarid =" & rsOppdragVikar("OppdragVikarID")
	   		end if

			If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
				ConnTrans.RollbackTrans
				CloseConnection(ConnTrans)
				set ConnTrans = nothing
				AddErrorMessage("Feil oppstod under oppdatering av Timeliste data!")
				call RenderErrorMessage()
			End if

			' Get next OPPDRAG_VIKAR
			rsOppdragVikar.MoveNext
		Loop

		' Update status on OPPDRAG to FULLSTENDIG
		strSQL = "UPDATE [oppdrag] SET [StatusID] = 5 WHERE [oppdragid] =" & LOppdragID

		If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
			ConnTrans.RollbackTrans
			CloseConnection(ConnTrans)
			set ConnTrans = nothing
			AddErrorMessage("Feil oppstod under oppdatering av Timeliste data!")
			call RenderErrorMessage()
		End if

		ConnTrans.CommitTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing

		'Rights should only be given to new timelister
		'Fetch vikarid based on task id
		strSQL = "SELECT v.vikarid FROM Vikar AS v, Oppdrag_Vikar AS ov " &_
				"WHERE v.VikarID = ov.VikarID " &_
				"AND ov.statusid = 4 AND oppdragid = " & LOppdragID

		set rsConsultant = GetFirehoseRS(strSQL, Conn)
		Lvikarid = rsConsultant("vikarid").value
		rsConsultant.close
		set rsConsultant = nothing

		strSQL = "SELECT COUNT(vikarid) AS Hasrights FROM web_rettigheter_vikar WHERE oppdragid = " & LOppdragID & " AND vikarID = " & Lvikarid
		set rsConsultant = GetFirehoseRS(strSQL, Conn)

		if (rsConsultant.fields("Hasrights") = 0) then
		'Give consultant & customer contact default web rights if they have web usernames.

			'First fetch consultant with status "Aksept" for this task.
			Oppdragskode = 0

			'if consultant has a web username, give him default userrights for this task..
			'Open connection to DNN

			iApp = Cint(Application("Application"))
			sUserServiceURL = Application("DNNUserServiceURL")

			Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
			objUserProxy.Url = sUserServiceURL

			sUserXml = objUserProxy.GetUser(iApp, Lvikarid,"V")

			
			if (sUserXml   = "") then
				'insert web rights for consultant..
				StrSQL = "INSERT INTO web_rettigheter_vikar(vikarid, oppdragid, rettighetid) " &_
						"VALUES( " & lvikarid & "," & Loppdragid & ", "

				' Update Consultant rights in database
				if (Oppdragskode = 0) then
					If ExecuteCRUDSQL(strSQL & C_Cons_Right_oppdrag & ")", Conn) = false then
						CloseConnection(Conn)
						set Conn = nothing
						CloseConnection(ConIMP)
						set ConIMP = nothing						
						AddErrorMessage("Feil oppstod under oppdatering av vikar rettigheter!")
						call RenderErrorMessage()
					End if
				elseif (Oppdragskode = 1) then
					If ExecuteCRUDSQL(strSQL & C_Cons_Right_Kurs & ")", Conn) = false then
						CloseConnection(Conn)
						set Conn = nothing
						CloseConnection(ConIMP)
						set ConIMP = nothing						
						AddErrorMessage("Feil oppstod under oppdatering av vikar rettigheter!")
						call RenderErrorMessage()
					End if
				end if
				If ExecuteCRUDSQL(strSQL  & C_Cons_Right_faktura & ")", Conn) = false then
					CloseConnection(Conn)
					set Conn = nothing
					CloseConnection(ConIMP)
					set ConIMP = nothing					
					AddErrorMessage("Feil oppstod under oppdatering av vikar rettigheter!")
					call RenderErrorMessage()
				End if
			end if

		end if
    ' Return to show OPPDRAG
	'Response.redirect "WebUI\OppdragView.aspx?OppdragID=" & LOppdragID
	 
	 'Redirect to Automatially unpublish page 
 	 Response.redirect "WebUI\UnpublishAddAP.aspx?commid=" & LOppdragID 
End If
 
if Request.Form("rbnKurskode") = 3 then  'for direct recurutiment
 
  ' Validate input		 
  
  noofopportunities = 1  'This is for direct recruitment. This field is not used for direct recruitments

    If Request.Form("dbxSOKontaktP") = "0" Then
		AddErrorMessage("Vennligst velg en kontaktperson!")
	End If
	
    If lenb(Request.Form("tbxRecruitmentDate")) = 0 Then
		AddErrorMessage("Rekruttering Dato mangler!")
	End If
	
	If lenb(Request.Form("FirmaID")) = 0 Then
		AddErrorMessage("Kontaktnr mangler!")
	End If

	If lenb(Request.Form("tbxBeskrivelse")) = 0 Then
		AddErrorMessage("Beskrivelse mangler!")
	End If	
	
	' When the commision status is Status 5 --> Fullstendig, Kommentar field must be mandatory
	If Request.form("dbxStatus") = 5 Then		
		If len(trim(Request.Form("tbxComment"))) = 0 Then			
			AddErrorMessage("Kommentar på faktura mangler!")
		End If
	End If
	
	If Request.Form("dbxtjenesteomrade") = "Ingen" Then
		AddErrorMessage("Du må velge tjenesteområde!")
	End If


		TimePris = 0


	if request("cbxPubliser")="on" Then
		webPub = 1
		if lenb(trim(request("tbxNotatWeb"))) = 0 Then
			AddErrorMessage("Webbeskrivelse må være utfylt!")
		elseif lenb(trim(request("tbxWebOverskrift"))) = 0 Then
			AddErrorMessage("Weboverskrift må være utfylt!")
		end if
	else
		webPub = 0
	end if

	'  Check on TIMERPRDAG

		TimerPrDag = 0
	

	 timeloenn=0
 
	 



	Lunsj = Request.Form("tbxLunsj")

	'  Check course / job type
	If CInt( Request.Form("rbnkurskode") ) = 0 Then
		Oppdragskode = 0
		Kurskode = 0
	Else
		Kurskode = Request.Form("rbnKurskode")
		Oppdragskode = 2
	End If

	' sjekk om oppdrag skal publiseres på web (0=nei,1=ja)
	if request("cbxPubliser")="on" THEN
		webPub = 1
	else
		webPub = 0
	end if

	'check if "ansvarlig" is selected
	if CInt(Request.form("dbxMedarbeider")) = 1 Then
 		AddErrorMessage("Det må velges ansvarlig!")
	End If
	
	'Check ReportingOverrule
	If (Request.Form("questback") <> "") Then
		bQuestback = 1
		ReportingContactID = 0
	Else
		bQuestback = 0
		if Request.Form("dbxQuestbackP") = "0" then
			ReportingContactID = Request.form("dbxSOKontaktP")
		else		
		ReportingContactID = Request.form("dbxQuestbackP")
		end if
	End If
	
	 if(len(trim(Request.Form("txtAmount"))) = 0 ) then
		   AddErrorMessage("Fyll inn lønn beløp!")

	Else
		amount = Request.Form("txtAmount")
		 
		' Convert from ',' to '.' in antTimer
		If Instr( amount , "," ) > 0 Then
			amount = Left( amount, Instr( amount , "," )-1) & "." & Mid( amount, Instr( amount , "," ) + 1  )
		End If
 
	 End If
	 
	 if(len(trim(Request.Form("txtNoOfPersons"))) = 0 ) then
		   AddErrorMessage("Fyll inn antall personer!")
	 Else
		noofpersons = Request.Form("txtNoOfPersons")
		' Validate value as an integer
		If IsNumeric(noofpersons) = False  Then
			AddErrorMessage("Antall personer må være en tallverdi!")
		End If 
	 End If
	
	
	if(HasError() = true) then
		call RenderErrorMessage()
	end if
	
	
else   ' for oppdrags
	' Validate input
	If Request.Form("dbxSOKontaktP") = "0" Then
		AddErrorMessage("Vennligst velg en kontaktperson!")
	End If
			
	If lenb(Request.Form("tbxFradato")) = 0 Then
		AddErrorMessage("Fradato mangler!")
	End If

	If lenb(Request.Form("tbxTildato")) = 0 Then
		AddErrorMessage("Tildato mangler!")
	End If

	if (ValidDateInterval(ToDateFromDDMMYY(Request.Form("tbxFradato")), ToDateFromDDMMYY(Request.Form("tbxTildato"))) = false) then
		AddErrorMessage("Fradato kan ikke være senere enn tildato!")
	end if

	If lenb(Request.Form("FirmaID")) = 0 Then
		AddErrorMessage("Kontaktnr mangler!")
	End If

	If lenb(Request.Form("tbxBeskrivelse")) = 0 Then
		AddErrorMessage("Beskrivelse mangler!")
	End If
	
	response.write parentAssignment
	'klll
	
	If(len(trim(Request.Form("txtNoOfOpportunities"))) = 0 and (Trim(Request.Form("hdnJobAction")) <> "copy") and (parentAssignment = 0)) Then
	 		AddErrorMessage("Antall personer mangler!")
	Else		
		If (Trim(Request.Form("hdnJobAction")) = "copy") or (isChildAssignment = "yes") Then
			noofopportunities = 1
		Else		
			noofopportunities = Request.Form("txtNoOfOpportunities")
		End If
		
		' Validate value as an integer
		If IsNumeric(noofopportunities) = False  Then
			AddErrorMessage("Antall personer er numerisk!")
		End If 
	 End If
	
	If Request.Form("dbxtjenesteomrade") = "Ingen" Then
		AddErrorMessage("Du må velge tjenesteområde!")
	End If

	'  Check on TIMEPRIS
	If lenb(Request.Form("tbxTimePris")) = 0 Then
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
	
	if request("cbxPubliser")="on" Then
		webPub = 1
		if lenb(trim(request("tbxNotatWeb"))) = 0 Then
			AddErrorMessage("Webbeskrivelse må være utfylt!")
		elseif lenb(trim(request("tbxWebOverskrift"))) = 0 Then
			AddErrorMessage("Weboverskrift må være utfylt!")
		end if
	else
		webPub = 0
	end if

	'  Check on TIMERPRDAG
	If lenb(Request.Form("tbxTimerprdag")) = 0 Then
		TimerPrDag = 0
	Else
		TimerPrDag = Request.Form("tbxTimerprdag")

		' Convert from ',' to '.' in antTimer
		If Instr( TimerPrDag , "," ) > 0 Then
			TimerPrDag = Left( TimerPrDag, Instr( TimerPrDag , "," )-1) & "." & Mid( TimerPrDag, Instr( TimerPrDag , "," )+1  )
		End If
	End If

	'  Check on timeloenn
	If lenb(Request.Form("tbxTimeloenn")) = 0 Then
		AddErrorMessage("Timelønn mangler!")
	Else
		timeloenn = Request.Form("tbxTimeloenn")
		' Convert from ',' to '.' in timeloenn
		If Instr( timeloenn , "," ) > 0 Then
			timeloenn = Left( timeloenn, Instr( timeloenn , "," )-1) & "." & Mid( timeloenn, Instr( timeloenn , "," ) + 1  )
		End If
	End If

	' Check input on lunsj
	If lenb(Request.Form("tbxLunsj")) = 0 Then
		AddErrorMessage("Lunsj mangler!")
	End If

	Lunsj = Request.Form("tbxLunsj")

	'  Check course / job type
	If CInt( Request.Form("rbnkurskode") ) = 0 Then
		Oppdragskode = 0
		Kurskode = 0
	Else
		Kurskode = Request.Form("rbnKurskode")
		Oppdragskode = 2
	End If

	' sjekk om oppdrag skal publiseres på web (0=nei,1=ja)
	if request("cbxPubliser")="on" THEN
		webPub = 1
	else
		webPub = 0
	end if

	'check if "ansvarlig" is selected
	if CInt(Request.form("dbxMedarbeider")) = 1 Then
 		AddErrorMessage("Det må velges ansvarlig!")
	End If
	
	'Check ReportingOverrule
	If (Request.Form("questback") <> "") Then
		bQuestback = 1
		ReportingContactID = 0
	Else
		bQuestback = 0
		if Request.Form("dbxQuestbackP") = "0" then
			ReportingContactID = Request.form("dbxSOKontaktP")
		else		
		ReportingContactID = Request.form("dbxQuestbackP")
		end if
	End If
	
	
	 'if trim(Request.Form("TermsCombo")) = "1" then
	 if ComTerms = "1" then

    if(len(trim(Request.Form("txtAmount"))) = 0 ) then
		   AddErrorMessage("Fyll inn lønn beløp!")
		 
	Else
		amount = Request.Form("txtAmount")
 
		' Convert from ',' to '.' in antTimer
		If Instr( amount , "," ) > 0 Then
			amount = Left( amount, Instr( amount , "," )-1) & "." & Mid( amount, Instr( amount , "," ) + 1  )
		End If
	 
	 End If	 
	
	
	if(len(trim(Request.Form("txtHourlyRate"))) = 0 ) then
		   AddErrorMessage("Skriv Timesats!")
		 
	Else
		hourlyRate = Request.Form("txtHourlyRate")
		 
		' Convert from ',' to '.' in antTimer
		If Instr( hourlyRate , "," ) > 0 Then
			hourlyRate = Left( hourlyRate, Instr( hourlyRate , "," )-1) & "." & Mid( hourlyRate, Instr( hourlyRate , "," ) + 1  )
		end if
	 
	End If
	
if(HasError() = true) then
		call RenderErrorMessage()
	end if
  end if
end if
	

' Action against database depending on Button pressed
'New job / oppdrag
If Trim( Request.Form("hdnJobAction")) = "lagre" AND clng(LOppdragID) = 0 Then

	if cint(Request.Form("dbxAvdeling")) = 0 then
		AddErrorMessage("Avdeling er ikke utfylt!")
	end if

	if cint(Request.Form("dbxregnskap")) = 0 then
		AddErrorMessage("Regnskapsavdeling er ikke utfylt!")
	end if

	if Request.Form("dbxKontaktP") = "0" then
		AddErrorMessage("Du må velge en kontaktperson hos Kontakt!")
	end if
	
	if Request.Form("dbxFAgreement") = "-1" then
		AddErrorMessage("Du må velge om det skal føres eget kunderegnskap på dette oppdrag!")
	end if
		
	
	if Request.Form("dbxCategory") = "0" then
		AddErrorMessage("Velg en kategori!")		
	end if
		
	if(HasError() = true) then
		call RenderErrorMessage()
	end if
	   


   'Create new oppdrag in database
   strSQL = "INSERT INTO Oppdrag(IsCustomerApproval, StatusID, TypeID, AnsMedID, AvdelingskontorID, AvdelingID, tomID, CategoryID, "&_
            "beskrivelse, fradato, frakl, tildato, tilkl, firmaid, ArbAdresse, bestiltDato, bestilltAv, SOPeID, oppdragskode, " &_
            "kurskode, " &_
            "bestiltklokken, timepris, timerprdag, timeloenn, lunch, " &_
            "notatansvarlig, notatokonomi, InvoiceComments, CustomerReference, " &_
            "webSted, webBeskrivelse, Weboverskrift,ReportingContactID,ReportingOverrule,NoOfPersons,Terms,FaID, webPub)  " &_
            "Values(" &_
            customerApprovalDefaultValue & ", " &_
            Request.form("dbxStatus") & ", " &_
             "0, " &_
            Request.form("dbxMedarbeider") & ", " &_
			Request.form("dbxAvdeling") & ", " &_
			Request.Form("dbxregnskap") & ", " &_
			Request.form("dbxtjenesteomrade") & ", " &_
			Request.form("dbxCategory") & ", " &_
			Quote( Request.form("tbxBeskrivelse") ) & ", " &_
			DbDate( Request.form("tbxFraDato") ) & ", " &_
			DbTime( Request.form("tbxFraKl") ) & ", " &_
			DbDate( Request.form("tbxTilDato") ) & ", " &_
			DbTime( Request.form("tbxTilKl") ) & ", " &_
			Request.form("FirmaID") & ", " &_
			Quote( Request.form("tbxArbAdresse") ) & ", " & _
			DbDate( Request.form("tbxBestDato") ) & ", " & _
			"NULL , " & _
			Request.form("dbxSOKontaktP") & ", " & _		
			Oppdragskode & ", " & _
			Kurskode & ", " & _
			DbTime( Request.form("tbxBestKl") ) & ", " &_
			Timepris & ", " &_
			Timerprdag & ", " &_
			Quote(timeloenn) & ", " &_
			DbTime( Lunsj ) & ", " &_
			Quote( PadQuotes(Request.form("tbxNotatAnsvarlig")) ) & ", " &_
			Quote( PadQuotes(Replace((Replace(Request.form("tbxNotatOkonomi"),"<p>","")),"</p>","")) ) & ", " &_			
			Quote( PadQuotes(Request.form("tbxInvoiceComments")) ) & ", " &_
			Quote( PadQuotes(Request.form("tbxCustomerReference")) ) & ", " &_
			Quote( PadQuotes(request("txtSted")) ) & ", " &_
			Quote( PadQuotes(request("tbxNotatWeb")) ) & ", " &_
			Quote( PadQuotes(request("tbxWebOverskrift")) ) & ", " &_
			ReportingContactID & ", " & _
			bQuestback & ", " & _
			noOfOpportunities & ", " & _
			ComTerms& ", " 
	if Request.Form("dbxFAgreement") = 0 then
		strSQL = strSQL +   "NULL," & webPub & ")"
	else
		strSQL = strSQL +  Request.Form("dbxFAgreement")& "," & webPub & ")"
	End If
	
	'response.write strSQL
	
	'strSQL = strSQL + bQuestback & ")"		
 
   ' Insert into database
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil ved oppretting av oppdrag!")
		call RenderErrorMessage()
	End if

    ' Get new oppdragID
    strSQL = "SELECT NewOppdragID = MAX( OppdragID ) FROM Oppdrag"
    set rsOppdrag = GetFirehoseRS(strSQL, Conn)
    
    'if fixed salary employee 
    if   (Kurskode = 0 and trim(ComTerms) = "1") then
    
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,HourlyRate) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",0," & amount & " , "& hourlyRate & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
    end if 
    
   'if direct employee employee
          
    if   (Kurskode = 3) then 
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,noofpersons,RecruitmentDate,Comment) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",1," & amount & "," & noofpersons & " ," & DbDate(Request.Form("tbxRecruitmentDate")) & "," & Quote(PadQuotes(Request.Form("tbxComment"))) & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
		dim commisionStatus  
		commisionStatus= Request.form("dbxStatus")
	 
		 strSQL = "SELECT Oppdragsstatus FROM  H_OPPDRAG_STATUS WHERE  OppdragsstatusID = " & commisionStatus
                   set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
             
            
            
            IF rsOppdragStatus("Oppdragsstatus") = "Fullstendig" then
            
           
             
             strSQL = "SELECT InvMaxNo = MAX(Fakturanr) FROM FAKTURAGRUNNLAG" 
              set rsInvNumber  =  GetFirehoseRS(strSQL, Conn)
            
            ' Create sql-statement insert to Add : 1st line for Direct recruitment (Recruitment text with values)      
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_ 
				", " &  rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   1   &_
				", " &   amount  &_
				", " &  1 &_
				", " & "301099" &_
				", " & "Rekruttering" & _
				", " & amount   &_
				", " & "NULL" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & 1 &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				    CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
								
			' Create sql-statement insert to Add : 2nd - Multiple line for Direct recruitment (Comment text, only if it exists)			
			linenumber = 1     ''' default starting number
			commentstring = Trim(Request.Form("tbxComment"))
			length = Len(commentstring)
						
			Do While length>0
			  If length>60 then
			    'document.write("<br/>Length : " &length)    
			    If Mid(commentstring,60,1) = " " then  			
			      strpart = Trim(Left(commentstring,60))
			      commentstring = Trim(Right(commentstring,length-60))
			    Else 					 
			      strpart = Trim(Left(commentstring,60))			      
			      position = InStrRev(strpart," ")			      
			      strpart = Trim(Left(commentstring,position))
			      commentstring = Trim(Right(commentstring,length-position))
			    End if 
			    'document.write("<br/>strpart :"&strpart)
			    'document.write("<br/>commentstring :"&commentstring)
			    length = Len(commentstring)
			  Else
			    strpart = commentstring    
			    length = 0			
			  End if 
			
			  ''' Insert comment line in DB
			  linenumber = linenumber + 1
			  strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes(strpart)) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & linenumber &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")
       
	       		' Check whether the text has any values befor inserting	            
				If ExecuteCRUDSQL(strSQL, conn) = false then
					    CloseConnection(Conn)
				        set Conn = nothing
				        AddErrorMessage("Feil på beløpet feltet!")
				        call RenderErrorMessage()
				End if
			Loop	
			''' LOOP ENDS HERE
			
			' Create sql-statement insert to Add : Final line for Direct recruitment (Oppdrag No)
			linenumber = linenumber + 1
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes( "Oppdrag "  & rsOppdrag("NewOppdragID"))) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & linenumber &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")
       
      
			If ExecuteCRUDSQL(strSQL, conn) = false then
				     CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
			
            
           END IF
		 
			
			
		
    end if 

    ' Do we have a KURS-PROGRAM ?
    If Request.form("dbxProgram") <> 0 Then

       ' Insert program AND kompnivå in OPPDRAG_KOMPETANSE
      strSQL = "INSERT INTO OPPDRAG_KOMPETANSE" &_
			"(oppdragID, K_TypeID, K_TittelID, K_LevelID ) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ", 3, " &_
			Request.form("dbxProgram") & "," & Request.form("dbxKompniva") & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil ved oppretting av kursprogram!")
			call RenderErrorMessage()
		End if
		
		

    End If
 
    ' set return value
    strOppdragID = rsOppdrag("NewOppdragID")
    
 
ElseIf Trim( Request.Form("hdnJobAction")) = "lagre" AND clng(LOppdragID)> 0 Then
 
	'If this task doesn't have an xis (not SuperOffice) contact person, 
	'the xis contact person is mandatory.
	dim SOBestiltAv
	dim bestiltAv
	
	if (lenb(Request.Form("dbxKontaktP")) = 0 and Request.Form("dbxSOPeID") = 0) then 
		AddErrorMessage("Du må velge en kontaktperson hos Kontakt!")
		call RenderErrorMessage()
	end if	
	
	if Request.Form("dbxFAgreement") = "-1" then
		AddErrorMessage("Du må velge om det skal føres eget kunderegnskap på dette oppdrag!")
		call RenderErrorMessage()
	end if
	
	if Request.Form("dbxCategory") = "0" then
		AddErrorMessage("Velg en kategori!")
		call RenderErrorMessage()
	end if
	
	if (Request.Form("dbxSOPeID") > 0) then  
		response.write "#" & Request.Form("dbxSOPeID") & "#"
		If (Request.Form("dbxNoOfTimesheets") = 0) Then
			SOBestiltAv = Request.form("dbxSOKontaktP")
		Else
			SOBestiltAv = Request.Form("dbxSOPeID")
		End If		
		bestiltAv = "NULL"
	end if
	
	if(lenb(Request.Form("dbxKontaktP")) > 0) then 		
		SOBestiltAv  = "NULL"
		bestiltAv = Request.Form("dbxKontaktP")
	end if	

	if(lenb(Request.Form("dbxKontaktPerson")) > 0) then 
		ReportingContactID = "NULL"
	end if
	
	if(len(ReportingContactID) = 0) Then
		ReportingContactID = 0
	end if

	'if Request.Form("dbxFAgreement") > 0 then
	'	faId = Request.Form("dbxFAgreement")	
	'else
	'	faId = "NULL"	
	'End If
	
   ' Update oppdrag
   dim opdragStatus 
   
   if(Request.form("dbxStatus")="") Then
     opdragStatus=Request.form("tbxStatus")
   else
     opdragStatus=Request.form("dbxStatus")
   end if
   
   'response.write "test" & SOBestiltAv
   'asdas
 
   strSQL = "UPDATE oppdrag SET " &_
          "StatusID = " & opdragStatus & ", " & _
          "AnsmedID = " & Request.form("dbxMedarbeider") & ", " & _
          "AvdelingskontorID = " & Request.form("dbxAvdeling") & ", " & _
          "AvdelingID = " & Request.form("dbxregnskap") & ", " & _
          "TomID = " & Request.form("dbxtjenesteomrade") & ", " & _
          "CategoryID = " & Request.form("dbxCategory") & ", " & _
          "Beskrivelse = " & Quote( Request.form("tbxBeskrivelse") ) & ", " & _
          "FraDato = " & DbDate( Request.form("tbxFraDato") ) & ", " & _
          "FraKl = " &  DbTime( Request.form("tbxFraKl") ) & ", " & _
          "TilDato = " & DbDate( Request.form("tbxTilDato") ) & ", " & _
          "TilKl = " &  DbTime( Request.form("tbxTilKl") ) & ", " & _
          "FirmaID = " & Request.form("FirmaID") & ", " & _
          "ArbAdresse = " & Quote( Request.form("tbxArbAdresse") ) & ", " & _
          "Bestiltdato = " & DbDate( Request.form("tbxBestDato") ) & ", " & _
          "Bestiltklokken = " & DbTime( Request.form("tbxBestKl") ) & ", " & _
		  "BestilltAv = " & bestiltAv & ", " & _          
          "SOPeID = " & SOBestiltAv & ", " & _
          "ReportingContactID = " & ReportingContactID & ", " & _
          "ReportingOverrule = " & bQuestback & ", " & _
          "Kurskode = " & Kurskode & ", " & _
          "Oppdragskode = " & Oppdragskode & ", " & _
          "Timepris = " & TimePris & ", " & _
          "Timerprdag = " & Timerprdag & ", " & _
          "Timeloenn = " & Quote(timeloenn) & ", " & _
          "Lunch = " & DbTime(Lunsj) & ", " & _          
          "NotatAnsvarlig = " & Quote( PadQuotes(Request.form("tbxNotatAnsvarlig")) )& ", " &_ 
          "NotatOkonomi = " & Quote( PadQuotes(Replace((Replace(Request.form("tbxNotatOkonomi"),"<p>","")),"</p>","")) )& ", "  & _          
          "InvoiceComments = " & Quote( PadQuotes(Request.form("tbxInvoiceComments")) )& ", " &_ 
          "CustomerReference = " & Quote( PadQuotes(Request.form("tbxCustomerReference")) )& ", " &_           
          "NoOfPersons = " & noOfOpportunities & ", " &_           
          "Terms = " & ComTerms& ", " 
	if Request.Form("dbxFAgreement") = 0 then
		strSQL = strSQL + "FaID = Null WHERE oppdragid =" & LOppdragID
		SetEmptyCategoryVikar(LOppdragID)
	else
		strSQL = strSQL + "FaID =" & Request.Form("dbxFAgreement") & " WHERE oppdragid =" & LOppdragID
	End If		 

	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil ved oppdatering av oppdrag!")
		call RenderErrorMessage()
	End if


    'if fixed salary employee 
    if   (Kurskode = 0 and trim(ComTerms) = "1") then
    
      strSQL = "UPDATE  EC_Oppdrag_Terms SET " &_
          "Amount = " & amount  & "  ," & _
            "HourlyRate = " & hourlyRate  & "   " & _
          "WHERE  oppdragid =" & LOppdragID
          
 
  
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
    end if 
    
    'if direct recruitment
    if   (Kurskode = 3) then    
    
      strSQL =  "UPDATE  EC_Oppdrag_Terms SET " &_
          "Amount = " & amount  & ", " & _
          "noofpersons = " & noofpersons  & ", " & _
           "RecruitmentDate = " & DbDate(Request.Form("tbxRecruitmentDate")) & ", " & _
          "Comment = " & Quote(PadQuotes(Request.Form("tbxComment"))) & " " & _
          "WHERE   oppdragid =" & LOppdragID

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
		
		
	' if status change to fullstending add data to fakturagrunnlag table
	if  Request.form("dbxStatus") <> "" Then
		  
		commisionStatus= Request.form("dbxStatus")
		 
 
	 
		 strSQL = "SELECT Oppdragsstatus FROM  H_OPPDRAG_STATUS WHERE  OppdragsstatusID = " & commisionStatus
           set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
     
            
            IF rsOppdragStatus("Oppdragsstatus") = "Fullstendig" then
            
              strSQL = "SELECT InvMaxNo = MAX(Fakturanr) FROM FAKTURAGRUNNLAG" 
              set rsInvNumber  =  GetFirehoseRS(strSQL, Conn)          

               
            ' Create sql-statement insert to Add : 1st line for Direct recruitment (Recruitment text with values)                  
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_ 
				", " &  LOppdragID &_
				", " & -1	&_
				", " &   1   &_
				", " &   amount  &_
				", " &  1 &_
				", " & "301099" &_
				", " & "Rekruttering" & _
				", " & amount   &_
				", " & "null" &_
				", " & Request.Form("dbxSOPeID") &_
				", " & 1 &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				     CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
			
			' Create sql-statement insert to Add : 2nd - Multiple line for Direct recruitment (Comment text, only if it exists)    						
			
			linenumber = 1     ''' default starting number
			commentstring = Trim(Request.Form("tbxComment"))			
			length = Len(commentstring)
						
			Do While length>0
			  If length>60 then
			    'document.write("<br/>Length : " &length)    
			    If Mid(commentstring,60,1) = " " then  			
			      strpart = Trim(Left(commentstring,60))
			      commentstring = Trim(Right(commentstring,length-60))
			    Else 					 
			      strpart = Trim(Left(commentstring,60))
			      position = InStrRev(strpart," ")
			      'document.write("<br/>Position : "&position)
			      strpart = Trim(Left(commentstring,position))
			      commentstring = Trim(Right(commentstring,length-position))
			    End if 
			    'document.write("<br/>strpart :"&strpart)
			    'document.write("<br/>commentstring :"&commentstring)
			    length = Len(commentstring)
			  Else
			    strpart = commentstring    
			    length = 0
			
			  End if 
			
			  ''' Insert comment line in DB
			  linenumber = linenumber + 1
			  strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   LOppdragID &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes(strpart)) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.Form("dbxSOPeID") &_
				", " & linenumber  &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")		
       
	       		' Check whether the text has any values befor inserting	            
				If ExecuteCRUDSQL(strSQL, conn) = false then
					    CloseConnection(Conn)
				        set Conn = nothing
				        AddErrorMessage("Feil på beløpet feltet!")
				        call RenderErrorMessage()
				End if
			Loop	
			''' LOOP ENDS HERE			
			
			' Create sql-statement insert to Add : Final line for Direct recruitment (Oppdrag No)
			linenumber = linenumber + 1
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   LOppdragID &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes( "Oppdrag "  &  LOppdragID )) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.Form("dbxSOPeID") &_
				", " & linenumber &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")       
      
			If ExecuteCRUDSQL(strSQL, conn) = false then
				     CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if			
           
           END IF
      END IF     
		
    end if 
    
    
   ' Set return value
   strOppdragID = LOppdragID
   
  
'==========================================================
ElseIf Trim( Request.Form("hdnJobAction") ) = "copy" Then

	if cint(Request.Form("dbxAvdeling")) = 0 then
		AddErrorMessage("Avdeling er ikke utfylt!")
	end if

	if cint(Request.Form("dbxregnskap")) = 0 then
		AddErrorMessage("Regnskapsavdeling er ikke utfylt!")
	end if

	if Request.Form("dbxKontaktP") = "0" then
		AddErrorMessage("Du må velge en kontaktperson hos Kontakt!")
	end if
	
	if Request.Form("dbxCategory") = "0" then
		AddErrorMessage("Velg en kategori!")		
	end if
	
	if Request.Form("dbxFAgreement") = "-1" then
		AddErrorMessage("Du må velge om det skal føres eget kunderegnskap på dette oppdrag!")
	end if
			
		
	if(HasError() = true) then
		call RenderErrorMessage()
	end if
	
	strSQl = "SELECT BekreftelseKunde, Notat, ArbAdrID, ProgramID, KompNiva, DokID, Deltagere, Notatkunde, NotatVikar, GmlOppdragID, VilkaarSendtDato, VilkaarMotattDato  FROM Oppdrag WHERE oppdragid =" & LOppdragID
	set rsOppdrag = GetFirehoseRS(strSQL, Conn)
	
	strSQL = "INSERT INTO Oppdrag(" &_
				"IsCustomerApproval,StatusID, TypeID, AnsMedID, AvdelingskontorID," &_
				"AvdelingID, tomID,CategoryID, beskrivelse, " &_
				"fradato, frakl, tildato, tilkl, " &_  
				"firmaid, ArbAdresse, bestiltDato, bestilltAv," &_
				"SOPeID, oppdragskode, " &_
				"kurskode, " &_
				"bestiltklokken, timepris, timerprdag, timeloenn, lunch, " &_
				"notatansvarlig, notatokonomi, InvoiceComments, CustomerReference, " &_
				"BekreftelseKunde, Notat, ArbAdrID, " &_
				"ProgramID, KompNiva, DokID, Deltagere, " &_
				"Notatkunde, NotatVikar, GmlOppdragID, VilkaarSendtDato, VilkaarMotattDato, NoOfPersons, ParentAssignment,Terms,webPub,IsCopied," &_
				"FaID) " &_
            "Values(" &_
            			customerApprovalDefaultValue & ", " &_
				Request.form("dbxStatus") & ", " &_
				"0, " &_
				Request.form("dbxMedarbeider") & ", " &_
				Request.form("dbxAvdeling") & ", " &_ 
				Request.Form("dbxregnskap") & ", " &_
				Request.form("dbxtjenesteomrade") & ", " &_
				Request.form("dbxCategory") & ", " &_
				Quote( Request.form("tbxBeskrivelse") ) & ", " &_
				DbDate( Request.form("tbxFraDato") ) & ", " &_
				DbTime( Request.form("tbxFraKl") ) & ", " &_
				DbDate( Request.form("tbxTilDato") ) & ", " &_
				DbTime( Request.form("tbxTilKl") ) & ", " &_
				Request.form("FirmaID") & ", " &_
				Quote( Request.form("tbxArbAdresse") ) & ", " & _
				DbDate( Request.form("tbxBestDato") ) & ", " & _
				"NULL , " & _
				Request.form("dbxSOKontaktP") & ", " & _		
				Oppdragskode & ", " & _
				Kurskode & ", " & _
				DbTime( Request.form("tbxBestKl") ) & ", " &_
				Timepris & ", " &_
				Timerprdag & ", " &_
				Quote(timeloenn) & ", " &_
				DbTime( Lunsj ) & ", " &_
				Quote( PadQuotes(Request.form("tbxNotatAnsvarlig")) ) & ", " &_
				Quote( PadQuotes(Replace((Replace(Request.form("tbxNotatOkonomi"),"<p>","")),"</p>","")) ) & ", " &_				
				Quote( PadQuotes(Request.form("tbxInvoiceComments")) ) & ", " &_
				Quote( PadQuotes(Request.form("tbxCustomerReference")) ) & ", " &_
				FixString(rsOppdrag("BekreftelseKunde")) & "," &_
				FixString(rsOppdrag("Notat")) & ","  &_
				FixString(rsOppdrag("ArbAdrID")) & "," &_
				FixString(rsOppdrag("ProgramID")) & "," &_
				FixString(rsOppdrag("KompNiva")) & "," &_
				FixString(rsOppdrag("DokID")) & "," &_
				FixString(rsOppdrag("Deltagere")) & "," &_
				FixString(rsOppdrag("Notatkunde")) & "," &_
				FixString(rsOppdrag("NotatVikar")) & "," &_
				FixString(rsOppdrag("GmlOppdragID")) & "," &_
				FixString(rsOppdrag("VilkaarSendtDato")) & "," &_
				FixString(rsOppdrag("VilkaarMotattDato")) & "," &_
				noOfOpportunities & ", " & _
				parentAssignment & ", " & _
				Quote(ComTerms) & ", " & _
				"0 ,1,"
			
	if Request.Form("dbxFAgreement") = 0 then
		strSQL = strSQL +   "NULL)"
	else
		strSQL = strSQL +  Request.Form("dbxFAgreement") + ")"
	End If			

	' Insert into database
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil ved oppretting av oppdrag!")
		call RenderErrorMessage()
	End if

    ' Get new oppdragID
    strSQL = "SELECT NewOppdragID = MAX( OppdragID ) FROM Oppdrag"
    set rsOppdrag = GetFirehoseRS(strSQL, Conn)


   'if fixed salary employee 
    if   (Kurskode = 0 and trim(ComTerms) = "1") then
    
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,HourlyRate) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",0," & amount & "," & hourlyRate & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
    end if 
    
    'if direct employeeement
    if   (Kurskode = 3) then
        
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,noofpersons,RecruitmentDate,Comment) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",1," & amount & "," & noofpersons & "," & DbDate(Request.Form("tbxRecruitmentDate")) & "," & Quote( PadQuotes(Request.Form("tbxComment"))) & ")"
 
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
		
			commisionStatus= Request.form("dbxStatus")
	 
		 strSQL = "SELECT Oppdragsstatus FROM  H_OPPDRAG_STATUS WHERE  OppdragsstatusID = " & commisionStatus
           set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
     
            
            IF rsOppdragStatus("Oppdragsstatus") = "Fullstendig" then            
          
               strSQL = "SELECT InvMaxNo = MAX(Fakturanr) FROM FAKTURAGRUNNLAG" 
              set rsInvNumber  =  GetFirehoseRS(strSQL, Conn)             
             
			' Create sql-statement insert to Add : 1st line for Direct recruitment (Recruitment text with values)      
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_ 
				", " &  rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   1   &_
				", " &   amount  &_
				", " &  1 &_
				", " & "301099" &_
				", " & "Rekruttering" & _
				", " & amount   &_
				", " & "NULL" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & 1 &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")

			If ExecuteCRUDSQL(strSQL, Conn) = false then
				    CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
			
			' Create sql-statement insert to Add : 2nd - Multiple line for Direct recruitment (Comment text, only if it exists)    			
			
			linenumber = 1     ''' default starting number
			commentstring = Trim(Request.Form("tbxComment"))			
			length = Len(commentstring)
						
			Do While length>0
			  If length>60 then
			    'document.write("<br/>Length : " &length)    
			    If Mid(commentstring,60,1) = " " then  			
			      strpart = Trim(Left(commentstring,60))
			      commentstring = Trim(Right(commentstring,length-60))
			    Else 					 
			      strpart = Trim(Left(commentstring,60))
			      position = InStrRev(strpart," ")
			      'document.write("<br/>Position : "&position)
			      strpart = Trim(Left(commentstring,position))
			      commentstring = Trim(Right(commentstring,length-position))
			    End if 
			    'document.write("<br/>strpart :"&strpart)
			    'document.write("<br/>commentstring :"&commentstring)
			    length = Len(commentstring)
			  Else
			    strpart = commentstring    
			    length = 0
			
			  End if 
			
			  ''' Insert comment line in DB
			  linenumber = linenumber + 1
			  strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes(strpart)) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & linenumber &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")		
       
	       		' Check whether the text has any values befor inserting	            
				If ExecuteCRUDSQL(strSQL, conn) = false then
					    CloseConnection(Conn)
				        set Conn = nothing
				        AddErrorMessage("Feil på beløpet feltet!")
				        call RenderErrorMessage()
				End if
			Loop	
			''' LOOP ENDS HERE					
			
			' Create sql-statement insert to Add : Final line for Direct recruitment (Oppdrag No)
			linenumber = linenumber + 1
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1 &_
				", " & "NULL" &_
				", " & Quote(PadQuotes( "Oppdrag "  & rsOppdrag("NewOppdragID"))) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.form("dbxSOKontaktP") &_
				", " & 3 &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")       
      
			If ExecuteCRUDSQL(strSQL, conn) = false then
				     CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
			
            
           END IF
		
    end if 



    ' Do we have a KURS-PROGRAM ?
    If Request.form("dbxProgram") <> 0 Then

       ' Insert program AND kompnivå in OPPDRAG_KOMPETANSE
      strSQL = "INSERT INTO OPPDRAG_KOMPETANSE" &_
			"(oppdragID, K_TypeID, K_TittelID, K_LevelID ) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ", 3, " &_
			Request.form("dbxProgram") & "," & Request.form("dbxKompniva") & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil ved oppretting av kursprogram!")
			call RenderErrorMessage()
		End if

    End If

    ' set return value
    strOppdragID = rsOppdrag("NewOppdragID")
'=============================================================
ElseIf Trim( Request.Form("hdnJobAction") ) = "copyassign" Then
	dim vikarid 
	vikarid = Request.Form("hdnVikarID")
	
	if cint(Request.Form("dbxAvdeling")) = 0 then
		AddErrorMessage("Avdeling er ikke utfylt!")
	end if

	if cint(Request.Form("dbxregnskap")) = 0 then
		AddErrorMessage("Regnskapsavdeling er ikke utfylt!")
	end if
	
	if Request.Form("dbxCategory") = "0" then
		AddErrorMessage("Velg en kategori!")		
	end if

	if Request.Form("dbxKontaktP") = "0" then
		AddErrorMessage("Du må velge en kontaktperson hos Kontakt!")
	end if
	
	if Request.Form("dbxFAgreement") = "-1" then
		AddErrorMessage("Du må velge om det skal føres eget kunderegnskap på dette oppdrag!")
	end if
		
	if(HasError() = true) then
		'call RenderErrorMessage()
	end if
	
	strSQl = "SELECT BekreftelseKunde, Notat, ArbAdrID, ProgramID, KompNiva, DokID, Deltagere, Notatkunde, NotatVikar, GmlOppdragID, VilkaarSendtDato, VilkaarMotattDato  FROM Oppdrag WHERE oppdragid =" & LOppdragID
	set rsOppdrag = GetFirehoseRS(strSQL, Conn)
	
	strSQL = "INSERT INTO Oppdrag(" &_
				"IsCustomerApproval,StatusID, TypeID, AnsMedID, AvdelingskontorID," &_
				"AvdelingID, tomID,CategoryID, beskrivelse, " &_
				"fradato, frakl, tildato, tilkl, " &_  
				"firmaid, ArbAdresse, bestiltDato, bestilltAv," &_
				"SOPeID, oppdragskode, " &_
				"kurskode, " &_
				"bestiltklokken, timepris, timerprdag, timeloenn, lunch, " &_
				"notatansvarlig, notatokonomi, InvoiceComments, CustomerReference, " &_
				"BekreftelseKunde, Notat, ArbAdrID, " &_
				"ProgramID, KompNiva, DokID, Deltagere, " &_
				"Notatkunde, NotatVikar, GmlOppdragID, VilkaarSendtDato, VilkaarMotattDato,NoOfPersons,ParentAssignment,Terms, webPub,IsCopied, " &_
				"FaID) " &_
            "Values(" &_
            			customerApprovalDefaultValue & ", " &_
				Request.form("dbxStatus") & ", " &_
				"0, " &_
				Request.form("dbxMedarbeider") & ", " &_
				Request.form("dbxAvdeling") & ", " &_ 
				Request.Form("dbxregnskap") & ", " &_
				Request.form("dbxtjenesteomrade") & ", " &_
				Request.form("dbxCategory") & ", " &_
				Quote( Request.form("tbxBeskrivelse") ) & ", " &_
				DbDate( Request.form("tbxFraDato") ) & ", " &_
				DbTime( Request.form("tbxFraKl") ) & ", " &_
				DbDate( Request.form("tbxTilDato") ) & ", " &_
				DbTime( Request.form("tbxTilKl") ) & ", " &_
				Request.form("FirmaID") & ", " &_
				Quote( Request.form("tbxArbAdresse") ) & ", " & _
				DbDate( Request.form("tbxBestDato") ) & ", " & _
				"NULL , " & _
				Request.form("dbxSOKontaktP") & ", " & _		
				Oppdragskode & ", " & _
				Kurskode & ", " & _
				DbTime( Request.form("tbxBestKl") ) & ", " &_
				Timepris & ", " &_
				Timerprdag & ", " &_
				Quote(timeloenn) & ", " &_
				DbTime( Lunsj ) & ", " &_
				Quote( PadQuotes(Request.form("tbxNotatAnsvarlig")) ) & ", " &_
				Quote( PadQuotes(Replace((Replace(Request.form("tbxNotatOkonomi"),"<p>","")),"</p>","")) ) & ", " &_				
				Quote( PadQuotes(Request.form("tbxInvoiceComments")) ) & ", " &_
				Quote( PadQuotes(Request.form("tbxCustomerReference")) ) & ", " &_
				FixString(rsOppdrag("BekreftelseKunde")) & "," &_
				FixString(rsOppdrag("Notat")) & ","  &_
				FixString(rsOppdrag("ArbAdrID")) & "," &_
				FixString(rsOppdrag("ProgramID")) & "," &_
				FixString(rsOppdrag("KompNiva")) & "," &_
				FixString(rsOppdrag("DokID")) & "," &_
				FixString(rsOppdrag("Deltagere")) & "," &_
				FixString(rsOppdrag("Notatkunde")) & "," &_
				FixString(rsOppdrag("NotatVikar")) & "," &_
				FixString(rsOppdrag("GmlOppdragID")) & "," &_
				FixString(rsOppdrag("VilkaarSendtDato")) & "," &_
				FixString(rsOppdrag("VilkaarMotattDato")) & "," &_
				noOfOpportunities & ", " & _			
				parentAssignment & ", " & _	
				Quote(ComTerms) & ", " & _
				"0 ,1,"
			
	if Request.Form("dbxFAgreement") = 0 then
		strSQL = strSQL +   "NULL)"
	else
		strSQL = strSQL +  Request.Form("dbxFAgreement") + ")"
	End If			

	' Insert into database
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil ved oppretting av oppdrag!")
		call RenderErrorMessage()
	End if

    ' Get new oppdragID
    strSQL = "SELECT NewOppdragID = MAX( OppdragID ) FROM Oppdrag"
    set rsOppdrag = GetFirehoseRS(strSQL, Conn)

	dim newopdragid 
	newopdragid = rsOppdrag("NewOppdragID")


   'if fixed salary employee 
    if   (Kurskode = 0 and trim(ComTerms) = "1") then
    
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,HourlyRate) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",0," & amount & "," &  hourlyRate & ")"
        
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
    end if 
    
    'if direct recruitment
    if   (Kurskode = 3) then   
    
      strSQL = "INSERT INTO EC_Oppdrag_Terms" &_
			"(oppdragID,TermType, Amount,noofpersons,RecruitmentDate,Comment) " & _
			"VALUES( " & rsOppdrag("NewOppdragID") & ",1," & amount & "," & noofpersons & "," & DbDate(Request.Form("tbxRecruitmentDate")) & "," & Quote( PadQuotes(Request.Form("tbxComment"))) & ")"
 
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil på beløpet feltet!")
			call RenderErrorMessage()
		End if
		
		
		
 
		commisionStatus= Request.form("dbxStatus")
	 
		 strSQL = "SELECT Oppdragsstatus FROM  H_OPPDRAG_STATUS WHERE  OppdragsstatusID = " & commisionStatus
           set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
     
            
            IF rsOppdragStatus("Oppdragsstatus") = "Fullstendig" then
            
                strSQL = "SELECT InvMaxNo = MAX(Fakturanr) FROM FAKTURAGRUNNLAG" 
              set rsInvNumber  =  GetFirehoseRS(strSQL, Conn)
           
            ' Create sql-statement for Procedure Lag_timeliste               
            
            ' Create sql-statement insert to Add : 2nd - Multiple line for Direct recruitment (Comment text, only if it exists)    			

			linenumber = 1     ''' default starting number
			commentstring = Trim(Request.Form("tbxComment"))			
			length = Len(commentstring)
						
			Do While length>0
			  If length>60 then
			    'document.write("<br/>Length : " &length)    
			    If Mid(commentstring,60,1) = " " then  			
			      strpart = Trim(Left(commentstring,60))
			      commentstring = Trim(Right(commentstring,length-60))
			    Else 					 
			      strpart = Trim(Left(commentstring,60))			      
			      position = InStrRev(strpart," ")
			      'document.write("<br/>Position : "&position)
			      strpart = Trim(Left(commentstring,position))
			      commentstring = Trim(Right(commentstring,length-position))
			    End if 
			    'document.write("<br/>strpart :"&strpart)
			    'document.write("<br/>commentstring :"&commentstring)
			    length = Len(commentstring)
			  Else
			    strpart = commentstring    
			    length = 0
			
			  End if 
			
			  ''' Insert comment line in DB
			  linenumber = linenumber + 1
			  strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_ 
				", " &  rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   1   &_
				", " &   amount  &_
				", " &  1 &_
				", " & "301099" &_
				", " & Quote(PadQuotes(strpart)) & _
				", " & amount   &_
				", " & "null" &_
				", " & Request.Form("dbxSOPeID") &_
				", " & linenumber &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")
       
	       		' Check whether the text has any values befor inserting	            
				If ExecuteCRUDSQL(strSQL, conn) = false then
					    CloseConnection(Conn)
				        set Conn = nothing
				        AddErrorMessage("Feil på beløpet feltet!")
				        call RenderErrorMessage()
				End if
			Loop	
			''' LOOP ENDS HERE             
			
			' Create sql-statement insert to Add : Final line for Direct recruitment (Comment text, only if it exists)  
			linenumber = linenumber + 1
			strSQL = "EXECUTE [EC_AddInvoiceLine] " & Request.form("FirmaID") &_  
				", " &   rsOppdrag("NewOppdragID") &_
				", " & -1	&_
				", " &   "NULL"  &_
				", " &   "NULL"  &_
				", " &  1  &_
				", " & "null" &_
				", " & Quote(PadQuotes( "Oppdrag "  & rsOppdrag("NewOppdragID"))) & _
				", " & "NULL"  &_
				", " & "null" &_
				", " & Request.Form("dbxSOPeID") &_
				", " & 2 &_
				", " & "NULL"  &_
				", " & rsInvNumber("InvMaxNo") + 1  &_
				", " & "NULL" &_
				", " & "NULL" &_
				", " & Request.form("dbxregnskap")
       
      
			If ExecuteCRUDSQL(strSQL, conn) = false then
				     CloseConnection(Conn)
			        set Conn = nothing
			        AddErrorMessage("Feil på beløpet feltet!")
			        call RenderErrorMessage()
			End if
			
            
           END IF
           
           
    end if 

    ' Do we have a KURS-PROGRAM ?
    If Request.form("dbxProgram") <> 0 Then

       ' Insert program AND kompnivå in OPPDRAG_KOMPETANSE
      strSQL = "INSERT INTO OPPDRAG_KOMPETANSE" &_
			"(oppdragID, K_TypeID, K_TittelID, K_LevelID ) " & _
			"VALUES( " & newopdragid & ", 3, " &_
			Request.form("dbxProgram") & "," & Request.form("dbxKompniva") & ")"

		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil ved oppretting av kursprogram!")
			call RenderErrorMessage()
		End if

    End If
    
    'assign temp
    strSQL = "SELECT OppdragVikarStatusID FROM H_OPPDRAG_VIKAR_STATUS WHERE Status = 'Aksept'"
	set rsOppdrag = GetFirehoseRS(strSQL, Conn)
	
	strSQL = "UPDATE OPPDRAG_VIKAR SET StatusID =" & rsOppdrag("OppdragVikarStatusID") & ", OppdragID=" & newopdragid & " WHERE OppdragID=" & LOppdragID & " AND VikarID=" & vikarid
	
	If ExecuteCRUDSQL(strSQL, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil ved oppretting av kursprogram!")
		call RenderErrorMessage()
	End if
	
	'if FaID is there 
	'Response.redirect "/xtra/oppdragVikarny.asp?OppdragVikarID=" & newOppdragVikarId & "&FaId=" & lFaId
	if Request.Form("dbxFAgreement") > 0 then
		strSQL = "SELECT oppdragvikarid FROM Oppdrag_Vikar WHERE OppdragId = " & newopdragid & " AND VikarId = " & vikarid	
		set rsOppdrag = GetFirehoseRS(strSQL, Conn)
		
		Dim ovid
		ovid = rsOppdrag("oppdragvikarid")
		
		CloseConnection(Conn)
		set Conn = nothing
									
		Response.redirect "/xtra/oppdragVikarny.asp?OppdragVikarID=" & ovid & "&FaId=" & Request.Form("dbxFAgreement")
	end if
	
	'register automatic activity for temp assign to commision
	Dim strActivity
	Dim strComment
	Dim sDate
	strActivity = "Vikar aksept"
	strComment = "Vikar akseptert. Dato: " & Request.form("tbxFraDato") & " - " & Request.form("tbxTilDato") & " Lønn: " & timeloenn
	set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", Conn)
	nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
	' Close and release recordset
      	rsActivityType.Close
      	Set rsActivityType = Nothing
	      	
      	sDate  = GetDateNowString()	      		      	
	strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
		"Values(" & nActivityTypeID & ",'" & sDate & "'," & vikarid & "," & Request.form("FirmaID") & "," & newopdragid & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
	
	If ExecuteCRUDSQL(strSql, Conn) = false then
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Aktivitetsregistrering for bytte av status feilet.")
		call RenderErrorMessage()
	End if
			
    ' set return value
    strOppdragID = newopdragid
    
'//=============================================================
ElseIf Trim( Request.Form("hdnJobAction") ) = "Slette" Then

	Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))
	ConnTrans.BeginTrans

   ' Delete OPPDRAG
   strSQL = "DELETE oppdrag WHERE oppdragid = " & LOppdragID

	If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
		ConnTrans.RollbackTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
		AddErrorMessage("Feil ved sletting av oppdrag!")
		call RenderErrorMessage()
	End if

   ' Create Sql-Statement
   strSQL = "DELETE oppdrag_Vikar WHERE oppdragid = " & LOppdragID

   ' Delete oppdrag-vikar in database
	If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
		ConnTrans.RollbackTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
		AddErrorMessage("Feil ved sletting av Oppdragvikar!")
		call RenderErrorMessage()
	End if

   ' Create Sql-Statement
   strSQL = "DELETE oppdrag_kompetanse WHERE oppdragid = " & LOppdragID

   ' Delete oppdrag-kompetanse in database
   ' Delete oppdrag-vikar in database
	If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
		ConnTrans.RollbackTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
		AddErrorMessage("Feil ved sletting av oppdrags kompetanse!")
		call RenderErrorMessage()
	End if
	

	ConnTrans.CommitTrans
	CloseConnection(ConnTrans)
	set ConnTrans = nothing
	Response.redirect "OppdragSoek.asp"
Else
	AddErrorMessage("Systemfeil: Parameter mangler!")
	call RenderErrorMessage()
End If

CloseConnection(Conn)
set Conn = nothing

' Redirect to updated page..
 If Request.form("dbxStatus") = "1" then
      
     Response.redirect "WebUI/OppdragView.aspx?OppdragID=" & strOppdragID
else
   'Redirect to Automatially unpublish page 
     Response.redirect "WebUI\UnpublishAddAP.aspx?commid=" & strOppdragID
  
End If
%>