<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Response.Expires = 0

%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.Economics.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	'Variables containing search criterias
	dim strSARating1
	dim strSARating2
	dim strSARating3
	dim lngVikarID
	dim strBil
	dim lngAnsattID
	dim strSoekForNavn
	dim strSoekEtterNavn
	dim strPostnr
	dim rsCategoryAreas
	dim strCategoryAreas
	dim lngManagerID
	dim strBranches
	dim strServiceAreas
	dim strOppsigelsestid
	dim aCategory
	dim strWorkType
	dim strCategorySearchType
	dim strCheckedWorkType0
	dim strCheckedWorkType1
	dim strCheckedWorkType2
	dim strCheckedWorkType3
	dim strCheckedOppsigelse0
	dim strCheckedOppsigelse1
	dim strCheckedOppsigelse2
	dim strCheckedOppsigelse3
	dim strCheckedOppsigelse4
	dim strCheckedOppsigelse5
	dim strCheckedInterviewed
	dim strFoererkort
	dim strCheckedDriverslicinseNo
	dim strCheckedDriverslicinseYes
	dim strFreeTextSearch
	
	'Variables to sort search -PDF
	dim strSortByCategory
	dim strSortDesc
	dim strCheckedSortCategory0
	dim strCheckedSortCategory1
	dim strCheckedSortCategory2
	dim strCheckedSortCategory3
	dim strCheckedSortCategory4
	dim strCheckedSortCategory5
	
	'Variables used to maintain page state
	dim chk
	dim i
	dim strFreeTextSearchType
	dim AWords
	dim n
	dim strDisplayOtherInfo : strDisplayOtherInfo = "none"
	dim strDisplayQualifications
	dim strDisplayFreeText : strDisplayFreeText = "none"
	dim strDisplayCourseSpecific : strDisplayCourseSpecific = "none"	
	dim strDisplayStatus
	dim chkManger
	
	'Variables used to connect to database and retrieve data
	dim objCon
	dim rsStatus
	dim rsBranches
	dim rsServiceAreas
	dim rsEmailTemplates
	dim rsAcceptedCommissions
	'Variables used to maintain temporary state
	dim strConsultantStatusID
	dim strCboSelected
	dim strCboProductSelected
	dim strExcludeServiceAreas
	dim strRatedAreas
	dim strRatedArea
	dim rsProductQualifications
	dim strExcludeProductQualifications
	dim strSAValue
	dim strSAName
	dim lngSARating
	dim strSALevel1
	dim strSALevel2
	dim strSALevel3
	dim intLevelID
	dim strLevel
	dim strRatedProductQualification
	dim rsSA
	dim strSearchSelectSQL
	dim strSearchFromSQL
	dim strSearchWhereSQL
	dim strSearchWhereRelationSQL
	dim strSearchSQL
	dim strSearchSortbySQL
	dim aAvdelingskontor
	dim aConsultantStatus
	dim aTjenesteomrader
	dim NOFArrayItems
	dim	strBranchSearchType
	dim	strSASearchType
	dim rsLevel
	dim rsProductRank
	dim strFraDato
	dim strTilDato
	dim StrAvdeling
	dim rsSoek
	dim strFullName
	dim strFullAdress
	dim strLoad
	dim blnFromJob	: blnFromJob = false
	dim lngoppdragID
	dim lngFirmaID

	dim intStatusID
	dim strUpdateSQL
	dim strSQL
	dim rsOppdrag
	dim strFraKl
	dim strTilkl
	dim TimerPrDag
	dim Timepris
	dim Lunsj
	dim Kurskode
	dim OppdragStatusID
	dim rsVikarLoenn
	dim rsVikarLocation
	dim lVikarID
	dim VikarLoenn
	dim VikarType
	dim rsVikar
	dim dblFaktor
	dim heading : heading = "Søk etter vikar"	'heading of page
	dim multy
	dim rsOppdragStatus
	dim rsOppdragVikarId
	

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList
	dim newOppdragVikarId
	dim lFaId
	
	dim oppdragVikarStatus
	dim rsRemovedState
	
	'Paging Constants
	const RECORDS_PER_PAGE = 100
	'Paging variables
	dim nPage
	dim nPageCount
	dim nRecCount
	dim postedPage:postedPage=""
	dim postedOppId
	dim postedFirmaId

	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = ""
	newOppdragVikarId = 0

	if Request.QueryString("Posted") <> "" then
		postedPage = Request.QueryString("Posted")
	elseif len(trim(Request("hdnPosted")))> 0 then
		postedPage = trim(request("hdnPosted"))
	end if
	
	if len(trim(Request("tbxOppdragID"))) > 0 then
		postedOppId = trim(Request("tbxOppdragID"))
	elseif Request.QueryString("OppdragId") <> "" then
		postedOppId = Request.QueryString("OppdragId")
	end if
	
	if len(trim(Request("tbxFirmaID"))) > 0 then
		postedFirmaId = trim(Request("tbxFirmaID"))
	elseif Request.QueryString("FirmaId") <> "" then
		postedFirmaId = Request.QueryString("FirmaId")
	end if
	

	'Initialize
	strExcludeServiceAreas = "''"
	strExcludeProductQualifications = "''"

	' Get the current page from the hidden control
	If len(trim(Request("hdnCurrentPage")))> 0 Then
	   nPage = clng(Request("hdnCurrentPage"))
	Else
	   nPage = 1
	End If

	' Open connection
	Set objCon = GetConnection(GetConnectionstring(XIS, ""))	
	
	
	' Get Qualification levels..
	set rsLevel = GetDynamicRS("select [K_LevelID],[Klevel] from [H_KOMP_LEVEL]", objCon)
	Set rsProductRank = GetDynamicRS("select [K_rangeringID],[K_rangering] from [H_KOMP_RANGERING] order by [K_rangeringID]", objCon)
	
	' Get removed state ID
    set rsRemovedState = GetFirehoseRS("select VikarstatusID FROM H_VIKAR_STATUS WHERE Vikarstatus = 'Fjernet'", objCon)

	lngManagerID = -1 'Session("medarbID") PDF

	'This is postback
	if (postedPage ="1") then

		lngVikarID				= request("txtVikarID")
		lngAnsattID				= request("txtAnsattID")
		strSoekForNavn			= request("txtForNavn")
		strSoekEtterNavn		= request("txtEtternavn")
		strConsultantStatusID	= request("chkConsultantStatus")
		lngManagerID			= request("cboManager")
		strBranches				= request("cboBranch")
		strServiceAreas			= request("cboServiceArea")
		strCategoryAreas			= request("cboCategory")
		strOppsigelsestid		= request("cboOppsigelse")
		strWorkType			= request("cboWorkType")
		strFoererkort			= request("cbodriverLicence")
		strPostnr				= request("txtPostnr")
		strBranchSearchType		= request("cboBranchSearchType")
		strSASearchType			= request("cboSASearchType")
		strCategorySearchType			= request("cboCategorySearchType")
		strFreeTextSearchType	= request("dbxFreeTextSearchType")
		strFreeTextSearch		= trim(request("txtFreeTextSearch"))
		strFraDato				= trim(Request("dtFromDate"))
		strTilDato				= trim(Request("dtToDate"))
		lngoppdragID			= postedOppId
		strCheckedInterviewed	= request("chkInterviewed")

		strSortByCategory		= request("cboSortCategory") 
		strSortDesc			= request("chkSortDesc") 
		
		if len(lngoppdragID) > 0 then
			blnFromJob	= true
			lngFirmaID	= postedFirmaId
			heading = "Tilknytt vikar til oppdrag"
		end if

		strCheckedOppsigelse0 = ""
		strCheckedOppsigelse1 = ""
		strCheckedOppsigelse2 = ""
		strCheckedOppsigelse3 = ""
		strCheckedOppsigelse4 = ""
		strCheckedOppsigelse5 = ""

		select case strOppsigelsestid
		case "0"
			strCheckedOppsigelse0 = "selected"
		case "1"
			strCheckedOppsigelse1 = "selected"
		case "2"
			strCheckedOppsigelse2 = "selected"
		case "3"
			strCheckedOppsigelse3 = "selected"
		case "4"
			strCheckedOppsigelse4 = "selected"
		case "5"
			strCheckedOppsigelse5 = "selected"
		end select

		strCheckedWorkType0 = ""
		strCheckedWorkType1 = ""
		strCheckedWorkType2 = ""
		strCheckedWorkType3 = ""
		
		select case strWorkType
		case "0"
			strCheckedWorkType0 = "selected"
		case "1"
			strCheckedWorkType1 = "selected"
		case "2"
			strCheckedWorkType2 = "selected"
		case "3"
			strCheckedWorkType3 = "selected"		
		end select

		if (len(trim(strBil)) = 0) then
			strBil = ""
		elseif (strBil = 0) then
			strBil = "Nei"
		elseif (strBil = 1) then
			strBil = "Ja"
		end if

		if (strFoererkort="0") then
			strCheckedDriverslicinseYes = "selected"
		elseif (strFoererkort="1") then
			strCheckedDriverslicinseNo = "selected"
		end if

		if (len(strPostnr) > 0 or len(strFoererkort) > 0 or len(strOppsigelsestid) > 0 or len(strWorkType) > 0 ) then
			strDisplayOtherInfo = ""
		end if

		if (len(trim(request("hdnServiceArea"))) > 0 or len(trim(request("hdnProductArea"))) > 0) then
			strDisplayQualifications = ""
		end if

		if (len(strFreeTextSearch) > 0 ) then
			strDisplayFreeText = ""
		end if
		
		if (len(strFraDato) > 0 OR len(strTilDato) > 0)  then
			strDisplayCourseSpecific = ""
		end if		
		
		'''STH bug fix start
		if lngoppdragID <> "" then
			strSQL = "select  count(*) as count , max(case isnull(statusid,0) " & _ 
		 		"when 4 then 1 " & _
		 		"else 0 end) as status " & _
		 		"from OPPDRAG_VIKAR where oppdragid=" & lngoppdragID
		else
			strSQL = "select  count(*) as count , max(case isnull(statusid,0) " & _ 
		 		"when 4 then 1 " & _
		 		"else 0 end) as status " & _
		 		"from OPPDRAG_VIKAR where oppdragid=0"
		End if
		 
		set rsOppdragStatus = GetFirehoseRS(strSQL, objCon)
	
		if (cint(rsOppdragStatus("count"))=0) then
			oppdragVikarStatus = 0	
		else
			oppdragVikarStatus = cint(rsOppdragStatus("status"))
		end if
	
		rsOppdragStatus.Close
		set rsOppdragStatus = Nothing
		''' End		
		
		if (request("hdnAddState")="0") then
		'Neither add or nor remove buttons have been clicked

			if (len(request("btnAxcept")) > 0 or len(request("btnOffer")) > 0) then
				'User has selected candidates for a job
				if len(request("btnAxcept")) > 0 Then
					intStatusID = 4 'aksept
					
				elseif len(request("btnOffer")) > 0 then
					intStatusID = 9 'tilbud (Short listed)
				End If

				' Hente oppdragsopplysninger
				strSQL = "SELECT [FirmaID], [fradato], [tildato], [Frakl], [Tilkl], [Timerprdag], [Timepris], [Timeloenn], [Lunch], [StatusID], [Kurskode], [FaID] " &_
						"FROM [oppdrag] where [oppdragid] = " & lngoppdragID

				set rsOppdrag = GetFirehoseRS(strSQL, objCon)

				strFraDato		= rsOppdrag("Fradato")
				strTilDato		= rsOppdrag("Tildato")
				lngFirmaID		= rsOppdrag("FirmaID")
				strFraKl		= FormatDateTime( rsOppdrag("FraKl"), 4)
				strTilkl		= FormatDateTime( rsOppdrag("TilKl"), 4)
				TimerPrDag		= rsOppdrag("TimerprDag")
				Timepris		= rsOppdrag("Timepris")
				Lunsj			= FormatDateTime( rsOppdrag("Lunch"), 3 )
				Kurskode		= rsOppdrag("Kurskode")
				lFaId			= rsOppdrag("FaID")
				OppdragStatusID = rsOppdrag("StatusID")
				
				if isnull(rsOppdrag("Timeloenn")) then
					VikarLoenn = 0
				else
					VikarLoenn = rsOppdrag("Timeloenn")
				end if
				
				Call fjernKomma(TimerPrDag)

				rsOppdrag.Close
				Set rsOppdrag = Nothing
				 
				' For hver vikar som er krysset av..
				For each chk in Request("tbxID")
					lVikarID = clng(chk)

					'Get salary pr hour 
					set rsVikarLoenn = GetFirehoseRS("SELECT [TypeID] FROM [VIKAR] WHERE [VikarID] = " & lVikarID, objCon)
					VikarType = rsVikarLoenn("TypeID")
         			rsVikarLoenn.Close
         			Set rsVikarLoenn = Nothing

					' er han allerde valgt men har en annen status?

					strSQL = "SELECT [VikarID] FROM [OPPDRAG_VIKAR] " &_
							"WHERE [OppdragID] = " & lngoppdragID &_
							" AND [VikarID] = " & lVikarID & " AND [StatusID] <> 4" 	'Put the status ids of Tilbud, Intervju, Avslag, Vurdert, Søkt stilling

					set rsVikar = GetFirehoseRS(strSQL, objCon)					

					If Not (rsVikar.EOF) Then
						AddErrorMessage("Vikaren er allerede koblet til oppdrag med en annen status enn aksept!")
						call RenderErrorMessage()
					End If
					rsVikar.Close
					Set rsVikar = Nothing

					if len(request("btnAxcept")) > 0 then
						set rsVikarLocation = GetFirehoseRS("exec [dbo].[spECGetVikarLocationAndIndustry] " & lVikarID, objCon)
						if(rsVikarLocation.EOF) then
							AddErrorMessage("Avdelingskontor, tjenesteområde og ansvarlig må være valgt på vikaren før aksept på oppdrag.")
							call RenderErrorMessage()
							
						End If
			
						if(len(trim(rsVikarLocation("Location"))) = 0 or len(trim(rsVikarLocation("Responsible"))) = 0 or len(trim(rsVikarLocation("Industry"))) = 0 or trim(rsVikarLocation("Responsible")) = "Ingen") then
							AddErrorMessage("Avdelingskontor, tjenesteområde og ansvarlig må være valgt på vikaren før aksept på oppdrag.")
							call RenderErrorMessage()
						end if	
						rsVikarLocation.Close
						set rsVikarLocation = nothing	

					end if 

					' Utregning av Faktor
					Response.Write Timepris
					Response.Write "@@@"
					Response.Write Vikarloenn
					Response.Write "----"
					
					
					' Convert from '.' to ',' in Timepris
					If Instr(Timepris, "." ) > 0 Then
						Timepris = Left( Timepris, Instr( Timepris, "." )-1) & "," & Mid(Timepris,Instr(Timepris, "." ) + 1)
					End If					
					
					' Convert from '.' to ',' in Vikarloenn
					If Instr(Vikarloenn, "." ) > 0 Then
						Vikarloenn = Left( Vikarloenn, Instr( Vikarloenn, "." )-1) & "," & Mid(Vikarloenn,Instr(Vikarloenn, "." ) + 1)
					End If	
					
					
					If Vikartype = 1 And Vikarloenn > 0  And Timepris > 0 Then   'Ansatt
						dblFaktor =  CDbl( Timepris / Vikarloenn )
					ElseIf Vikartype > 1 And Vikarloenn > 0 And Timepris > 0 Then 'Selvstendig og AS
						dblFaktor = CDbl( Timepris / ( VikarLoenn  / XIS_FACTOR )  )
					Else
						dblFaktor = 0
					End If
					
					Call fjernKomma(dblFaktor)
					Call fjernKomma(VikarLoenn)
					Call fjernKomma(Timepris)

					'lagring av vikar på oppdraget
					strSQL = "INSERT INTO [Oppdrag_Vikar](" &_
						"[OppdragID], " &_
						"[StatusID], " &_
						"[VikarID], " &_
						"[Timeloenn], " &_
						"[Timepris], " &_
						"[Timeliste], " &_
						"[fradato], " &_
						"[tildato], " &_
						"[frakl], " &_
						"[tilkl], " &_
						"[Lunch], " &_
						"[FirmaID], " &_
						"[Faktor], " &_
						"[Anttimer]," &_
						"[IsInAppliedList]" &_
					") Values (" &_
						lngoppdragID & "," & _
		      			intStatusID & "," & _
						lVikarID & "," &_
						VikarLoenn & "," & _
						Timepris & "," & _
 						"0," & _
						DbDate( strFraDato ) & "," & _
						DbDate( strTilDato ) & "," & _
						DbTime( strFraKl ) & "," & _
						DbTime( strTilKl ) & "," & _
						DbTime( Lunsj ) & "," & _
						lngFirmaID & "," & _
						dblFaktor & "," & _
						TimerPrDag & "," &_
						0 & ")"

					if (ExecuteCRUDSQL( strSql, objCon) = false) then
						AddErrorMessage("Feil oppstod under tilknytning av vikar til oppdrag!")
						call RenderErrorMessage()
					Else
					
						if (len(request("btnAxcept")) > 0) And  (isNull(lFaId)=  false ) Then
							strSQL = "SELECT  MAX(oppdragvikarid)AS oppdragvikarid FROM Oppdrag_Vikar WHERE OppdragId = " & lngoppdragID & " AND VikarId = " & lVikarID								
		 
							set rsOppdragVikarId = GetFirehoseRS(strSQL, objCon)	
						
							newOppdragVikarId = rsOppdragVikarId.Fields("oppdragvikarid").value
										
	
							rsOppdragVikarId.Close
							set rsOppdragVikarId = Nothing
							
							
						End If
					
						
					
					End if

					'Register activity for new Oppdrag
					Dim  strActivity
					Dim rsActivityType
					Dim strComment
					Dim nActivityTypeID
					Dim rsOppdrag2
					strActivity = "Tilknyttet oppdr."
					
					set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", objCon)
					nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
					' Close and release recordset
				      	rsActivityType.Close
				      	Set rsActivityType = Nothing
					
					set rsOppdrag2 = GetFirehoseRS("SELECT Status FROM H_OPPDRAG_VIKAR_STATUS WHERE OppdragVikarStatusID = " & intStatusID, objCon)
								
					strComment = "Tilknyttet oppdraget med status " & rsOppdrag2("Status")
				      	' Close and release recordset
				      	rsOppdrag2.Close
				      	Set rsOppdrag2 = Nothing
				      	
				      	Dim dt
				      	Dim strDate
				      	Dim sDate
				      	'Set dt = Now()
				      	sDate = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
				      					      	
					strSql = "INSERT INTO AKTIVITET( AktivitetTypeID, AktivitetDato, VikarID, FirmaID, OppdragID, Notat, Registrertav, RegistrertAvID, AutoRegistered )" &_
						"Values(" & nActivityTypeID & ",'" & sDate & "'," & lVikarID & "," & lngFirmaID & "," & lngoppdragID & ",'" & strComment & "','" & Session("Brukernavn") & "'," & Session("medarbID") & ", 1)"
					
					If ExecuteCRUDSQL(strSql, objCon) = false then
						response.write(strSql)
						AddErrorMessage("Aktivitetsregistrering for knytte vikar til oppdrag feilet.")
						call RenderErrorMessage()
					End if
							
				Next 'andre som er krysset av
				
								
				
				CloseConnection(objCon)
				set objCon = nothing
				
				'Redirect
				 if isNull(newOppdragVikarId) or (newOppdragVikarId = 0) then
				Response.redirect "/xtra/WebUI/OppdragView.aspx?OppdragID=" & lngoppdragID
				 Else
				 	Response.redirect "/xtra/oppdragVikarny.asp?OppdragVikarID=" & newOppdragVikarId & "&FaId=" & lFaId
				 End if
				
			end if


			'Do search

			'Init SQL string
			strSearchFromSQL = ""
			strSearchWhereSQL = ""
			strSearchSQL = ""
			strSearchSelectSQL = ""
			strSearchSortbySQL = ""
			
			If (len(trim(lngVikarID)) > 0) Then
				strSearchWhereSQL = " AND [V].[VikarID] = '" & lngVikarID & "'"
			end if

			If (len(trim(lngAnsattID)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL &  " AND [va].[ansattnummer] = '" & lngAnsattID & "'"
				strSearchWhereRelationSQL = strSearchWhereRelationSQL &	" AND [V].[VikarID] = [VA].[VikarID] "
			else
				strSearchWhereRelationSQL = strSearchWhereRelationSQL &	" AND [V].[VikarID] *= [VA].[VikarID] "
			end if

			if len(trim(strConsultantStatusID)) > 0 then
				aConsultantStatus = split(strConsultantStatusID,",")
				if Not IsEmpty(aConsultantStatus) then
					NOFArrayItems = ubound(aConsultantStatus)
					strSearchWhereSQL = strSearchWhereSQL & " AND ([V].[StatusID] = '" & trim(aConsultantStatus(0)) & "'"
					For I = 1 To NOFArrayItems
						strSearchWhereSQL =  strSearchWhereSQL & " OR [V].[StatusID] = '" & trim(aConsultantStatus(I)) & "'"
					Next
					strSearchWhereSQL = strSearchWhereSQL & ")"
				end if				
			end if
			'######
			'If (strConsultantStatusID <> "0") Then
				'strSearchWhereSQL =  strSearchWhereSQL & " AND [V].[StatusID] = '" & lngConsultantStatusID & "'"
			'end if

			If (lngManagerID <> 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[AnsMedID] = '" & lngManagerID & "'"
			end if
			If (len(trim(strSoekEtterNavn)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[Etternavn] LIKE '" & strSoekEtterNavn & "%'"
			end if

			If (len(trim(strSoekForNavn)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[Fornavn] LIKE '" & strSoekForNavn & "%'"
			end if

			If (len(trim(strPostnr)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [A].[postnr] LIKE '" & strPostnr & "'"
			end if

			If (len(trim(strFoererkort)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[foererkort] = '" & strFoererkort & "'"
			end if

			If (len(trim(strOppsigelsestid)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[oppsigelsestid] <= '" & strOppsigelsestid & "' "
			end if

			If (len(trim(strWorkType)) > 0) Then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[WorkType] = '" & strWorkType & "' "
			end if

			if (strCheckedInterviewed = "interviewed") then
				strSearchWhereSQL = strSearchWhereSQL & " AND [V].[Intervjudato] IS NOT NULL "
			end if

			if len(trim(strServiceAreas)) > 0 then
				aTjenesteomrader = split(strServiceAreas,",")
				if not IsEmpty(aTjenesteomrader) then
					NOFArrayItems = ubound(aTjenesteomrader)
					if (strSASearchType="AND") then
						For I = 0 To NOFArrayItems
							strSearchFromSQL = strSearchFromSQL & ", [VIKAR_TJENESTEOMRADE] AS [VT" & I & "]"
							strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VT" & I & "].VikarID = [V].VikarID"
							strSearchWhereSQL = strSearchWhereSQL  & " " & strSASearchType & " [VT" & I & "].[TomID] = '" &  trim(aTjenesteomrader(I)) & "'"
						Next
					elseif (strSASearchType="OR") then
						strSearchFromSQL = strSearchFromSQL & ", [VIKAR_TJENESTEOMRADE] AS [VT]"
						strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VT].[VikarID] = [V].[VikarID]"
						strSearchWhereSQL = strSearchWhereSQL  & " AND ("
						For I = 0 To NOFArrayItems
							if (I>0) then
								strSASearchType = "OR"
							else
								strSASearchType = ""
							end if
							strSearchWhereSQL = strSearchWhereSQL  & " " & strSASearchType & " [VT].TomID = '" &  trim(aTjenesteomrader(I)) & "'"
						Next
						strSearchWhereSQL = strSearchWhereSQL  & ")"
					end if
				end if
			end if


			
			if len(trim(strCategoryAreas)) > 0 then
				aCategory = split(strCategoryAreas,",")
				if not IsEmpty(aCategory) then
					NOFArrayItems = ubound(aCategory)
					if (strCategorySearchType="AND") then
						For I = 0 To NOFArrayItems
							strSearchFromSQL = strSearchFromSQL & ", [VIKAR_CATEGORY] AS [VCA" & I & "]"
							strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VCA" & I & "].VikarID = [V].VikarID"
							strSearchWhereSQL = strSearchWhereSQL  & " " & strCategorySearchType & " [VCA" & I & "].[CategoryID] = '" &  trim(aCategory(I)) & "'"
						Next
					elseif (strCategorySearchType="OR") then
						strSearchFromSQL = strSearchFromSQL & ", [VIKAR_CATEGORY] AS [VCA]"
						strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VCA].[VikarID] = [V].[VikarID]"
						strSearchWhereSQL = strSearchWhereSQL  & " AND ("
						For I = 0 To NOFArrayItems
							if (I>0) then
								strCategorySearchType = "OR"
							else
								strCategorySearchType = ""
							end if
							strSearchWhereSQL = strSearchWhereSQL  & " " & strCategorySearchType & " [VCA].CategoryID = '" &  trim(aCategory(I)) & "'"
						Next
						strSearchWhereSQL = strSearchWhereSQL  & ")"
					end if
				end if
			end if

			if len(trim(strBranches)) > 0 then
				aAvdelingskontor = split(strBranches,",")
				if not IsEmpty(aAvdelingskontor) then
					NOFArrayItems = ubound(aAvdelingskontor)
					if (strBranchSearchType="AND") then
						For I = 0 To NOFArrayItems
							strSearchFromSQL = strSearchFromSQL & ", [VIKAR_ARBEIDSSTED] AS [VAS" & I & "]"
							strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VAS" & I & "].[VikarID] = [V].[VikarID] "
							strSearchWhereSQL = strSearchWhereSQL  & " " & strBranchSearchType & " [VAS" & I & "].AvdelingskontorID = '" &  trim(aAvdelingskontor(I)) & "'"
						Next
					elseif (strBranchSearchType="OR") then
						strSearchFromSQL = strSearchFromSQL & ", [VIKAR_ARBEIDSSTED] AS [VAS]"
						strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VAS].[VikarID] = [V].[VikarID] "
						strSearchWhereSQL = strSearchWhereSQL  & " AND ("
						For I = 0 To NOFArrayItems
							if (I>0) then
								strBranchSearchType = "OR"
							else
								strBranchSearchType = ""
							end if
							strSearchWhereSQL = strSearchWhereSQL  & " " & strBranchSearchType & " [VAS].[AvdelingskontorID] = '" &  trim(aAvdelingskontor(I)) & "'"
						Next
						strSearchWhereSQL = strSearchWhereSQL  & ") "
					end if
				end if
			end if

			i = 1
			for each chk in request("hdnServiceArea")
			'Bygg opp fagkompetanse SQL
				strSearchFromSQL = strSearchFromSQL & ", [vikar_kompetanse] AS [VK" & i & "] "
				strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VK" & i & "].[VikarID] = [v].[VikarID] "
				strSAValue = trim(chk)
				lngSARating = trim(request("cboNivaa" & strSAValue))
				strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[K_TittelID]=" & strSAValue & " "
				if cint(lngSARating)>0 then
					strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[Relevant_WorkExperience]=" & lngSARating & " "
				end if
				lngSARating = trim(request("cboBrukerNivaa" & strSAValue))
				if cint(lngSARating)>0 then
						strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[Relevant_Education]=" & lngSARating & " "
				end if
				i = i + 1
			next

			for each chk in request("hdnProductArea")
			'Bygg opp produktkompetanse SQL
				strSearchFromSQL = strSearchFromSQL & ", [vikar_kompetanse] AS [VK" & i & "] "
				strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [VK" & i & "].[VikarID] = [v].[VikarID] "
				strSAValue = trim(chk)
				lngSARating = trim(request("dbxLevel" & strSAValue))
				strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[K_TittelID]=" & strSAValue & " "
				if cint(lngSARating)>0 then
					strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[k_levelID]=" & lngSARating & " "
				end if
				lngSARating = trim(request("cboProductRating" & strSAValue))
				if cint(lngSARating)>0 then
						strSearchWhereSQL = strSearchWhereSQL  & " AND [VK" & i & "].[Rangering]=" & lngSARating & " "
				end if
				i = i + 1
			next

			if len(strFreeTextSearch)> 0 then
				strSearchFromSQL = strSearchFromSQL & ", [CV] , [vikar_kompetanse] AS [VK" & i & "], [cv_Data] AS [DAT], [CV_References] AS [REF] "
				strSearchWhereRelationSQL = strSearchWhereRelationSQL & " AND [CV].[ConsultantID] = [V].[VikarID] " & _
				" AND [VK" & i & "].[VikarID] = [v].[VikarID] "	& _
				" AND [DAT].[CvID] = [CV].[CvID] " & _
				" AND [REF].[CvID] = [CV].[CvID] "
				strSearchWhereSQL = strSearchWhereSQL & " AND [CV].[Type] = 'C' "
				if strFreeTextSearchType="EXACT" then
					strSearchWhereSQL = strSearchWhereSQL & " AND ( " & _
					" [CV].[Key_Qualifications] like '%" & strFreeTextSearch & "%'" & _
					" OR [CV].[other_information] like '%" & strFreeTextSearch & "%'" & _
					" OR [VK" & i & "].[Kommentar] like '%" & strFreeTextSearch & "%'" & _
					" OR [DAT].[Place] like '%" & strFreeTextSearch & "%'" & _
					" OR [DAT].[Title] like '%" & strFreeTextSearch & "%'" & _
					" OR [DAT].[Description] like '%" & strFreeTextSearch & "%'" & _
					" OR [REF].[Name] like '%" & strFreeTextSearch & "%'" & _
					" OR [REF].[Title] like '%" & strFreeTextSearch & "%'" & _
					" OR [REF].[Comment] like '%" & strFreeTextSearch & "%'" & _
					" OR [REF].[Firma] like '%" & strFreeTextSearch & "%'" & _
					" ) "
				elseif strFreeTextSearchType="ALLWORDS" then
					Awords = split(strFreeTextSearch," ")
					NOFArrayItems = ubound(Awords)
					strSearchWhereSQL = strSearchWhereSQL & " AND ( "
					For n = 0 To NOFArrayItems
						if n > 0 then
							strSearchWhereSQL = strSearchWhereSQL & " OR"
						end if
						strSearchWhereSQL = strSearchWhereSQL  & _
						" [CV].[Key_Qualifications] like '%" & Awords(n) & "%'" & _
						" OR [CV].[other_information] like '%" & Awords(n) & "%'" & _
						" OR [VK" & i & "].[Kommentar] like '%" & Awords(n) & "%'" & _
						" OR [DAT].[Place] like '%" & Awords(n) & "%'" & _
						" OR [DAT].[Title] like '%" & Awords(n) & "%'" & _
						" OR [DAT].[Description] like '%" & Awords(n) & "%'" & _
						" OR [REF].[Name] like '%" & Awords(n) & "%'" & _
						" OR [REF].[Title] like '%" & Awords(n) & "%'" & _
						" OR [REF].[Comment] like '%" & Awords(n) & "%'" & _
						" OR [REF].[Firma] like '%" & Awords(n) & "%'"
					Next
					strSearchWhereSQL = strSearchWhereSQL & " ) "
				end if
			end if

			if (len(strFraDato) > 0 and len(strFraDato) > 0) then
				strSearchWhereSQL = strSearchWhereSQL & "AND [V].[VikarID] not in (" & _
				" SELECT [OV].[VikarID] FROM [OPPDRAG_VIKAR] AS [OV], [OPPDRAG] AS [O] " & _
				" WHERE [OV].[OPPDRAGID] = O.OPPDRAGID " & _
				" AND [OV].[StatusID] = 4 " & _
				" AND ([OV].[tildato] >= " & dbDate(strFraDato) & " AND [OV].[fradato] <= " & dbDate(strTilDato) & ")"& _
				" ) "
			end if

			'SortBY sQL
			strSearchSortbySQL = "[V].[etternavn],[v].[Fornavn] "
			
			strCheckedSortCategory0 = ""
			strCheckedSortCategory1 = "" 
			strCheckedSortCategory2 = ""
			strCheckedSortCategory3 = ""
			strCheckedSortCategory4 = ""
			strCheckedSortCategory5 = ""
			
			select case strSortByCategory
			case "0"	'Etternavn
				strSearchSortbySQL = "[V].[etternavn] "
				strCheckedSortCategory0 = "selected"
			case "1"	'VikarID
				strSearchSortbySQL = "[V].[VikarID] "
				strCheckedSortCategory1 = "selected"
			case "2"	'Ansattnummer
				strSearchSortbySQL = "[VA].[AnsattNummer] "
				strCheckedSortCategory2 = "selected"
			case "3"	'Intervjudato
				strSearchSortbySQL = "[V].[Intervjudato] "
				strCheckedSortCategory3 = "selected"
			case "4"	'Status
				strSearchSortbySQL = "[V].[StatusID] "
				strCheckedSortCategory4 = "selected"
			case "5"	'Ansvarlig
				strSearchSortbySQL = "[V].[AnsMedID] "
				strCheckedSortCategory5 = "selected"
			end select
			if strSortDesc = "desc" then
				strSearchSortbySQL = strSearchSortbySQL & " DESC"
			end if
					            						
			'DATA DELETION PRO@EC 2007/03/28 -> select except removed status
			strSearchSQL = "SELECT DISTINCT " & _
			strSearchSelectSQL & _
			" [VA].[AnsattNummer], [V].[VikarID],[v].[Fornavn],[v].[Etternavn], [v].[CVID]," & _
			" [A].[Adresse], [A].[Postnr], [A].[Poststed], [V].[Telefon], [V].[MobilTlf], [V].[EPost], [V].[Intervjudato], [V].[StatusID], [V].[AnsMedID] " & _
			" FROM [vikar] as [V],[ADRESSE] AS [A] " & _
			", [VIKAR_ANSATTNUMMER] AS [VA] " & _
			strSearchFromSQL & _
			" WHERE " & _
			" (NOT ([V].[StatusID] = '" & rsRemovedState("VikarstatusID") & "')) " & _
			" AND [V].[VikarID] = [A].[AdresserelID] " & _
			strSearchWhereRelationSQL & _
			" AND [A].[AdresseRelasjon]= '2' " & _
			" AND [A].[AdresseType] = '1' " & _
			strSearchWhereSQL & _
			" ORDER BY " & strSearchSortbySQL '[V].[etternavn],[v].[Fornavn] "


			set rsSoek = GetFirehoseRecSet(strSearchSQL, objCon)
			
			strExcludeServiceAreas = "'," & replace(request("hdnServiceArea")," ","") & ",'"
			strExcludeProductQualifications = "'," & replace(request("hdnProductArea")," ","") & ",'"

		elseif (request("hdnAddState") = "1") then
		'User added a experience / qualification
			strRatedArea = trim(request("cboFagkompetanse"))
			strExcludeServiceAreas = request("hdnServiceArea")
			if (len(strRatedArea) > 0) then
				strExcludeServiceAreas = "'," & strExcludeServiceAreas & ", " & strRatedArea & ",'"
			end if	
			strExcludeProductQualifications = "'," & replace(request("hdnProductArea")," ","") & ",'"			
		elseif (request("hdnAddState")="2") then
		'User remove one or more qualification
			strExcludeServiceAreas = "'," & replace(request("hdnServiceArea")," ","") & ",'"
			for each chk in request("chkRemove")
				'Remove qualification from ignorelist
				 if instr(1, strExcludeServiceAreas,"," & trim(chk) & ",") then
					strExcludeServiceAreas = replace(strExcludeServiceAreas, trim(chk) & ",","")
				 end if
			next
			strExcludeProductQualifications = "'," & replace(request("hdnProductArea")," ","") & ",'"
		elseif (request("hdnAddState")="3") then
		'User added a experience
			strRatedProductQualification = trim(request("cboProduktkompetanse")) 'List of existing product areas allready added
			strExcludeProductQualifications = request("hdnProductArea") 'Selected product area to be added
			if (len(strRatedProductQualification) > 0) then
				strExcludeProductQualifications = "'," & strExcludeProductQualifications & ", " & strRatedProductQualification & ",'"
			end if
			strExcludeServiceAreas = "'," & replace(request("hdnServiceArea")," ","") & ",'"
		elseif (request("hdnAddState")="4") then
		'User has removed one or more product qualifications
			strExcludeProductQualifications = "'," & replace(request("hdnProductArea")," ","") & ",'"
			for each chk in request("chkProductRemove")
				'Remove qualification from ignorelist
				 if instr(1, strExcludeProductQualifications, "," & trim(chk) & ",") then
					strExcludeProductQualifications = replace(strExcludeProductQualifications, trim(chk) & ",","")
				 end if
			next
			strExcludeServiceAreas = "'," & replace(request("hdnServiceArea"), " ", "") & ",'"
		end if
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<title>Vikar søk</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<link type="text/css" rel="stylesheet" href="css/main.css" title="default style">
		<link type="text/css" rel="stylesheet" href="css/print.css" title="default print style" media="print">
		<script type="text/javascript" src="js/contentMenu.js"></script>
		<script language='javascript' src='js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>		
		<script language='javascript' src='js/ecajax.js'></script>	
		<script language="javascript">
			var oWndHandle = null;
			document.onkeydown = shortKey;			

			function VisCV(nVikarID) {
				window.open('vikarCVutskrift.asp?VikarID='+nVikarID);
			}

			function ChangeSearch()
			{
				window.location = "HTMLSearch.asp"
			}
			function AddQualification()
			{
				document.all.hdnAddState.value="1";
				document.frmConsultantSearch.submit();
			}

			function RemoveQualification()
			{
				document.all.hdnAddState.value="2";
				document.frmConsultantSearch.submit();
			}
			function AddProductQualifications()
			{
				document.all.hdnAddState.value="3";
				document.frmConsultantSearch.submit();
			}

			function RemoveProductQualifications()
			{
				document.all.hdnAddState.value="4";
				document.frmConsultantSearch.submit();
			}

			function MultipleAcceptance(obj,displayAlert)
			{
				var count=0;
					
				if(obj.checked && displayAlert)
					alert('Vikaren jobber I et annet oppdrag I samme periode. Vil du fortsette?');
					
				for (var i=0; i<document.frmConsultantSearch.elements.length ;i++) {
				
					if(document.frmConsultantSearch.elements[i].name == "tbxID")
					{
						var object = document.frmConsultantSearch.elements[i];
						if(object.checked)
						{
							count = count + 1;
						}
					}
					
				
				}

				
				if (count > 1)
				{
					document.frmConsultantSearch.btnAxcept.disabled="disabled";
				}
				else
				{
					<% If oppdragVikarStatus <> 1 Then %>
						document.frmConsultantSearch.btnAxcept.disabled=false;
					<% Else %>
						document.frmConsultantSearch.btnAxcept.disabled="disabled";
					<% End If %>
					
				}			


			}			
			
			function ViewPage(page)
			{
				if(page == 'pre')					
					document.all.hdnCurrentPage.value = parseInt(document.all.hdnCurrentPage.value) -1;
				else if(page == 'next')
					document.all.hdnCurrentPage.value = parseInt(document.all.hdnCurrentPage.value) + 1;
				else if(page == 'last')
					document.all.hdnCurrentPage.value = parseInt(document.all.hdnLastPage.value);
				else if(page == 'first')
					document.all.hdnCurrentPage.value = 1;									
				document.all.frmConsultantSearch.submit();
			}
			
			function PopHelp(nhelpId)
			{
				if (oWndHandle && oWndHandle.open && !oWndHandle.closed)
				{
					oWndHandle.location = "help/Help"+nhelpId+".html";
					oWndHandle.focus();
				}else
				{
					oWndHandle = window.open("help/Help"+nhelpId+".html","Hjelp","height=200, width=400, scrollbars=yes, resizable=no, status=yes, toolbar=no, menubar=no, location=no")
					oWndHandle.focus();
				};
				event.cancelBubble=true;
			}

			function shortKey(e) 
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;
			
				if (event.keyCode == 13)
				{
					document.all.frmConsultantSearch.submit();
				}
				<% 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>				
					if (modKey && modKey2 && keyChar == "S")
					{	
						parent.frames[funcFrameIndex].location=("/xtra/VikarSoek.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
					%>					
					if (modKey && modKey2 && keyChar == "Y")
					{	
						parent.frames[funcFrameIndex].location=("/xtra/VikarNy.asp");
					}
					<% 
				End If 
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>					
					if (modKey && modKey2 && keyChar == "W")
					{	
						parent.frames[funcFrameIndex].location = ("/xtra/jobb/SuspectList.asp");
					}
					<% 
				End If 
				%>					
			}
			//Her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
			
			//PDF
			function SortSearchResult()
			{
				//alert('Test')
				document.frmConsultantSearch.submit();
			}
				
			var handleCallback = function (result) 
			{
				if(result) 
				{
					var str = "<select ID='cboManager' NAME='cboManager' onkeydown='typeAhead()'><option VALUE='0'></option>";
					str = str.concat(result,"</SELECT>")
					document.getElementById('divResponsible').innerHTML = str;
				}
			};
		
			function ShowAll(ansID)
			{
				var url;
				if(document.all.chkShowAllRes.checked)			
					url = "Callback.asp?FuncName=GetAllCoWorkersAsOptionList&SelectedID=";			
				else			
					url = "Callback.asp?FuncName=GetActiveCoWorkersAsOptionList&SelectedID=";							
				url = url.concat(ansID);
				
				asynchronousCall(url,handleCallback);
			
			}
				
			function showDialog(ddl,hdnControl)
			{
			    showModalDialog('WebUI/EmailEditor.aspx?TemplateId='+ document.getElementById(ddl).value + '&VikarIds='+ document.getElementById(hdnControl).value,null,'dialogWidth:800px;dialogHeight:480px;location:0;toolbar:0;resizable=0;status:0;menubar=0');
			}
			
			function addVikarId(chk,control,vikarId)
			{   	
				var count=0;		
			    	var Ids = document.getElementById(control).value;
			        if(document.getElementById(chk).checked && Ids.indexOf(vikarId) == -1)
			            Ids += vikarId + ",";            
			        else
			           Ids =  Ids.replace(vikarId+ ',','');
			
			        document.getElementById(control).value = Ids;			        
			       
				for (var i=0; i<document.frmConsultantSearch.elements.length ;i++) {
				
					if(document.frmConsultantSearch.elements[i].name == "mailCHK")
					{
						var object = document.frmConsultantSearch.elements[i];
						if(object.checked)						
							count = count + 1;						
					}									
				}
				if (count > 50)
				{				
					document.frmConsultantSearch.cboEmailTemplates.disabled=true;				
					document.frmConsultantSearch.imgSendMail.disabled=true;
					document.frmConsultantSearch.imgSendMail.src="images/envelop_dis.gif";
				}
				else
				{
					document.frmConsultantSearch.cboEmailTemplates.disabled=false;			
					document.frmConsultantSearch.imgSendMail.disabled=false;
					document.frmConsultantSearch.imgSendMail.src="images/envelop.gif";
				}
			}
				
		</script>
		<script language="javaScript" src="Js/javascript.js"></script>
	</head>
	<%
	if ((postedPage = "1") AND (request("hdnAddState") = "0")) then
		strLoad = "onLoad='javascript:document.all.Result.scrollIntoView(true);'"
	else
		strLoad = "onLoad='javascript:document.all.txtAnsattID.focus();'"
	end if
	%>
	<body <%=strLoad%>>
		<form action="vikarSoek.asp" method="post" id="frmConsultantSearch" name="frmConsultantSearch">
			<input type="hidden" name="hdnPosted" id="hdnPosted" value="1">
			<input type="hidden" name="hdnAddState" id="hdnAddState" value="0">
			<input type="hidden" name="tbxOppdragID" id="tbxOppdragID" value="<%=lngoppdragID%>">
			<input type="hidden" name="tbxFirmaID" id="tbxFirmaID" value="<%=lngFirmaID%>">		
			<div class="pageContainer" id="pageContainer">
				<div class="contentHead1">
					<h1><%=heading%></h1>
					<div class="contentMenu">
						<input name='hdnCurrentPage' id='hdnCurrentPage' TYPE='hidden' VALUE='<%= nPage %>'>
						<table cellpadding="0" cellspacing="0" width="96%" ID="Table1">
							<tr>								
								<td>
									<table cellpadding="0" cellspacing="2" ID="Table2">
										<tr>
											<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												Søk i
												<select ID="cboSearchIn" NAME="cboSearchIn" onchange="javascript:ChangeSearch()">
													<option selected>Database</option>
													<option>Dokumentarkiv</option>
												</select>
											</td>
											<td class="menu" id="menu7" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<img src="/xtra/images/icon_search.gif" width="18" height="15" alt="" align="absmiddle">
												<a onClick="javascript:document.all.hdnCurrentPage.value = 1;document.all.frmConsultantSearch.submit();" href="#" title="Søke etter vikarer">Utfør søk</a>
											</td>
											<td class="menu" id="menu8" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a onClick="javascript:window.location='vikarSoek.asp';" href="#" title="Blanker ut alle feltene">Blank ut</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="right">
									<!--#include file="Includes/contentToolsMenu.asp"-->
								</td>
							</tr>
						</table>
					</div>
				</div>
				<div class="content">
					<table ID="Table3">
						<col width="22%">
						<col width="22%">
						<col width="28%">
						<col width="28%">
						<tr>
							<td align="left">Ansattnummer:&nbsp;<input type="text" class="sizeC" ID="txtAnsattID" NAME="txtAnsattID" size="8" maxlength="8" value="<%=lngAnsattID%>"></td>
							<td align="left">Vikarnummer:&nbsp;<input type="text" class="sizeC" ID="txtVikarID" NAME="txtVikarID" size="8" maxlength="8" value="<%=lngVikarID%>"></td>
							<td align="left">Etternavn:&nbsp;<input type="text" ID="txtEtternavn" NAME="txtEtternavn" size="25" maxlength="50" value="<%=strSoekEtterNavn%>"></td>
							<td align="left">Fornavn:&nbsp;<input type="text" ID="txtForNavn" NAME="txtForNavn" size="25" maxlength="50" value="<%=strSoekForNavn%>"></td>
						</tr>
					</table>
				</div>
				<div class="contentHead contentHeadOpen" id="toggleView101" onclick="toggle('view101', this.id);" onMouseOver="toggleOver(this.id);" onMouseOut="toggleOut(this.id);">
					<h2 style="position:inline; width:94.8%;">Tilhørighet og status</h2>
					<div style="float:right; width:5.2%; margin:-19px 0 0 0;"><img src="images/icon_help.gif" onClick="PopHelp(1);"></div>
				</div>
				<div class="content" id="view101" style="display:<%=strDisplayStatus%>;">
					<table ID="Table4">
						<col width="13%">
						<col width="13%">
						<col width="5%">
						<col width="23%">
						<col width="23%">
						<col width="23%">
						<tr>
							<td align="left">
							<div>Status:&nbsp;
								<div class="divContainer">									
									<%									
									set rsStatus = GetFirehoseRS("exec [dbo].[GetAllConsultantStatus]", objCon)									
									if (not rsStatus.EOF) then
										if Not IsEmpty(aConsultantStatus) then
											NOFArrayItems = ubound(aConsultantStatus)
											Dim tempBool
											while (not rsStatus.EOF)
												tempBool = false
												For I = 0 To NOFArrayItems													
													if (clng(trim(aConsultantStatus(I)))=clng(rsStatus.fields("VikarStatusID"))) then
														%>
														<div><input class="checkbox" Checked TYPE='CHECKBOX' ID='chkConsultantStatus' NAME='chkConsultantStatus' VALUE='<%=rsStatus.fields("VikarStatusID")%>'><%=rsStatus.fields("VikarStatus")%></div>
														<!--<option <%=strCboSelected%> value="<%=rsStatus.fields("VikarStatusID")%>"><%=rsStatus.fields("VikarStatus")%></option>-->
														<%	
														tempBool = true														
														Exit For
													end if													
												Next
												
												if Not tempBool then
												%>
												<div><input class="checkbox" TYPE='CHECKBOX'  ID='chkConsultantStatus' NAME='chkConsultantStatus' VALUE='<%=rsStatus.fields("VikarStatusID")%>'><%=rsStatus.fields("VikarStatus")%></div>
												<%
												else
												
												%>
												  
												<%
												end if
												rsStatus.movenext
											wend
										else
											while (not rsStatus.EOF)											
												%>
												<div><input class="checkbox" TYPE='CHECKBOX' ID='chkConsultantStatus' NAME='chkConsultantStatus' VALUE='<%=rsStatus.fields("VikarStatusID")%>'  <%if(InStr(1, request("chkConsultantStatus"),rsStatus.fields("VikarStatusID"))>0) then Response.write("CHECKED")  End If %>><%=rsStatus.fields("VikarStatus")%> </div>
												<%
												rsStatus.movenext
											wend	
										end if
									end if
									rsStatus.close
									set rsStatus = nothing
									%>								
								</div></div>
								<br />								
								Intervjuet:
								<input class="checkbox" type="checkbox" <%if strCheckedInterviewed = "interviewed" then %>checked<%  end if %> id="chkInterviewed" name="chkInterviewed" value="interviewed">
							</td>
							<td align="left" valign="top">Ansvarlig:<br>
								<div id="divResponsible">
								<select ID="cboManager" NAME="cboManager" onkeydown="typeAhead()">
									<option VALUE="0"></option>
									<%
      									' Get ansvarlig medarbeider
 										response.write GetCoWorkersAsOptionList(lngManagerID) '-1
 									%>
								</select>
								</div>
								
							</td>
							
							<td align="left" valign="top"><br>
								<input id='chkShowAllRes' name='chkShowAllRes' type='checkbox' class='checkbox' Value='1' onClick="ShowAll(<%=lngManagerID%>);">Vis Alle
							</td>
							
							<td align="left" valign="top">Avdelingskontor:<br>
								<select name="cboBranchSearchType" class="LongStyle" id="cboBranchSearchType">
								<%
								if (request("cboBranchSearchType")="AND") then
								%>
								<option value="OR">Minst en (OR-søk)</option>
								<option value="AND" selected>Alle (AND-søk)</option>
								<%
								else
								%>
								<option value="OR" selected>Minst en (OR-søk)</option>
								<option value="AND">Alle (AND-søk)</option>
								<%
								end if
								%>
								</select><br>
								<select ID="cboBranch" NAME="cboBranch" class="LongStyle" size="8" multiple>
									<%
									set rsBranches = GetFirehoseRS("exec [dbo].[GetAllBranchesForDropDown]", objCon)
									strBranches = "," & trim(replace(strBranches," ","")) & ","
									if (not rsBranches.EOF) then
										while (not rsBranches.EOF)
											if (instr(1, strBranches,"," & trim(rsBranches.fields("id")) & ",")) then
												strCboSelected = "selected"
											else
												strCboSelected = ""
											end if
										%>
											<option <%=strCboSelected%> value="<%=rsBranches.fields("id")%>"><%=rsBranches.fields("navn")%></option>
										<%
											rsBranches.movenext
										wend
									end if
									rsBranches.close
									set rsBranches = nothing
									%>
								</select>
							</td>
							
							<td align="left" valign="top">Tjenesteområde:<br>
								<select name="cboSASearchType" class="LongStyle" id="cboSASearchType">
								<%
								if (request("cboSASearchType")="AND") then
								%>
								<option value="OR">Minst en (OR-søk)</option>
								<option value="AND" selected>Alle (AND-søk)</option>
								<%
								else
								%>
								<option value="OR" selected>Minst en (OR-søk)</option>
								<option value="AND">Alle (AND-søk)</option>
								<%
								end if
								%>
								</select><br>
								<select ID="cboServiceArea" NAME="cboServiceArea" class="LongStyle" size="8" multiple>
									<%
									set rsServiceAreas = GetFirehoseRS("exec [dbo].[GetAllServiceAreasForDropDown]", objCon)
									strServiceAreas = "," & trim(replace(strServiceAreas," ","")) & ","
									if (not rsServiceAreas.EOF) then
										while (not rsServiceAreas.EOF)
											if (instr(1, strServiceAreas,"," & trim(rsServiceAreas.fields("tomid")) & ",")) then
												strCboSelected = "selected"
											else
												strCboSelected = ""
											end if
											%>
											<option <%=strCboSelected%> value="<%=rsServiceAreas.fields("tomid")%>"><%=rsServiceAreas.fields("navn")%></option>
											<%
											rsServiceAreas.movenext
										wend
									end if
									rsServiceAreas.close
									set rsServiceAreas = nothing
									%>
								</select>
							</td>
							
							<!-- Category Drop down -->
							<td align="left" valign="top">Kategori:<br>
								<select name="cboCategorySearchType" class="LongStyle" id="cboCategorySearchType">
								<%
								if (request("cboCategorySearchType")="AND") then
								%>
								<option value="OR">Minst en (OR-søk)</option>
								<option value="AND" selected>Alle (AND-søk)</option>
								<%
								else
								%>
								<option value="OR" selected>Minst en (OR-søk)</option>
								<option value="AND">Alle (AND-søk)</option>
								<%
								end if
								%>
								</select><br>								
								
								<select ID="cboCategory" NAME="cboCategory" class="LongStyle" size="8" multiple>
									<%
									set rsCategoryAreas = GetFirehoseRS("exec [dbo].[spECGetAllCategoriesList]", objCon)
									strCategoryAreas = "," & trim(replace(strCategoryAreas," ","")) & ","
									
									if (not rsCategoryAreas.EOF) then
										while (not rsCategoryAreas.EOF)
											if (instr(1, strCategoryAreas,"," & trim(rsCategoryAreas.fields("CategoryID")) & ",")) then
												strCboSelected = "selected"
											else
												strCboSelected = ""
											end if
											%>
											<option <%=strCboSelected%> value="<%=rsCategoryAreas.fields("CategoryID")%>"><%=rsCategoryAreas.fields("Name")%></option>
											<%
											rsCategoryAreas.movenext
										wend
									end if
									rsCategoryAreas.close
									set rsCategoryAreas = nothing
									
									
									%>
								</select>
								
							</td>							
						</tr>
					</table>
				</div>
				<div class="contentHead contentHeadClosed" id="toggleView102" onclick="toggle('view102', this.id);" onMouseOver="toggleOver(this.id);" onMouseOut="toggleOut(this.id);" title="klikk for å åpne/lukke">
					<h2 style="position:inline; width:94.8%;">Andre opplysninger</h2>
					<div style="float:right; width:5.2%; margin:-18px 0 0 0;"><img src="images/icon_help.gif" onClick="PopHelp(2);"></div>
				</div>
				<div class="content" id="view102" style="display:<%=strDisplayOtherInfo%>;">
					<table ID="Table5" width="100%">
						<col width="25%">
						<col width="25%">
						<col width="25%">
						<col width="25%">
						<tr>
							<td align="left">Postnummer:&nbsp;<input type="text" ID="txtPostnr" NAME="txtPostnr" size="4" maxlength="4" value="<%=strPostnr%>"></td>
							<td align="left">
								Førerkort:&nbsp;<select id="cbodriverLicence" NAME="cbodriverLicence">
									<option></option>
									<option <%=strCheckedDriverslicinseNo%> value="0">Nei</option>
									<option <%=strCheckedDriverslicinseYes%> value="1">Ja</option>
								</select>
							</td>
							<td align="left">
								Oppsigelsestid mindre enn:&nbsp;
								<select ID="cboOppsigelse" NAME="cboOppsigelse">
									<option value=""></option>
									<option <%=strCheckedOppsigelse0%> value="0">Ingen</option>
									<option <%=strCheckedOppsigelse1%> value="1">14 dager</option>
									<option <%=strCheckedOppsigelse2%> value="2">1 m&aring;ned</option>
									<option <%=strCheckedOppsigelse3%> value="3">2 m&aring;neder</option>
									<option <%=strCheckedOppsigelse4%> value="4">3 m&aring;neder</option>
									<option <%=strCheckedOppsigelse5%> value="5">3 m&aring;neder eller mer</option>
								</select>
							</td>
							<td align="left">
								Stillingsbrøk:&nbsp;
								<select ID="cboWorkType" NAME="cboWorkType">
									<option value=""></option>
									<option <%=strCheckedWorkType0%> value="0">Ikke angitt</option>
									<option <%=strCheckedWorkType1%> value="1">Kun fulltid</option>
									<option <%=strCheckedWorkType2%> value="2">Kun deltid</option>
									<option <%=strCheckedWorkType3%> value="3">Fulltid og deltid</option>
								</select>
							</td>
						</tr>
					</table>
				</div>
				<div class="contentHead contentHeadOpen" id="toggleView103" onclick="toggle('view103', this.id);" onMouseOver="toggleOver(this.id);" onMouseOut="toggleOut(this.id);" title="klikk for å åpne/lukke">
					<h2 style="position:inline; width:94.8%;">Produkt- og fagkompetanse</h2>
					<div style="float:right; width:5.2%; margin:-18px 0 0 0;"><img src="images/icon_help.gif" onClick="PopHelp(3);"></div>
				</div>
				<div class="content" id="view103" style="display:<%=strDisplayQualifications%>;">
					<h3>Fagkompetanse</h3>
					<table cellpadding="0" cellspacing="0" width="96%">
						<tr>
							<td>
								<table ID="Table7">
									<tr>
										<td valign="top">
											<select name="cboFagkompetanse" id="cboFagkompetanse" size="8" style="width:240px;" onkeydown="typeAhead()">
											<%
											set rsServiceAreas = GetFirehoseRS("exec [dbo].[GetQualificationsExcludeSpecified] " & strExcludeServiceAreas, objCon)
											if (not rsServiceAreas.EOF) then
												while (not rsServiceAreas.EOF)
												%>
													<option value="<%=rsServiceAreas.fields("K_TittelID")%>"><%=rsServiceAreas.fields("FagOmrade") & " - " & rsServiceAreas.fields("ktittel")%></option>
												<%
													rsServiceAreas.movenext
												wend
											end if
											rsServiceAreas.close
											set rsServiceAreas = nothing
											%>
											</select>
										</td>
										<td valign="middle">&nbsp;<input type="button" onClick="javascript:AddQualification()" ID="btnAdd" NAME="btnAdd" value=">" title="Legg til fagkvalifikasjoner"><br><br>
											&nbsp;<input type="button" onClick="javascript:RemoveQualification()" ID="btnRemove" NAME="btnRemove" value="<" title="fjern fagkvalifikasjoner">
											&nbsp;
										</td>
										<td valign=top">
											<div class="Searchcontainer listing">
												<table ID="Table8">
													<tr>
														<th width="5%">&nbsp;</th>
														<th width="75%">Fagkompetanse</th>
														<th width="10%">Erfaring</th>
														<th width="10%">Utdannelse</th>
													</tr>
													<%
													if (request("hdnAddState")="0") OR (request("hdnAddState")="1") OR (request("hdnAddState")="2") OR (request("hdnAddState")="3") OR (request("hdnAddState")="4") then
														'New qualification
														if (len(strRatedArea) > 0) then
															strSAValue = strRatedArea

															set rsSA = GetFirehoseRS("exec [dbo].[GetQualificationByID] " & strSAValue, objCon)
															strSAName = rsSA.fields("FagOmrade").value & " - " &  rsSA.fields("ktittel").value
															rsSA.close
															set rsSA = nothing
															%>
															<tr>
																<td>
																	<input type="hidden" ID="Hidden1" NAME="hdnServiceArea" value="<%=strSAValue%>">
																	<input type="checkbox" class="checkbox" id="Checkbox1" name="chkRemove" value="<%=strSAValue%>">
																</td>
																<td><%=strSAName%>&nbsp;</td>
																<td>
																	<select id="cboNivaa2" name="cboNivaa<%=strSAValue%>">
																		<option value="0"></option>
																		<option value="1">Lite</option>
																		<option value="2">Noe</option>
																		<option value="3">Mye</option>
																	</select>
																</td>
																<td>
																	<select id="cboBrukerNivaa2" name="cboBrukerNivaa<%=strSAValue%>">
																		<option value="0"></option>
																		<option value="1">Lite</option>
																		<option value="2">Noe</option>
																		<option value="3">Mye</option>
																	</select>
																</td>
															</tr>
														<%
														end if
														strExcludeServiceAreas = replace(strExcludeServiceAreas," ","")
														for each chk in request("hdnServiceArea")
														'Recreate old qualifications
															strSAValue = trim(chk)
															if instr(1, strExcludeServiceAreas, "," & strSAValue & ",")>0 then

																set rsSA = objCon.execute("exec [dbo].[GetQualificationByID] " & strSAValue)
																strSAName = rsSA.fields("FagOmrade").value & " - " &  rsSA.fields("ktittel").value
																set rsSA = nothing
																strSARating1 = ""
																strSARating2 = ""
																strSARating3 = ""
																lngSARating = trim(request("cboNivaa" & strSAValue))
																if len(lngSARating)>0 then
																	if clng(lngSARating)=1 then
																		strSARating1 = "selected"
																	elseif clng(lngSARating)=2 then
																		strSARating2 = "selected"
																	elseif clng(lngSARating)=3 then
																		strSARating3 = "selected"
																	end if
																end if
																lngSARating = trim(request("cboBrukerNivaa" & strSAValue))
																if len(lngSARating)>0 then
																	if clng(lngSARating)=1 then
																		strSALevel1 = "selected"
																	elseif clng(lngSARating)=2 then
																		strSALevel2 = "selected"
																	elseif clng(lngSARating)=3 then
																		strSALevel3 = "selected"
																	end if
																end if
																%>
																<tr>
																	<td>
																		<input type="hidden" ID="hdnServiceArea" NAME="hdnServiceArea" value="<%=strSAValue%>">
																		<input type="checkbox" class="checkbox" id="chkRemove" name="chkRemove" value="<%=strSAValue%>">
																	</td>
																	<td><%=strSAName%>&nbsp;</td>
																	<td>
																		<select id="cboNivaa" name="cboNivaa<%=strSAValue%>">
																			<option value="0"></option>
																			<option <%=strSARating1%> value="1">Lite</option>
																			<option <%=strSARating2%> value="2">Noe</option>
																			<option <%=strSARating3%> value="3">Mye</option>
																		</select>
																	</td>
																	<td>
																		<select id="cboBrukerNivaa" name="cboBrukerNivaa<%=strSAValue%>">
																			<option value="0"></option>
																			<option <%=strSALevel1%> value="1">Lite</option>
																			<option <%=strSALevel2%> value="2">Noe</option>
																			<option <%=strSALevel3%> value="3">Mye</option>
																		</select>
																	</td>
																</tr>
																<%
															end if
														next
													end if
													%>
												</table>
											</div>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					<h3>Produktkompetanse</h3>
					<table ID="Table9">
						<tr>
							<td>
								<table ID="Table10">
									<tr>
										<td valign=top">
											<select name="cboProduktkompetanse" id="Select1" size="8" style="width:240px;" onkeydown="typeAhead()">
											<%
											set rsProductQualifications = GetFirehoseRS("exec [dbo].[GetProductQualificationsExcludeSpecified] " & strExcludeProductQualifications, objCon)
											if (not rsProductQualifications.EOF) then
												while (not rsProductQualifications.EOF)
												%>
													<option value="<%=rsProductQualifications.fields("K_TittelID")%>"><%=rsProductQualifications.fields("ktittel")%></option>
												<%
													rsProductQualifications.movenext
												wend
											end if
											rsProductQualifications.close
											set rsProductQualifications = nothing
											%>
											</select>
										</td>
										<td  valign="middle">&nbsp;<input type="button" onClick="javascript:AddProductQualifications()" ID="Button1" NAME="btnAdd"  value=">" title="Legg til produktkvalifikasjoner"><br><br>
											&nbsp;<input type="button" onClick="javascript:RemoveProductQualifications()" ID="Button2" NAME="btnRemove" value="<" title="fjern produktkvalifikasjoner">
											&nbsp;
										</td>
										<td  valign=top">
											<div class="Searchcontainer listing">
												<table ID="Table11">
													<tr>
														<th width="5%">&nbsp;</th>
														<th width="75%">Produktkompetanse</th>
														<th width="10%">Brukernivå</th>
														<th width="10%">Kursnivå</th>
													</tr>
													<%
													if (request("hdnAddState")="0") OR (request("hdnAddState")="1") OR (request("hdnAddState")="2") OR (request("hdnAddState")="3") OR (request("hdnAddState")="4") then
														'New product qualification
														if (len(strRatedProductQualification) > 0) then
															strSAValue = strRatedProductQualification
															set rsSA = GetFirehoseRS("exec [dbo].[GetProductQualificationByID] " & strSAValue, objCon)
															strSAName = rsSA.fields("ktittel").value
															rsSA.close
															set rsSA = nothing
															%>
															<tr>
																<td>
																	<input type="hidden" ID="Hidden2" NAME="hdnProductArea" value="<%=strSAValue%>">
																	<input type="checkbox" class="checkbox" id="Checkbox2" name="chkProductRemove" value="<%=strSAValue%>">
																</td>
																<td><%=strSAName%>&nbsp;</td>
																<td>
																	<select id="Select2" name="cboProductRating<%=strSAValue%>" style="width:auto;">
																		<option value="0"></option>
																			<%
																			rsProductRank.movefirst
																			While (NOT rsProductRank.EOF)
																				strLevel = rsProductRank("K_rangeringID")
																				%>
																		<option value="<%=strLevel%>"><%=rsProductRank.fields("K_rangering").value%></option>
																				<%
																				rsProductRank.movenext
																			wend
																			%>
																	</select>
																</td>
																<td>
																	<select name="dbxLevel<%=strSAValue%>" id="dbxLevel<%=strSAValue%>">
																		<option value="0"></option>
																	<%
																	rsLevel.movefirst
																	While (NOT rsLevel.EOF)
																		strLevel = rsLevel("K_LevelID")
																		%><option VALUE="<%=strLevel%>"><%=rsLevel("KLevel")%></option><%
																		rsLevel.MoveNext
																	Wend
																	%>
																	</select>
																</td>
															</tr>
														<%
														end if
														strExcludeProductQualifications = replace(strExcludeProductQualifications," ","")
														for each chk in request("hdnProductArea")
														'Recreate old qualifications
															strSAValue = trim(chk)
															if instr(1, strExcludeProductQualifications, "," & strSAValue & ",")>0 then

																set rsSA = GetFirehoseRS("exec [dbo].[GetProductQualificationByID] " & strSAValue, objCon)
																strSAName =  rsSA.fields("ktittel").value
																rsSA.close
																set rsSA = nothing
																lngSARating = trim(request("cboProductRating" & strSAValue))
																if len(lngSARating) = 0 then
																	lngSARating = 0
																end if
																intLevelid = trim(request("dbxLevel" & strSAValue))
																if len(intLevelid) = 0 then
																	intLevelid = 0
																end if
																%>
																<tr>
																	<td>
																		<input type="hidden" ID="Hidden3" NAME="hdnProductArea" value="<%=strSAValue%>">
																		<input type="checkbox" class="checkbox" id="Checkbox3" name="chkProductRemove" value="<%=strSAValue%>">
																	</td>
																	<td><%=strSAName%>&nbsp;</td>
																	<td>
																		<select id="Select4" name="cboProductRating<%=strSAValue%>">
																			<option value="0"></option>
																			<%
																			rsProductRank.movefirst
																			While (NOT rsProductRank.EOF)
																				strLevel = rsProductRank("K_rangeringID")
																				if rsProductRank.fields("K_rangeringID").value=cint(lngSARating) then
																					strCboSelected = "SELECTED"
																				Else
																					strCboSelected = ""
																				End If
																				%>
																				<option <%=strCboSelected%> value="<%=strLevel%>"><%=rsProductRank.fields("K_rangering").value%></option>
																				<%
																				rsProductRank.movenext
																			wend
																			%>
																		</select>
																	</td>
																	<td>
																		<select name="dbxLevel<%=strSAValue%>" ID="Select3">
																			<option value="0"></option>
																			<%
																			rsLevel.movefirst
																			While (NOT rsLevel.EOF)
																				strLevel = rsLevel("K_LevelID")
																				If cint(rsLevel("K_LevelID")) = cint(intLevelid) Then
																					strCboSelected = "SELECTED"
																				Else
																					strCboSelected = ""
																				End If

																				%><option VALUE="<%=strLevel%>" <%=strCboSelected%>><%=rsLevel("KLevel")%></option><%
																				rsLevel.MoveNext
																			Wend
																			%>
																		</select>
																	</td>
																</tr>
																<%
															end if
														next
													end if
													%>
												</table>
											</DIV>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</div>
				<div class="contentHead contentHeadClosed" id="toggleView104" onclick="toggle('view104', this.id);" onMouseOver="toggleOver(this.id);" onMouseOut="toggleOut(this.id);" title="klikk for å åpne/lukke">
					<h2 style="position:inline; width:94.8%;">Fritekstsøk i CV</h2>
					<div style="float:right; width:5.2%; margin:-18px 0 0 0;"><img src="images/icon_help.gif" onClick="PopHelp(4);"></div>
				</div>
				<div class="content" id="view104" style="display:<%=strDisplayFreeText%>;">
					Fritekstsøk:
					<select NAME="dbxFreeTextSearchType" ID="Select5">
							<%
							if (strFreeTextSearchType="EXACT") then
								%>
								<option VALUE="ALLWORDS">Alle ordene</option>
								<option SELECTED VALUE="EXACT">Eksakt frase</option>
								<%
							else
								%>
								<option SELECTED VALUE="ALLWORDS">Alle ordene</option>
								<option VALUE="EXACT">Eksakt frase</option>
								<%
							end if
							%>
					</select>
					<input type="text" ID="txtFreeTextSearch" NAME="txtFreeTextSearch" value="<%=strFreeTextSearch%>">
				</div>
				
				<div class="contentHead contentHeadClosed" id="toggleView105" onclick="toggle('view105', this.id);" onMouseOver="toggleOver(this.id);" onMouseOut="toggleOut(this.id);" title="klikk for å åpne/lukke">
					<h2 style="position:inline; width:94.8%;">Periode</h2>
					<div style="float:right; width:5.2%; margin:-18px 0 0 0;"><img src="images/icon_help.gif" onClick="PopHelp(5);"></div>
				</div>
				
				<div>
					<table><tr><td>&nbsp;</td></tr></table>
					<br/>
				</div>
				
				<div class="content" id="view105" style="display:<%=strDisplayCourseSpecific%>;">
					<table>
						<tr>
							<td width="40%">Ledig fom. dato:<input type="text" class="sizeC" ID="dtFromDate" NAME="dtFromDate" ONBLUR="dateCheck(this.form, this.name)" value="<%=strFraDato%>"></td>
							<td width="60%">Ledig tom. dato:<input type="text" class="sizeC" ID="dtToDate" NAME="dtToDate" ONBLUR="dateCheck(this.form, this.name)" value="<%=strTilDato%>"></td>
						</tr>
					</table>
				</div>
				<%
				if ((postedPage ="1") AND (request("hdnAddState")="0")) then
				%>
				<a id="Result"></a>
				<br>
					<div class="contentHead1">
					<table width="85%">
					<tr>
					<td>	<h1>Søkeresultat</h1> </td>										
					<td style="vertical-align:middle;">
					Sorter på:
						<select name="cboSortCategory" id="cboSortCategory" style="width: 150px" onchange="SortSearchResult()">
					            <option <%=strCheckedSortCategory0%> value="0">Etternavn</option>
					            <option <%=strCheckedSortCategory1%> value="1">VikarID</option>
					            <option <%=strCheckedSortCategory2%> value="2">Ansattnummer</option>
					            <option <%=strCheckedSortCategory3%> value="3">Intervjudato</option>
					            <option <%=strCheckedSortCategory4%> value="4">Status</option>
					            <option <%=strCheckedSortCategory5%> value="5">Ansvarlig</option>
					        </select>
					        <%if (strSortDesc="desc") then
							%>
							<input class="checkbox" type="checkbox" id="chkSortDesc" name="chkSortDesc" value="desc" onclick="SortSearchResult()" checked/>Desc</td>	
							<%
						else
							%>
							<input class="checkbox" type="checkbox" id="chkSortDesc" name="chkSortDesc" value="desc" onclick="SortSearchResult()" />Desc</td>
							<%
						end if
						%>
						<%
						if (blnFromJob	= false) then						
							%>
							<td style="vertical-align:middle;">
							<%						
							Response.write("&nbsp;&nbsp;&nbsp;&nbsp;Send E-post:&nbsp;")
							%>
							<select name="cboEmailTemplates" id="cboEmailTemplates">
								<%
								set rsEmailTemplates = GetFirehoseRS("exec [dbo].[spECGetAllEmailTemplates]", objCon)
								if (not rsEmailTemplates.EOF) then
									while (not rsEmailTemplates.EOF)
									%>
										<option value="<%=rsEmailTemplates.fields("TemplateId")%>"><%= rsEmailTemplates.fields("Name") %></option>
									<%
										rsEmailTemplates.movenext
									wend
								end if
								rsEmailTemplates.close
								set rsEmailTemplates = nothing
								%>
							</select>&nbsp;
							<img id="imgSendMail" name="imgSendMail"  src="images/envelop.gif" style="cursor: pointer;cursor: hand;" onclick="showDialog('cboEmailTemplates','hdnVikarIds');">
							<input name='hdnVikarIds' id='hdnVikarIds' TYPE='hidden'>
							</td>
							<%
						end if
						%>
						</tr>
						</table>
					</div>
					<div class="content">
						<div class='listing'>
							<%="<h3>" & StrAvdeling & "</h3>"%>
							<table cellspacing='1' cellpadding='0' width="96%" ID="Table13">
								<tr>
									<%
									if (blnFromJob = true) then
									%>
									<th>Tilknytt</th>
									<%
									else
									%>
									<th>E-post</th>
									<%
									end if
									%>
									<th>Ansattnr.</th>
									<th>Navn</th>
									<th>Hjemmeadresse</th>
									<th>Telefon</th>
									<th>Mobil</th>
									<th>Epost</th>
									<th nowrap>Vis CV</th>
								</tr>
								<%
								if(HasRows(rsSoek)) then
								
									nRecCount = rsSoek.RecordCount
									rsSoek.PageSize = RECORDS_PER_PAGE
									nPageCount = rsSoek.PageCount
									
									If nPage < 1 Or nPage > nPageCount Then
										nPage = 1			
									End If
									
									' Position recordset to the page we want to see
									rsSoek.AbsolutePage = nPage
									
									Do Until (rsSoek.EOF OR rsSoek.AbsolutePage <> nPage)
										strFullName   = rsSoek("Etternavn") & " " & rsSoek("Fornavn")
										strFullAdress = rsSoek("Adresse") & " " & rsSoek("Postnr") & " " & rsSoek("Poststed")
										%>
										<tr>
											<%
											if (blnFromJob	= true) then
												set rsAcceptedCommissions = GetFirehoseRS("exec [dbo].[spECGetAlreadyAcceptedCommissions] " & postedOppId & "," & rsSoek("VikarID"), objCon)
												if (not rsAcceptedCommissions.EOF) then
												%>
												<td><INPUT TYPE="CHECKBOX" class="checkbox" name='tbxID' value="<%=rsSoek("VikarID")%>" ID="Checkbox4" onclick=" MultipleAcceptance(this,true);" ></td>
												<%
												else
												%>
												<td><INPUT TYPE="CHECKBOX" class="checkbox" name='tbxID' value="<%=rsSoek("VikarID")%>" ID="Checkbox4" onclick="MultipleAcceptance(this,false);" ></td>
												<%
												end if
												rsAcceptedCommissions.close
												set rsAcceptedCommissions = nothing
											else
												%>
												<td><INPUT TYPE="CHECKBOX" class="checkbox" name='mailCHK' ID="<%= "mailCHK_" & rsSoek("VikarID") %>" onclick="<%= "addVikarId(this.id,'hdnVikarIds','" & rsSoek("VikarID") & "');"%>" ></td>
												<%
											end if
											if (rsSoek("ansattnummer") > 0) then
												%>
												<td class="left"><%=rsSoek("ansattnummer")%></td>
												<%
											else
												%>
												<td class="left">--</td>
												<%
											end if
											%>
											<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsSoek("vikarid"), strFullName, "Vis vikar " & strFullName )%></td>
											<td><%=strFullAdress%>&nbsp;</td>
											<td class='nowrap'><%=rsSoek("Telefon")%>&nbsp;</td>
											<% if(blnFromJob) then %>
											<td class='nowrap'><%=rsSoek("MobilTlf")%>&nbsp;
											<% 
												if(Not isNULL(rsSoek("MobilTlf")) and len(trim(rsSoek("MobilTlf"))) >0) then
											%>
												<a id="uxSMSLink" style="cursor: pointer; cursor: hand;" onclick="showModalDialog('WebUI/Sms/XisSms.aspx?VikarId=<% =rsSoek("vikarid") %>&CommissionId=<% =lngoppdragID %>&FirmaId=<% =lngFirmaID %>' ,null,'dialogWidth:490px;dialogHeight:360px;location:0;toolbar:0;resizable=0;status:0;menubar=0');"><img name="uxSMSImage" src="/xtra/images/smsbutton.gif" /></a>
											<%
											else
											%>
												<a id="uxSMSLink" disabled onclick=""><img name="uxSMSImage" src="/xtra/images/smsbuttondis.jpg" /></a>
											<% 
											end if
											%>
											</td>
											<% else %>											
											<td class='nowrap'><%=rsSoek("MobilTlf")%>&nbsp;
											<% 
												if(Not isNULL(rsSoek("MobilTlf")) and len(trim(rsSoek("MobilTlf"))) >0) then
											%>
												<a id="uxSMSLink" style="cursor: pointer; cursor: hand;" onclick="showModalDialog('WebUI/Sms/XisSms.aspx?VikarId=<% =rsSoek("vikarid") %>' ,null,'dialogWidth:490px;dialogHeight:360px;location:0;toolbar:0;resizable=0;status:0;menubar=0');"><img name="uxSMSImage" src="/xtra/images/smsbutton.gif" /></a>
											<%
											else
											%>
												<a id="uxSMSLink" disabled onclick=""><img name="uxSMSImage" src="/xtra/images/smsbuttondis.jpg" /></a>
											<% 
											end if
											%>
											</td>
											<% end if %>
																						
											<td class='nowrap'><A Href='mailto:<%=rsSoek("EPost")%>'><%=rsSoek("EPost")%></a></td>
											<td class='center'><A Href="javascript:VisCV('<%=rsSoek("VikarID")%>')"><img src="/xtra/images/icon_cv.gif" width="14" height="14" alt="" align="absmiddle"></a></td>
										</tr>
										<%
										rsSoek.MoveNext
									Loop
									rsSoek.Close
								end if
								%>
								
								</table>
								<%
							set rsSoek = Nothing
						%>
						</div>
						<div style="text-align:center">
						<table cellspacing='1' cellpadding='0'>
							<tr> 								
								<td><input name='hdnLastPage' id='hdnLastPage' TYPE='hidden' VALUE='<%= nPageCount %>'></td>
							</tr>
							<tr align="center">
								
								<td><a href="javascript:ViewPage('first');" style='visibility:<% If nPage = 1 Then Response.Write("hidden") Else  Response.Write("visible") %>'>&#171; Første</a></td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td><a href="javascript:ViewPage('pre');" style='visibility:<% If nPage = 1 Then Response.Write("hidden") Else  Response.Write("visible") %>'>&#8249; Forrige</a></td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td><% If ((nPage * RECORDS_PER_PAGE) < nRecCount) Then Response.Write(((nPage -1)* RECORDS_PER_PAGE  + 1) & " - " & nPage * RECORDS_PER_PAGE  & " av " & nRecCount) Else If (nRecCount > 0) Then Response.Write(((nPage -1)* RECORDS_PER_PAGE  + 1) & " - " & nRecCount  & " av " & nRecCount)End If End If%></td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td><a href="javascript:ViewPage('next');" style='visibility:<% If ((nPage = nPageCount) OR (nRecCount = 0)) Then Response.Write("hidden") Else  Response.Write("visible") %>'>Neste &#8250;</a></td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td><a href="javascript:ViewPage('last');" style='visibility:<% If ((nPage = nPageCount) OR (nRecCount = 0)) Then Response.Write("hidden") Else  Response.Write("visible") %>'>Siste &#187;</a></td>														
							</tr>
						</table>
						</div>
						<div>
						<table>
							<tr>
							<td>
						
					
					<%
					if (blnFromJob	= true) then
						%>
						<INPUT TYPE=SUBMIT NAME="btnOffer" VALUE="Legg til i shortlist " ID="Submit1">
						<INPUT TYPE="SUBMIT" NAME="btnAxcept" VALUE="Lagre med status Aksept " ID="Submit2">
						<%
					end if
								%>
							</td>
							</tr>
							<tr>
							<td>
							&nbsp;
							</td>
							</tr>
						
						</table>
						</div>
						
					</div>
					<%
				end if
				%>
			</div>
		</form>
</body>
</html>
<%
rsLevel.close
set rsLevel = nothing
rsProductRank.close
set rsProductRank = nothing
CloseConnection(objCon)
set objCon = nothing
%>
