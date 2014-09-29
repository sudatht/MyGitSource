<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\MailLib.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.Settings.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes\Xis.Security.Utils.inc"-->
<!--#INCLUDE FILE="includes\DNN.Users.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim lngVikarID
	dim strMottattSkattekort
	dim strRegDato
	dim strPresentasjon
	dim strBankKontonr
	dim	strOppsigelsestid
	dim	strBil
	dim	strFoererkort
	dim objFSO
	dim objSkjemaFolder
	dim objFilesCol
	dim NofFiles
	dim objFile
	dim FilePath
	dim strBgColor
	dim strType
	dim strName
	dim strSkjemaFileName
	dim strFileTypes
	dim profil
	dim objCon
	dim cons
	dim blnShowHotList
	dim strBruker
	dim rsBruker
	dim bDeletePersonInfo
	dim strResName
	
	dim strLink1URLShort
	dim strLink2URLShort
	dim strLink3URLShort

	dim isAccountLocked
	dim EnableUser

	isAccountLocked = false

	'Path to icons
	Dim strPicturePath

	'CV varibles
	dim CV
	dim hasCV
	dim isCVLocked

	hasCV = false

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType

	'Consultant menu variables
	dim strClass
	dim strJSEvents

	dim strRelevancyDate
	dim strRelevancyClass
	dim strRelevancyJSEvents
	dim strRelevancyDisabled
	dim nActivityTypeID
	dim strActivity

	brukerID = Session("BrukerID")

	' Move parameters to local variables
	lngVikarID = Request.Querystring("VikarID")

	' Check VikarID
	If lngVikarID = "" Then
		AddErrorMessage("Feil: Mangler parameter vikarID. Noter navn på vikar og kontakt systemansvarlig.")
		call RenderErrorMessage()
	End If

   ' Open database connection  
   
   
   Set objCon = GetConnection(GetConnectionstring(XIS, ""))

	' Get all consultant data
	If lngVikarID <> "" Then

		strVikar = "SELECT " & _
		"VIKAR.VikarID, " & _
		"VIKAR.StatusID, " & _
		"VIKAR.Etternavn, " & _
		"VIKAR.Fornavn, " & _
		"VIKAR.foedselsdato, " & _
		"VIKAR.personnummer, " & _
		"VIKAR.loenn1, " & _
		"VIKAR.StatusID, " & _
		"VIKAR.IntervjuDato, " & _
		"VIKAR.Telefon, " & _
		"VIKAR.MobilTlf, " & _
		"VIKAR.Fax, " & _
		"VIKAR.Epost, " & _
		"VIKAR.Hjemmeside, " & _
		"VIKAR.kontraktsendt, " & _
		"VIKAR.kontraktmottatt, " & _
		"VIKAR.notat, " & _
		"VIKAR.InterestedJobs, " & _
		"VIKAR.oppsummering_intervju, " & _
		"VIKAR.oppsummering_ref_sjekk, " & _
		"VIKAR_ANSATTNUMMER.ansattnummer, " & _
		"VIKAR.GodkjentAv, " & _
		"VIKAR.RegDato, " & _
		"VIKAR.GodkjentDato, " & _
		"VIKAR.MottattSkattekort, " & _
		"VIKAR.KundePresentasjon, " & _
		"VIKAR.bankkontonr, " & _
		"VIKAR.Link1URL, " & _
		"VIKAR.Link2URL, " & _
		"VIKAR.Link3URL, " & _
		"isnull(COUNTRY.PrintableName,'') AS PrintableName, " & _
		"MFornavn = MEDARBEIDER.Fornavn, " & _
		"MEtterNavn = MEDARBEIDER.Etternavn, " & _
		"ADRESSE.Adresse, " & _
		"ADRESSE.Postnr, " & _
		"ADRESSE.PostSted, " & _
		"H_VIKAR_TYPE.Vikartype, " & _
		"H_VIKAR_STATUS.VikarStatus, " & _
		"VIKAR.Foererkort, VIKAR.Bil, VIKAR.hasCar, VIKAR.Oppsigelsestid, VIKAR.WorkType " &_
		"FROM VIKAR " & _
		"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
		"LEFT OUTER JOIN H_VIKAR_TYPE ON VIKAR.TypeID = H_VIKAR_TYPE.VikarTypeID " & _
		"LEFT OUTER JOIN H_VIKAR_STATUS ON VIKAR.StatusID = H_VIKAR_STATUS.VikarStatusID " & _
		"LEFT OUTER JOIN MEDARBEIDER ON VIKAR.AnsMedID = MEDARBEIDER.MedID " & _
		"LEFT OUTER JOIN ADRESSE ON VIKAR.VikarID = ADRESSE.adresseRelID " & _
		"LEFT OUTER JOIN COUNTRY ON VIKAR.Country = Country.CountryID " & _
		"WHERE VIKAR.VikarID = '" & lngVikarID & "' " & _
		"AND ADRESSE.AdresseRelasjon = '2' " & _
		"AND ADRESSE.AdresseType = '1' "

		set rsVikar = GetFirehoseRS(strVikar, objCon)

		strVikarID				= rsVikar("VikarID").Value
		strAnsattnummer			= rsVikar("ansattnummer").Value
		strEtternavn			= rsVikar("Etternavn").Value
		strFornavn				= rsVikar("Fornavn").Value
		strFoedselsdato			= rsVikar("Foedselsdato").Value
		strRegDato				= rsVikar("RegDato").Value
		strPersonnummer			= rsVikar("Personnummer").Value
		strNotat				= rsVikar("Notat").Value
		strInterestedJobs       = rsVikar("InterestedJobs").Value
		strPresentasjon			= rsVikar("KundePresentasjon").Value
		lTimeloenn				= rsVikar("loenn1").Value
		strStatus				= rsVikar("VikarStatus").Value
		IStatusID				= rsVikar("StatusID").Value
		strVikarType			= rsVikar("Vikartype").Value
		strIntervjudato			= rsVikar("IntervjuDato").Value
		strTelefon				= rsVikar("Telefon").Value
		strFax					= rsVikar("Fax").Value
		strMobilTlf				= rsVikar("MobilTlf").Value
		strEPost				= rsVikar("EPost").Value
		strHjemmeside			= rsVikar("Hjemmeside").Value
		strGodkjentAv			= rsVikar("GodkjentAv").Value
		strGodkjentDato			= rsVikar("GodkjentDato").Value
		strAnsMedName			= rsVikar("MFornavn").Value & " " & rsVikar("MEtternavn").Value
		strKontraktSendt		= rsVikar("KontraktSendt").Value
		strKontraktmottatt		= rsVikar("Kontraktmottatt").Value
		strAdresse				= rsVikar("Adresse").Value
		strOppsumIntervju		= rsVikar("Oppsummering_intervju").Value
		strOppsumRefSjekk		= rsvikar("Oppsummering_ref_sjekk").Value
		strMottattSkattekort	= rsvikar("mottattskattekort").Value
		strFoererkort			= rsVikar("Foererkort").Value
		strBil					= rsVikar("hasCar").Value
		strOppsigelsestid		= rsVikar("Oppsigelsestid").Value
		strWorkType		= rsVikar("WorkType").Value
		strBankKontonr			= rsVikar("bankkontonr").Value
		strLink1URL				= rsVikar("Link1URL").Value
		strLink2URL				= rsVikar("Link2URL").Value
		strLink3URL				= rsVikar("Link3URL").Value	
		strCountry				= rsVikar("PrintableName").Value
		
								
		if (len(trim(strLink1URL))>=25) then
			strLink1URLShort =  Mid(strLink1URL, 1, 25) + ".."
		else 
		    strLink1URLShort = strLink1URL
		end if
		
		if (len(trim(strLink2URL))>25) then
			strLink2URLShort =  Mid(strLink2URL, 1, 25) + ".." 
		else
			strLink2URLShort = strLink2URL
		end if
		
		if (len(trim(strLink3URL))>25) then
			strLink3URLShort =  Mid(strLink3URL, 1, 25)  + ".."
		else
			strLink3URLShort = strLink3URL
		end if			
			
			
		' if the url consist of an http:// this should be removed here
		if (Mid((trim(strLink1URL)),1,7)="http://") then
			strLink1URL =  Mid(strLink1URL, 8, len(strLink1URL)) 		
		end if
		if (Mid((trim(strLink2URL)),1,7)="http://") then
			strLink2URL =  Mid(strLink2URL, 8, len(strLink2URL)) 		
		end if
		if (Mid((trim(strLink3URL)),1,7)="http://") then
			strLink3URL =  Mid(strLink3URL, 8, len(strLink3URL)) 		
		end if


        'Data Deletion 
        'PRO@EC               
        strBruker = "SELECT Navn, DeletePersonInfo FROM BRUKER WHERE ID= '" & brukerID & "'"
        set rsBruker = GetFirehoseRS(strBruker, objCon)
        
        bDeletePersonInfo       = rsBruker("DeletePersonInfo").Value
        strResName = rsBruker("Navn").Value
        rsBruker.Close
        set rsBruker = Nothing

		if (len(trim(strFoererkort))=0) then
			strFoererkort = ""
		elseif (strFoererkort=0) then
			strFoererkort = "Nei"
		elseif (strFoererkort=1) then
			strFoererkort = "Ja"
		end if

		if (len(trim(strBil))=0) then
			strBil = ""
		elseif (strBil = 0) then
			strBil = "Nei"
		elseif (strBil = 1) then
			strBil = "Ja"
		end if

		select case strOppsigelsestid
		case ""
			strOppsigelsestid = "selected"
		case "0"
			strOppsigelsestid = "Ingen"
		case "1"
			strOppsigelsestid = "14 dager"
		case "2"
			strOppsigelsestid = "1 m&aring;ned"
		case "3"
			strOppsigelsestid = "2 m&aring;neder"
		case "4"
			strOppsigelsestid = "3 m&aring;neder"
		case "5"
			strOppsigelsestid = "Over 3 m&aring;neder"
		end select

		select case strWorkType
		case "0"
			strWorkType = "Ikke angitt"
		case "1"
			strWorkType = "Kun fulltid"
		case "2"
			strWorkType = "Kun deltid"
		case "3"
			strWorkType = "Fulltid og deltid"
		end select

		' Create poststed
		strPoststed = rsVikar("Postnr").Value & " " & rsVikar("PostSted").Value
		strHeading = rsVikar("Fornavn").Value & " " &rsVikar("Etternavn").Value

		' Close and release recordset
		rsVikar.close
		Set rsVikar = Nothing

		set cons = Server.CreateObject("XtraWeb.Consultant")
		cons.XtraConString = Application("XtraWebConnection")
		cons.GetConsultant(lngVikarID)

		'Determine CV information
		set cv	= cons.CV
		cv.XtraConString = Application("Xtra_intern_ConnectionString")
		cv.XtraDataShapeConString = Application("ConXtraShape")
		cv.Refresh

		if cons.CV.DataValues.Count = 0 then
			cons.CV.Save
		else
			if (lcase(trim(Request("editCV"))) = "yes") then
				cv.unlockCV()
			end if
		end if

		hasCV = true
		isCVLocked = cv.islocked

		set cv = nothing
		cons.CV.cleanup
		cons.cleanup
		set cons = nothing
	End If

	'Sjekk for å forhindre dobbelreg i Hotlist
	set rsTest = GetFirehoseRS("Select * from HOTLIST Where status=3 And BrukerID=" & brukerID & " And navnID=" & strVikarID, objCon)
	if (HasRows(rsTest) = false) then
		blnShowHotList = true
		strAddToHotlistLink = "addHotlist.asp?kode=3&vikarNavn=" & server.URLEncode(strEtternavn & " " & strFornavn) & "&vikarNr=" & lngVikarID
		strHotlistType = "vikar"
	else
		blnShowHotList = true
		strAddToHotlistLink = ""
		strHotlistType = "vikar"
	end if

	' DNN user handling
	' - Resettes password if someone has pressed the 'Tilbakestill passord' button.
	'Variables used by dnn user routines.
	dim strUsername 'string
	dim strPassword 'Users new password
	dim strReset 	'string
	Dim iApp
	Dim sUserServiceURL
	Dim objUserProxy
	Dim objUserDom
	Dim sUserXml

	iApp = Cint(Application("Application"))
	sUserServiceURL = Application("DNNUserServiceURL")
	strReset = Request("Reset")

	'dersom vikar har status "ANSATT" or "Kandidat" or pt.optatt
       
'if ((cint(IStatusID) = 2) OR (cint(IStatusID) = 3) OR (cint(IStatusID) = 1) OR (cint(IStatusID) = 8))  then 
                
		Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
	    objUserProxy.Url = sUserServiceURL
	    
	    sUserXml = objUserProxy.GetUser(iApp, strVikarID,"V")
		if strReset="YES" then
			'Generate new password
			strPassword = GeneratePassword(6)
			'TODO implement password reset
			Set objUserDom = CreateUserDom(iApp, strVikarID, "", strPassword, "", "", "","V")
			Call objUserProxy.SaveUser(objUserDom.xml)
			Call SendResetPWD(strVikarID)
		end if
		set objUserProxy = nothing
		if sUserXml <> "" then
			Set objUserDom = Server.CreateObject("Microsoft.XmlDom")
			objUserDom.LoadXml sUserXml
			strUsername = objUserDom.selectSingleNode("/user/userName").Text
		end if
		set objUserDom = nothing
	'end if-->

'/IMP Brukerhåndtering

	'Sub that sends a mail informing substitute that his/her password has been reset.
	sub SendResetPWD(strVikarID)
		dim strSubject
		dim strBody
		dim strRemotehost
		dim strFromname
		dim strFromAdress
		dim strToAdress
		dim StrToName
		dim StrNavn
		dim emailText

		StrNavn = strFornavn & " " & strEtternavn

		emailText = getsetting("Mail_RESETPWD")
		emailText = replace(emailText, "%NAME%", strNavn)
		emailText = replace(emailText, "%MAILBODYRESETPWD%", Application("MailBodyresetPWD"))
		emailText = replace(emailText, "%PASSWORD%", strPassword)
		emailText = replace(emailText, "%MAILBODYSIGN%", Application("MailBodySign") )

		strSubject = Application("MailSubjectNewUID")
		strBody = emailText
		strRemotehost = Application("XtraMailServer")
		strFromname  = Application("XtraSenderName")
		strFromAdress = Application("XtraSenderMail")
		strToAdress = strEPost
		StrToName = strFornavn & " " & strEtternavn

		call sendMail(strFromAdress, StrFromName, strSubject, strBody, strRemoteHost, strToAdress, StrToName)
	end sub
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language="javascript" src="js/menu.js" id="MenuScript" ></script>
		<title><%=strHeading %></title>
	</head>
	<script language='javascript' src='js/menu.js' id='menuScripts'></script>
	<script language='javascript' src='js/navigation.js' id='navigationScripts'></script>
	<script language="javaScript" type="text/javascript">
		function shortKey(e)
		{
			var keyChar = String.fromCharCode(event.keyCode);
			var modKey  = event.ctrlKey;
			var modKey2 = event.shiftKey;

			//linker i submeny
			<%
			If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
				%>
				if (modKey && modKey2 && keyChar=="S")
				{
					parent.frames[funcFrameIndex].location=("/xtra/vikarSoek.asp");
				}
				<%
			End If
			If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
				%>
				if (modKey && modKey2 && keyChar=="Y")
				{
					parent.frames[funcFrameIndex].location=("/xtra/VikarDuplikatSoek.asp");
				}
				<%
			End If
			If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
				%>
				if (modKey && modKey2 && keyChar=="W")
				{
					parent.frames[funcFrameIndex].location=("/xtra/jobb/SuspectList.asp");
				}
				<%
			End If
			%>
		}
		//her catches eventen som trigger shortcut'en
		document.onkeydown = shortKey;

		function ResetPWD()
		{
			location = ("Vikarvis.asp?Reset=YES&VikarID=<%=strVikarID%>");
		}

/*
		function GrantCVAccess()
		{
			location = ("Vikarvis.asp?editCV=yes&VikarID=<%=strVikarID%>");
			
			 Kommentert ut ettersom denne boksen aldri kom opp riktig kjørt i en pane i SuperOffice
			
			var response = confirm("Du har valgt å åpne CVen til vikar for redigering.\nDersom du lagrer nå vil HTML formatering komme til å forsvinne\nnår brukeren redigerer CVen på web.\nFor å åpne CV''en trykk på 'OK', trykk på 'Avbryt' for å avbryte.");
			if(response == true)
			{
				location = ("Vikarvis.asp?editCV=yes&VikarID=<%=strVikarID%>");
				return;
			}
			else if(response == false)
			{
				return;
			}
			
		}
*/
		function Upload()
		{
			window.open("fileUpload.asp?vikarid=<%=strVikarID%>","Oversikt","height=150, width=400, status=yes, toolbar=no, menubar=no, location=no")
		}

		function ShowConsultantPicture()
		{
			var oWnd
			oWnd = window.open("ShowConsultantPicture.asp?vikarid=<%=strVikarID%>", null, "height=560, width=680, scrollbars=yes, status=yes, resizable=yes, toolbar=no, menubar=no, location=no")
			oWnd.focus();
		}
	</script>
	<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
	<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%" id="Table1">
						<tr>
							<td>
								<%
								strClass = "menu disabled"
								strJSEvents = ""
								
								strRelevancyClass = "menu"
								strRelevancyJSEvents = "onMouseOver=""menuOver(this.id);"" onMouseOut=""menuOut(this.id);"""
								strRelevancyDisabled = ""
								
								strActivity = "Oppfriskningsdato"
									
								set rsActivityType = GetFirehoseRS("SELECT AktivitetTypeID FROM H_AKTIVITET_TYPE WHERE AktivitetType = '" & strActivity & "'", objCon)
								nActivityTypeID = CInt(rsActivityType("AktivitetTypeID"))
								' Close and release recordset
						      	rsActivityType.Close
						      	Set rsActivityType = Nothing
						      	'response.write "zzz"
'response.write nActivityTypeID
					      		set rsRelevancyDate = GetFirehoseRS("Select  top 1 AktivitetDato from Aktivitet where vikarid =" & lngVikarID & "  and aktivitettypeid = " & nActivityTypeID & " order by aktivitetdato desc", objCon)
					      		if not (rsRelevancyDate.EOF) then
					      			if datediff("d",datevalue(rsRelevancyDate("AktivitetDato")),datevalue(now())) <= 30 then
								      		strRelevancyClass = "menu disabled"
											strRelevancyJSEvents = ""
											strRelevancyDisabled = "disabled"
									end if			
					      			strRelevancyDate = "Sist oppfrisket: " & rsRelevancyDate("AktivitetDato")
					      		
					      		else
					      			strRelevancyDate = "Ingen oppfriskningsdato registrert"
					      		end if
					      		
					      		' Close and release recordset
						      	rsRelevancyDate.Close
						      	Set rsActivityType = Nothing 
						      	
								%>
								<table cellpadding="0" cellspacing="2" id="Table2">
									<tr>
										<td class="<%=strClass%>" id="menu1" <%=strJSEvents%>>
											<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
										</td>
										<td class="<%=strClass%>" id="menu2" <%=strJSEvents%>>
											<img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle">Vis
										</td>
										<%
										If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
											%>
											<td class="menu" id="menu3" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);">
												<form action="vikarny.asp?VikarID=<%=strVikarID %>" method="POST" id="frmConsultantChange">
													<input name="pbnDataAction" type="hidden" value="Endre kons.opplysninger" id="Hidden1">
													<a href="javascript:document.all.frmConsultantChange.submit();" title="Endre vikar">
													<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre</a>
												</form>
											</td>
											<%
										End If
										If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
											%>
											<td class="menu" id="menu4" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);"><strong>CV&nbsp;</strong><select id="cboCVChoice" onchange="javascript:Vis_CV(<%=lngVikarID%>);" name="cboCVChoice"><option value="0"></option><option value="1">Se</option><%If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then%><option value="2">Endre</option><%end if%><option value="3">Presentere</option></select></td>
										 
											<%
										End If
										%>
										<td class="menu" id="menu6" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);">
											<form action="vikar-kunder.asp?VikarID=<%=strVikarID%>" method="POST" id="frmConsultantFormerClients">
												<a href="javascript:document.all.frmConsultantFormerClients.submit();" title="Vis tidligere oppdragsgivere"><img src="/xtra/images/icon_tidl-kunder.gif" alt="" width="18" height="15" border="0" align="absmiddle">Tidligere oppdragsgivere</a>
											</form>
										</td>
										<td class="menu" id="menu7" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);">
											<form action="AktivitetVikar.asp?VikarID=<%=strVikarID%>" method="POST" id="frmConsultantActivities">
												<a href="javascript:document.all.frmConsultantActivities.submit();" title="Vis aktiviteter for vikaren"><img src="/xtra/images/icon_activities.gif" alt="" width="18" height="15" border="0" align="absmiddle">Aktiviteter</a>
											</form>
										</td>
										<%
										If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
										%>
										<td class="<%=strRelevancyClass%>" id="menu9" <%=strRelevancyJSEvents%>>
											<form ACTION="RelevancyActivity.asp?VikarID=<%=lngVikarID%>" METHOD="POST" id="frmRelevanyRegister">
												<%
												 	if strRelevancyDisabled = "" then
												 	%>
												 		<a href="javascript:document.all.frmRelevanyRegister.submit();" title="<% =strRelevancyDate %>"><img src="/xtra/images/exclamation.png" alt="" border="0" align="absmiddle">Oppfrisk!</a>												
												 	<%
												 	else
												 	%>
												 		<a disabled title="<% =strRelevancyDate %>"><img src="/xtra/images/exclamation.png" alt="" border="0" align="absmiddle">Oppfrisk!</a>
												 	<%
												 	end if	
												%>
											</form>
										</td>
										<%
										End If										
										'Data Deletion : PRO@EC 
										If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) and bDeletePersonInfo Then
											%>
											<td class="menu" id="menu8" onmouseover="menuOver(this.id);" onmouseout="menuOut(this.id);">											
												<form action="/xtra/WebUI/DataDeletion/ShowResult.aspx?VikarID=<%=strVikarID %>" method="POST" id="frmDeletePersonInfo" >
													<input name="DeletePageID" type="hidden" value="OnDemandDelete" id="DeletePageID" />
													<input name="VikarName" type="hidden" value="<% Response.write(strFornavn & " " & strEtternavn) %>" id="VikarName" />
													<input name="Fornavn" type="hidden" value="<% Response.write(strFornavn) %>" id="Fornavn" />
													<input name="Etternavn" type="hidden" value="<% Response.write(strEtternavn) %>" id="Etternavn" />
													<input name="EPost" type="hidden" value="<%=strEPost%>" id="EPost" />
													<input name="ResName" type="hidden" value="<%=strResName%>" id="ResName" />
													<input name="RegDate" type="hidden" value="<%=strRegDato%>" id="RegDate" />
													<input name="IntvDate" type="hidden" value="<%=strIntervjudato%>" id="IntvDate" />													
													
													<a href="javascript:document.all.frmDeletePersonInfo.submit();" onclick="return confirm('Er du helt sikker på at du vil slette all informasjon for denne vikaren?');" title="Slett vikarinfo">
													<img src="/xtra/images/icon_delete.gif" width="14" height="14" alt="" align="absmiddle" />Slett</a>
												</form>
											</td>
											<% 
										End If
										'/Data Deletion
										%>
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
			<table class="layout" cellpadding="0" cellspacing="0" id="Table3">
				<col width="33%">
				<col width="34%">
				<col width="33%">
				<tr>
					<td>
						<table id="Table4">
							<tr>
								<th>Vikarid:</td>
								<td><%=strVikarID%></td>
							</tr>
							<tr>
								<th>Ansattnummer:</th>
								<td>
									<%
									if (strAnsattnummer <> "") then
										Response.Write strAnsattnummer
									else
										Response.write "---"
									end if
									%>
								</td>
							</tr>
							<tr>
								<th>Opprettet:</th>
								<td><%=strRegDato%>&nbsp;</td>
							</tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr>
								<th>Fornavn:</th>
								<td><%=strFornavn %>
								<%
								strPicturePath = Application("ConsultantImages") & strVikarID & ".jpg"
								Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

								If objFSO.FileExists(strPicturePath) then
									%>
									<img alt="Vis foto av <%=strFornavn%>" style="cursor:pointer;" onclick="javascript:ShowConsultantPicture()" src="images/icon_consultantPortrait.gif" width="15" height="15">
									<%
								end if
								set objFSO = nothing
								%>

								</td>
							</tr>
							<tr>
								<th>Etternavn:</th>
								<td><%=strEtternavn %>&nbsp;</td>
							</tr>
							<tr>
								<th>Fødselsdato:</th>
								<td><%=strFoedselsdato %>&nbsp;</td>
							</tr>
							<tr>
								<th>Nasjonalitet:</th>
								<td><%=strCountry %>&nbsp;</td>
							</tr>
						</table>
					</td>
					<td>
	<!-- second -->
						<table id="Table5">
							<tr>
								<th>Telefon:</th>
								<td><%=strTelefon%>&nbsp;</td>
							</tr>
							<tr>
								<th>Mobil:</th>
								<td><%=strMobilTlf%>&nbsp;								
								<% 
									if(Not isNULL(strMobilTlf) and len(trim(strMobilTlf)) >0) then
								%>
									<a id="uxSMSLink" onclick="showModalDialog('WebUI/Sms/XisSms.aspx?VikarId=<% =lngVikarID %>' ,null,'dialogWidth:490px;dialogHeight:360px;location:0;toolbar:0;resizable=0;status:0;menubar=0');"><img name="uxSMSImage" src="/xtra/images/smsbutton.gif" /></a>								
								<%
								else
								%>
									<a id="uxSMSLink" disabled onclick=""><img name="uxSMSImage" src="/xtra/images/smsbuttondis.jpg" /></a>								
								<% 
								end if
								%>
								</td>
							</tr>
							<tr>
								<th>Fax:</th>
								<td><%=strFax%>&nbsp;</td>
							</tr>
							<tr>
								<th>E-Post:</th>
								<td><a href="mailto:<%=strEPost%>"><%=strEPost%>&nbsp;</td>
							</tr>
							<!--
							<tr>
								<th>Hjemmeside:</th>
								<%if (len(trim(strHjemmeside))>0) then%>
								<td><a href="http://<%=strHjemmeside%>" target="_NEW"><%=strHjemmeside%></td>
								<%else%>
								<td>&nbsp;</td>
								<%end if%>
							</tr
							-->
							<tr>
								<th>Har f&oslash;rerkort:</th>
								<td><%=strFoererkort%>&nbsp;</td>
							</tr>
							<tr>
								<th>Disponerer bil:</th>
								<td><%=strBil%>&nbsp;</td>
							</tr>
							<tr>
								<th>Oppsigelsestid:</th>
								<td><%=strOppsigelsestid%>&nbsp;</td>
							</tr>
							<tr>
								<th>Stillingsbrøk:</th>
								<td><%=strWorkType%>&nbsp;</td>
							</tr>
						</table>
					</td>
					<td>
					<table >
					<tr>
					<th>
				
						<strong>Web brukernavn:</strong>&nbsp;
					
					</th>
					<td>
					<%=strUserName%>
					</td>
					</tr>
					</table>
	<!-- third -->
						<table id="Table6">
							<tr>
								<th>Hjemmeadresse:</th>
								<td><%=strAdresse%>&nbsp;</td>
							</tr>
							<tr>
								<td>Poststed:</td>
								<td><%=strPoststed %>&nbsp;</td>
							</tr>
						</table>
						<div class="insideElement">
							<form action="adresse.asp?Relasjon=2&amp;ID=<%=strVikarId%>" method="POST" id="frmConsultantAdress" name="frmConsultantAdress">
								<table id="Table7">
									<%
									set rsAdress = GetFirehoseRS("Select A.AdrId, A.Adresse , T.AdresseType, A.Postnr, A.poststed from ADRESSE A, H_ADRESSE_TYPE T where A.AdresseRelasjon = 2 and A.adresseRelID = " & lngVikarID & "and A.AdresseType = T.AdrTypeID and A.adressetype > 1 ", objCon)
									Do Until rsAdress.EOF
										strFullAdress = rsAdress("Adresse") & " " & rsAdress("PostNr") & " " & rsAdress("PostSted")
										%>
										<tr>
											<th><%=rsAdress("AdresseType") %>:</th>
											<td>
												<%
												If  HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
													%>
													<a href="adresse.asp?Relasjon=2&amp;ID=<%=lngVikarID %>&amp;AdrID=<%=rsAdress("AdrId") %>"><%= strFullAdress %> </a>
													<%
												Else
													Response.Write strFullAdress
												End If
												%>
											</td>
										</tr>
										<%
										rsAdress.MoveNext
									Loop
									' Close and release recordset
									rsAdress.Close
									Set rsAdress = Nothing
									%>
								</table>
								<%
								If  HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
									%>
									<span class="menuInside" title="Opprette ny adresse"><a href="#" onclick="javascript:document.all.frmConsultantAdress.submit();"><img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Ny adresse</a></span>
									<%
								End If
								%>
							</form>
						</div>
						<table >
							<tr>
								<th>Link 1:</th>
								<td title="<%=strLink1URL%>">
								  <a href="http:\\<%=strLink1URL%>" target="_blank"><%=strLink1URLShort%></a>&nbsp;
								</td>
								
							</tr>
							
							<tr>
								<th>Link 2:</th>
								<td title="<%=strLink2URL%>">
								  <a href="http:\\<%=strLink2URL%>" target="_blank"><%=strLink2URLShort%></a>&nbsp;
								</td>
							</tr>
							
							<tr>
								<th>Link 3:</th>
								<td title="<%=strLink3URL%>">
								  <a href="http:\\<%=strLink3URL%>" target="_blank"><%=strLink3URLShort%></a>&nbsp;
								</td>
							</tr>							
						</table>
						
					</td>
				</tr>
			</table>
		</div>

		<div class="contentHead"><h2>Notater</h2></div>
		<div class="content">
			<table class="layout" cellpadding="10" cellspacing="0" id="Table8">
				<col width="50%">
				<col width="50%">
				<tr>
					
					<td>
						<h3>Notat:</h3>
						<p><%=strNotat%></p>
						<h3>Stillings-interesse:</h3>
						<p><%=strInterestedJobs%></p>
						<h3>Kandidatpresentasjon:</h3>
						<p><%=strPresentasjon%></p>
					</td>
					<td>
						<h3>Etter intervju:</h3>
						<p><%=strOppsumIntervju%></p>

						<h3>Etter referansesjekk:</h3>
						<p><%=strOppsumRefSjekk%></p>
					</td>
					
					
				</tr>
			</table>
		</div>
		<div class="contentHead">
			<h2>Ansattinformasjon, avdelingskontor og tjenesteområder</h2>
		</div>
		<div class="content">
			<table class="layout" cellpadding="0" cellspacing="0" id="Table9">
				<col width="33%">
				<col width="34%">
				<col width="33%">
				<tr>
					<td>
	<!-- first -->
						<table id="Table10">
							<tr>
								<th>Intervjdato:</th>
								<td> <%=strIntervjudato%>&nbsp;</td>
							</tr>
							<tr>
								<th>Ansvarlig:</th>
								<td> <%=strAnsMedName%>&nbsp;</td>
							</tr>
							<tr>
								<th>Status:</th>
								<td> <%=strStatus%>&nbsp;</td>
							</tr>
							<tr>
								<th>Type:</th>
								<td><%=strVikartype%>&nbsp;</td>
							</tr>
							<tr>
								<th>Ønsket timelønn:</th>
								<td><%=lTimeloenn%>&nbsp;</td>
							</tr>
							<tr>
								<th>Kontonummer:</th>
								<td><%=strBankKontonr%>&nbsp;</td>
							</tr>

							<tr><td colspan="2">&nbsp;</td></tr>
							<tr>
								<th>Kontr. Sendt:</th>
								<td><%=strkontraktsendt%>&nbsp;</td>
							</tr>
							<tr>
								<th>Kontr. Mottatt:</th>
								<td><%=strKontraktmottatt%>&nbsp;</td>
							</tr>
						</table>

					</td>
					<td>
	<!-- second -->
						<table id="Table11">
							<tr>
								<th>Godkjent av:</th>
								<td><%=strGodkjentav %>&nbsp;</td>
							</tr>
							<tr>
								<th>Godkjent dato:</th>
								<td><%=strGodkjentdato%>&nbsp;</td>
							</tr>
							<tr>
								<th>Skattekort:</th>
								<%If (strMottattSkattekort = -1) Then%>
								<td>Nei&nbsp;</td>
								<%Else%>
								<td><%=strMottattSkattekort%>&nbsp;</td>
								<%End If%>
							</tr>
							<%
								dim strAvdelingskontorer
								' Get all connected AVDELING
								strSql = "select distinct t2.id, t2.navn, t1.vikarid " &_
									"from vikar_arbeidssted t1, avdelingskontor t2 " &_
									"where t1.AvdelingskontorID = t2.id " &_
									"and vikarid = " & strVikarID

								set rsVikarAvdeling = GetFirehoseRS(strSql, objCon)
								' Print lead text
								Response.Write "<tr><th>" & "Avdelingskontorer: " & "</th><td>"
								strAvdelingskontorer = ""
								' loop on result and display in table
								Do while not rsVikarAvdeling.EOF
							 		strAvdelingskontorer = strAvdelingskontorer & rsVikarAvdeling("navn" ) & ", "
							 		rsVikarAvdeling.MoveNext
								Loop
								if len(strAvdelingskontorer)>0 then
									strAvdelingskontorer = Mid(strAvdelingskontorer, 1, len(strAvdelingskontorer)-2)
								end if
								Response.Write strAvdelingskontorer & "&nbsp;</td></tr>"
								' Close and release recordset
								rsVikarAvdeling.close
								Set rsVikarAvdeling = Nothing
							%>

						</table>
					</td>
					<td>
	<!-- third -->
						<table id="Table12">
						<%
							dim strTjenesteomrader 'as string
							dim rsTjenesteomrader 'as adodb.recordset

							' Get all connected Tjenesteomrader
							strSql = "select distinct t2.tomid, t2.navn, t1.vikarid "  &_
						 		"from vikar_tjenesteomrade t1, tjenesteomrade t2 "  &_
						 		"where t1.tomID = t2.tomid "  &_
						 		"and vikarid = " & strVikarID

							set rsTjenesteomrader = GetFirehoseRS(strSql, objCon)
							' Print lead text
							Response.Write "<tr><th>" & "Tjenesteområder: " & "</th><td>"
							strTjenesteomrader = ""
							' loop on result and display in table
							with rsTjenesteomrader
								Do while not .EOF
									strTjenesteomrader = strTjenesteomrader & .fields("navn") & ", "
									.MoveNext
								Loop
								if len(strTjenesteomrader)>0 then
									strTjenesteomrader = Mid(strTjenesteomrader, 1, len(strTjenesteomrader)-2)
								end if
								Response.Write strTjenesteomrader & "&nbsp;</td></tr>"
							end with
							rsTjenesteomrader.close
							Set rsTjenesteomrader = Nothing
							
							dim strCategory
							dim rsCategory
							
							strSql = "select t2.categoryid, t2.name, t1.vikarid " &_
									"from vikar_category t1 inner join oppdrag_category t2 on t1.categoryid = t2.categoryid " &_
									"where  t1.vikarid = " & strVikarID
							
							set rsCategory = GetFirehoseRS(strSql, objCon)
							' Print lead text
							Response.Write "<tr><th>" & "Kategori: " & "</th><td>"
							strCategory = ""
							' loop on result and display in table
							with rsCategory
								Do while not .EOF
									strCategory = strCategory & .fields("name") & ", "
									.MoveNext
								Loop
								if len(strCategory)>0 then
									strCategory = Mid(strCategory, 1, len(strCategory)-2)
								end if
								Response.Write strCategory & "&nbsp;</td></tr>"
							end with
							rsCategory.close
							Set rsCategory = Nothing
							
							
						%>
						</table>
					</td>
				</tr>
			</table>
			</div>
			<div class="contentHead">
				<h2>Dokumentarkiv</h2>
			</div>
			<div class="content">
				<table width="70%" id="Table13">
				<tr>
				<td width="70%">
				<%
				dim util
				set  util = Server.CreateObject("XisSystem.Util")

				'Get location of all files
				FilePath = Application("ConsultantFileRoot") & strVikarID & "\"

				if (util.EnsurePathExists(FilePath) = false) then
					'Response.Clear
					dim str
					str = "Feil under aksessering av nettverksressurs (" & FilePath & ")"
					AddErrorMessage(str)
					'AddErrorMessage("Feil under aksessering av nettverksressurs.")
					call RenderErrorMessage()
				end if

				call util.Logon()

				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

				'Valid files with registered icons
				strFileTypes = "|pdf|xls|doc|txt|msg|"
				set objSkjemaFolder = objFSO.getfolder(FilePath)
				set objFilesCol = objSkjemaFolder.Files
				NofFiles = objFilesCol.count
				%>
				<div class="fileContainer">
				<%
					if (NofFiles > 0) then
				%>
				<div class="listing">
					<table width="98%" cellpadding="0" cellspacing="1" id="Table14">
						<tr>
							<th>&nbsp;</th>
							<th>Fil</th>
							<th>Endret</th>
						</tr>
						<%
						for each objFile in objFilesCol
							strName = objfile.name
							strSkjemaFileName = strName
							strDotPos = instr(strName,".")
							if (strDotPos > 0) then
								strType = mid(strName, strDotPos+1)
								strName = mid(strName,1, strDotPos-1)
							end if
							if ((NofRecords MOD 2) = 0) then
								strBgColor = "#ffffff"
							else
								strBgColor = "#ffffff"
							end if
							if (instr(strFileTypes,"|" & strType & "|") > 0) then
								response.write "<tr style='background-color:" & strBgColor & "'><td><img title="""" hspace=""3"" src=""/xtra/images/icon_" & strType & ".gif""></td>"
							else
								response.write "<tr style='background-color:" & strBgColor & "'><td><img hspace=""3"" src=""/xtra/images/icon_document.gif""></td>"
							end if
							response.write "<td><a href='" & FilePath & strSkjemaFileName & "' target='new'>" & strName & "<a>&nbsp;</td>"
							response.write "<td>" & objfile.DateLastModified & "</td>"

							NofRecords = NofRecords  + 1
						next
						set objFSO = nothing
						set objSkjemaFolder = nothing
						set objFile = nothing
						%>
						</table>
						<%
					end if
					%>
					</div>
					</td>
						<td width="30%">
							<div class="menuInside" title="Last opp ny fil" onmouseover="this.style.cursor = 'hand';" onclick='javascript:Upload()'><img src="/xtra/images/icon_upload.gif" width="18" height="15" alt="Last opp ny fil" border="0" align="absmiddle">Last opp fil</div>
							<div class="menuInside" title="Jeg vil se mappa hans!"><a href="<%=FilePath%>" target="_blank"><img src="/xtra/images/icon_open2.gif" width="18" height="15" alt="Jeg vil se mappa hans!" border="0" align="absmiddle">&Aring;pne mappe</a></div>
							<div class="menuInside" title="Oppdater innholdet i fil-listen" onmouseover="this.style.cursor = 'hand';" onclick='javascript:window.location.reload()'><img src="/xtra/images/icon_refresh.gif" width="18" height="15" alt="Oppdater innholdet i fil-listen" border="0" align="absmiddle">Oppdater</div>
						</td>
					</tr>
				</table>
				</div>
				<%
				call util.Logoff()
				set util = nothing
				%>
			<div class="contentHead" id="competance"><h2>Fag- og produktkompetanse</h2></div>
			<div class="content">
				<%
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
				%>
				<table class="layout" cellpadding="10" cellspacing="0" id="Table15">
					<col width="50%">
					<col width="50%">
					<tr>
						<td>
							<%
							End If
							set rsKompetanse = GetFirehoseRS("exec GetProfessionsForConsultant " & lngVikarID, objCon)
							if (not rskompetanse.EOF) then
							%>
							<div class="listing">
								<table cellpadding="0" cellspacing="0" id="Table16">
									<tr>
										<th>Fagkompetanse</th>
										<th>Erfaring</th>
										<th>Utdannelse</th>
										<th>Kommentar</th>
										<th>Slette</th>
									</tr>
									<%
									While (not rskompetanse.EOF)
										%>
										<tr>
											<%
											If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
											%>
												<td>
													<a name="AKompModify" href="kompetanse.asp?VikarID=<%=strVikarId%>&TypeID=4&KompetanseID=<%=rskompetanse.fields("KompetanseID")%>"><%=rskompetanse.fields("KTittel")%></a>
												</td>
											<%
											else
											%>
												<td><%=rskompetanse.fields("KTittel")%></td>
											<%
											end if
											%>
											<td><%=rskompetanse.fields("K_Erfaring").value%>&nbsp;</td>
											<td class="right"><%=rskompetanse.fields("k_nivaa").value%>&nbsp;</td>
											<td><%=rskompetanse.fields("kommentar").value%>&nbsp;</td>
											<td class="center"><a href="kompetanseDB.asp?tbxVikarID=<%=strVikarID%>&tbxAction=slette&tbxKompetanseID=<%=rskompetanse.fields("kompetanseID").value%>"><img src="/xtra/images/icon_delete.gif" alt="Slette" width="14" height="14" border="0"></a></td>
										</tr>
							   			<%
							   			rskompetanse.MoveNext
									wend
								' Close and release recordset
								rsKompetanse.close
								Set rsKompetanse = Nothing
								%>
								</table>
							</div>
							<%
							end if
							If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
								%>
								<form action="vikarJobbOnskeliste.asp?VikarID=<% =strVikarID %>" method="POST" id="frmConsultantNewCompetence">
						   			<input name="tbxVikarID" type="HIDDEN" value="<%=strVikarID %>" id="Hidden2">
									<div class="right"><span class="menuInside" title="Legg til Ny fagkompetanse"><a href="javascript:document.all.frmConsultantNewCompetence.submit();"><img src="/xtra/images/icon_New2.gif" width="14" height="14" alt="" align="absmiddle"> Ny fagkompetanse</a></span></div>
								</form>
								<%
							end if
							%>
							</td>
							<td>
							<%
							set rsKompetanse = GetFirehoseRS("exec GetQualificationsForConsultant  " & lngVikarID , objCon)
							if (NOT rskompetanse.EOF) then
								%>
								<div class="listing">
								<table cellpadding="0" cellspacing="0" id="Table17">
									<tr>
										<th>Produkt</th>
										<th>Kursniv&aring;</th>
										<th>Bruker niv&aring;</th>
										<th>Kommentar</th>
										<th>Slette</th>
									</tr>
									<%
									while (not rskompetanse.EOF)
										If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
										%>
										<tr>
											<td><a name="AKompModify" href="kompetanse.asp?VikarID=<%=strVikarId%>&TypeID=3&KompetanseID=<%=rskompetanse.fields("KompetanseID")%>"><%=rskompetanse.fields("KTittel")%></a></td>
										<%
										else
										%>
										<tr>
											<td><%=rskompetanse.fields("KTittel")%></td>
										<%
										end if
										%>
											<td><%=rskompetanse.fields("KLevel").value%>&nbsp;</td>
											<td class="right"><%=rskompetanse.fields("k_Rangering").value%>&nbsp;</td>
											<td><%=rskompetanse.fields("kommentar").value%>&nbsp;</td>
											<td class="center"><a href="kompetanseDB.asp?tbxVikarID=<%=strVikarID%>&tbxAction=slette&tbxKompetanseID=<%=rskompetanse.fields("kompetanseID").value%>"><img src="/xtra/images/icon_delete.gif" alt="Slette" width="14" height="14" border="0"></a></td>
										</tr>
						   				<%
						   				rskompetanse.MoveNext
									wend
									' Close and release recordset
									rsKompetanse.close
									Set rsKompetanse = Nothing
									%>
								</table>
								</div>
								<%
							end if
							If HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) Then
								%>
								<form action="vikarKompliste.asp?VikarID=<% =strVikarID %>" method="POST" id="frmConsultantNewProdCompetence">
									<input name="tbxVikarID" type="HIDDEN" value="<%=strVikarID %>" id="Hidden3">
									<div class="right"><span class="menuInside" title="Legg til produktkompetanse"><a href="javascript:document.all.frmConsultantNewProdCompetence.submit();"><img src="/xtra/images/icon_New2.gif" alt="" width="14" height="14" border="0" align="absmiddle"> Ny produktkompetanse</a></span></div>
								</form>
								<%
							end if
							%>
						</td>
					</tr>
				</table>
			</div>
		</div>
		<br/>
		<br/>
	</body>
</html>

<%
CloseConnection(objCon)
set objCon = nothing
%>
