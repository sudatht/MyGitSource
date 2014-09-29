<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	'Page parameters
	dim lngOppdragID
	dim lngFirmaID : lngFirmaID = 0
	dim SOcuID
	dim NoOfTimesheets
	dim lngBrukerID
        dim linkUser
        dim linkPassword
        dim linkDomain
        dim strKunde
       
	dim strSQL
	dim rsOppdrag
	dim rsHotlist
	dim rsContact
	dim lAvdelingID
	dim itomID
	dim icategoryID
	dim lRegnskapID
	dim contactDisabled

	dim rsAvdeling
	dim rsRegnskapskontor
	dim rsDefaultTOM
	dim rsTjenesteOmrade
	dim rsCategory
	dim rsDefaultRA
	dim rsVikarStatus
	dim sel
	dim strKontaktperson
	dim	strBestiltDato
	dim	strBestiltKl
	dim	strCustomerReference
	dim	strBeskrivelse
	dim	strFrakl
	dim	strTilkl
	dim	tmpDate
	dim	Timerprdag
	dim	Lunsj
	dim	strOppdrag
	dim strRecruitment
	dim isDisabledDirectRec
	dim isDisabledDirectNoCopy
	dim isDisableOppdrag
	dim	lAnsMedID
	dim	lStatusID
	dim	lFaId 
	dim bChooseSelected
	dim lBestilltAv
	dim timeloenn
	dim webPub
	dim bReportingOverrule
	dim ReportingContactID
	dim webPubEnabled
    dim terms
	Dim cts 
	Dim personsHTML
    dim comments
    dim reruitmentDate
    dim amount
    dim noofpersons
    dim noofopportunities
	dim parentAssignment
    dim hourlyRate
	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList
	
	dim strVikarFirstName
	dim strVikarLastName
	dim hasAcceptedTemp
	dim ansattNummer
	dim acceptedVikar
	dim OppdragStatusEnable 
	'copy commission PRO@EC
	dim ccAction
	dim vikarID	
	dim ovid
	ccAction = false
	
	// Text Editor Code - FckEditor	
	'Dim editor	
	'Set editor = New FCKeditor
	'editor.BasePath = "fckeditor/"
	'editor.Width = "1000"
	'editor.Height = "130"
	'editor.ToolBarSet = "Basic"   ' Default or Basic
	'editor.Value = ""	
	
	strVikarFirstName = ""
	strVikarLastName = ""
	acceptedVikar = 0
	ansattNummer = 0
	hasAcceptedTemp = false
	bChooseSelected = true

	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = ""

	lngBrukerID = Session("medarbID")

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	' Move parameter OppdragsID to variable
	If len(trim(Request("OppdragID"))) > 0 Then
	   lngOppdragID = Request.QueryString("OppdragID")
	Else
	   lngOppdragID = 0
	End If

	' Move parameter FirmaID to variable
	If len(trim(Request("FirmaID")))> 0 Then
	   lngFirmaID = Request("FirmaID")

		strSQL = "SELECT SOCUID FROM Firma WHERE FirmaId = " & lngFirmaID
		set rsContact = GetFirehoseRS(strSQL, conn)
		if(HasRows(rsContact) = true) then
			SOcuID = rsContact("SOCUID")
		end if
		rsContact.close
		set rsContact = nothing
	Else
		if (len(request("cuid")) > 0) then
			strSQL = "SELECT FirmaId FROM Firma WHERE SOCUID = " & request("cuid")
			set rsContact = GetFirehoseRS(strSQL, conn)
			if(HasRows(rsContact) = true) then
				lngFirmaID = rsContact("FirmaId")
			end if
			rsContact.close
			set rsContact = nothing
		end if
	End If

	'is this copy commission
	Dim tmpCCAction 
	tmpCCAction = Request.QueryString("OppdragAct")
	if len(trim(tmpCCAction)) > 0 and tmpCCAction = "copy" then
		ccAction = true
		vikarID = Request.QueryString("vikarID")
		ovid = Request.QueryString("OVID")
	end if

	' Get Oppdrag Information ?
	' Do we have an existing job ?
	If lngOppdragID > 0 Then

		strSQL = "SELECT OppdragID, StatusID, AvdelingskontorID, AvdelingID, tomID, CategoryID, AnsmedID,Terms, Typeid, Oppdrag.FirmaID, FraDato, FraKl, TilDato, TilKl, Beskrivelse," &_
		" Bestiltdato, bestiltklokken, BestilltAv, SOPeID, ArbAdresse, timepris, timerprdag, timeloenn, kurskode, Oppdragskode, ReportingOverrule, ReportingContactID, " &_
		" notatansvarlig, Kontaktperson = K.Fornavn + ' ' + K.Etternavn, notatokonomi, isnull(InvoiceComments,'') as InvoiceComments,lunch, notatvikar, webPub, webSted, isnull(CustomerReference,'') AS CustomerReference, " &_
		" webBeskrivelse, WebOverskrift, NoOfPersons, isnull(ParentAssignment,0) AS ParentAssignment, FaID " &_
		"FROM Oppdrag, KONTAKT K " &_
		"WHERE OppdragID = " & lngOppdragID & " " &_
		" AND Bestilltav *= K.KontaktID "

		' Get job data		
		set rsOppdrag = GetFirehoseRS(StrSQL, Conn)

		' Move from recordset to variables
		lngOppdragID		= rsOppdrag("OppdragID")
		lngFirmaID			= rsOppdrag("FirmaID")			
		lAnsMedID			= rsOppdrag("AnsMedID")
		lFaId				= rsOppdrag("FaID")
		lTypeID				= rsOppdrag("TypeID")
		lStatusID			= rsOppdrag("StatusID")
		lAvdelingID			= rsOppdrag("AvdelingskontorID")
		lRegnskapID			= rsOppdrag("AvdelingID")
		itomID				= rsOppdrag("TomID")
		icategoryID			= rsOppdrag("CategoryID")
		strFraDato			= rsOppdrag("FraDato")
		noOfOpportunities		= rsOppdrag("NoOfPersons")
		parentAssignment = rsOppdrag("parentAssignment")
		strTilDato			= rsOppdrag("TilDato")
		strFraKl			= FormatDateTime( rsOppdrag("FraKl"), 4 )
		strTilKl			= FormatDateTime( rsOppdrag("TilKl"), 4 )
		Timerprdag			= rsOppdrag("Timerprdag")
		strBeskrivelse		= rsOppdrag("Beskrivelse")
		strBestiltdato		= rsOppdrag("BestiltDato")
		lBestilltAv			= rsOppdrag("Bestilltav")
		SOPeID				= rsOppdrag("SOPeID")
		terms               = rsOppdrag("Terms")
		ReportingContactID = rsOppdrag("ReportingContactID")
		
	
		bReportingOverrule		= rsOppdrag("ReportingOverrule")
		strKontaktperson	= rsOppdrag("Kontaktperson")
		if(lenb(strKontaktperson) > 0) then
			SOPeID = 0
		else
			lBestilltAv = 0
		end if		
		strBestiltkl		= FormatDateTime( rsOppdrag("Bestiltklokken"), 4)
		strCustomerReference    = rsOppdrag("CustomerReference")
		strNotatAnsvarlig	= rsOppdrag("NotatAnsvarlig")
		strNotatOkonomi		= rsOppdrag("NotatOkonomi")
		strInvoiceComments      = rsOppdrag("InvoiceComments")		
		Timepris		= rsOppdrag("Timepris")
		Timepris		= FormatNumber(Timepris,2)		
		timeloenn		= rsOppdrag("Timeloenn")
		timeloenn		= FormatNumber(timeloenn,2)		
		Lunsj			= FormatDateTime( rsOppdrag("lunch"), 4)
		strArbAdresse		= rsOppdrag("ArbAdresse")
		strOppdragskode		= rsOppdrag("Oppdragskode")
		iwebpub			= rsOppdrag("webPub")
		strSted			= rsOppdrag("webSted")
		strNotatWeb		= rsOppdrag("webBeskrivelse")
		strWebOverskrift	= rsOppdrag("WebOverskrift")
		strKurskode		= rsOppdrag("Kurskode")
		
		if(parentAssignment = 0) THEN
			tmpCCAction = tmpCCAction
		else
			tmpCCAction = "copy"
		end if

		if iwebpub = 1 THEN
			webPub = "CHECKED"
		else
			webPub = ""
		end If
		
		
		if(lStatusID <> 1) then
			webPubEnabled = " DISABLED "
			webPub = ""
		else
			webPubEnabled = ""
		end if

		If( CInt(strOppdragskode )  = 0   )Then
			strOppdrag = "CHECKED"
			isDisabledDirectRec ="DISABLED"
			if len(trim(tmpCCAction)) > 0 and tmpCCAction = "copy" then

			else
				isDisabledDirectNoCopy = "DISABLED"
			end if

			if(terms=1) then 
		   		strSQL = "SELECT Amount,isnull(noofpersons,0) AS noofpersons,HourlyRate FROM EC_Oppdrag_Terms WHERE  oppdragID=" & lngOppdragID

		    SET rsOppdragTerms =GetFirehoseRS(StrSQL, Conn)
		    
		    if(HasRows(rsOppdragTerms) = true) then
				amount			= rsOppdragTerms("amount")
				noofpersons		= rsOppdragTerms("noofpersons")
				hourlyRate      = rsOppdragTerms("HourlyRate")
			end if

			rsOppdragTerms.close
			set rsOppdragTerms  = nothing
			
			end if
		end if
		
	   if CInt( strOppdragskode )  = 2 Then
		    strRecruitment ="CHECKED"
		    isDisableOppdrag = "DISABLED"
		 
		    if lStatusID = 5 then
				OppdragStatusEnable= "DISABLED"
		  end if 
		    strSQL = "SELECT   Amount,isnull(noofpersons,0) AS noofpersons,RecruitmentDate,Comment FROM EC_Oppdrag_Terms WHERE TermType=1 AND oppdragID=" & lngOppdragID
		    SET rsOppdragTerms = GetFirehoseRS(StrSQL, Conn)
		    
		    if(HasRows(rsOppdragTerms) = true) then
				amount			= rsOppdragTerms("amount")
				noofpersons			= rsOppdragTerms("noofpersons")
				 reruitmentDate		= rsOppdragTerms("RecruitmentDate")
				comments			= rsOppdragTerms( "Comment" )
				
			end if
			
			rsOppdragTerms.close
			set rsOppdragTerms  = nothing
			
		End If

		' Close and release recordset
		rsOppdrag.Close
		Set rsOppdrag = Nothing
		
		' update from oppdrag_vikar
		if (ccAction and len(trim(ovid)) > 0) then
		
			strSQL = "SELECT OppdragVikarID, VikarID, Fradato, Tildato, OppdragID, FraKl, TilKl, Timeloenn, AntTimer, Timepris, Lunch " &_
			"FROM OPPDRAG_VIKAR " &_
			"where OppdragVikarID=" & ovid 
		
			set rsOppdragVikar = GetFirehoseRS(StrSQL, Conn)
		
			if(HasRows(rsOppdragVikar) = true) then
				strFraDato			= rsOppdragVikar("Fradato")
				strTilDato			= rsOppdragVikar("Tildato")
				strFraKl			= FormatDateTime( rsOppdragVikar("FraKl"), 4 )
				strTilKl			= FormatDateTime( rsOppdragVikar("TilKl"), 4 )
				Timerprdag			= rsOppdragVikar("AntTimer")
				Timepris			= rsOppdragVikar("Timepris")
				Timepris			= FormatNumber(Timepris,2)
				timeloenn			= rsOppdragVikar("Timeloenn")
				timeloenn			= FormatNumber(timeloenn,2)
				Lunsj				= FormatDateTime( rsOppdragVikar("Lunch"), 4)
			end if
			
			rsOppdragVikar.close
			set rsOppdragVikar  = nothing
		end if

		StrSQL = "Select * from HOTLIST Where status=1 And BrukerID=" & lngBrukerID & " And oppdragID=" & lngOppdragID
		set rsHotlist = GetFirehoseRS(StrSQL, Conn)
		if (HasRows(rsHotlist) = false) then
			blnShowHotList = true
			strAddToHotlistLink = "AddHotlist.asp?kode=1&oppdragID=" & lngOppdragID & "&kundeNavn=" & server.URLEncode(strFirma) & "&kundeNr=" & lngFirmaID
			strHotlistType = "oppdrag"
		else
			rsHotlist.close
		end if
		set rsHotlist = nothing
	Else
		strBestiltDato = CStr(Date)
		strBestiltKl = CStr(FormatDateTime( Time , 4 ))
		strBeskrivelse =""
		strFrakl = "08:00"
		strTilkl = "16:00"
		tmpDate = cdate(strTilkl) - cdate(strFrakl)
		Timerprdag = datepart("h", tmpDate, 2, 2) & "." & (datepart("n", tmpDate, 2, 2) / 60 * 100)
		Lunsj = "00:00"
		strOppdrag = "CHECKED"
		lAnsMedID = lngBrukerID
		lStatusID = 1
		If Trim(lngBrukerID) = "" or isNull(lngBrukerID) Then
			lAnsMedID = 0
		End If

		strSQL = "SELECT [vikar_arbeidssted].[AvdelingskontorID] " & _
		"FROM [medarbeider] " & _
		"INNER JOIN [vikar] ON [medarbeider].[vikarID] = [vikar].[vikarID] " & _
		"INNER JOIN [vikar_arbeidssted] ON [vikar_arbeidssted].[vikarID] = [vikar].[vikarID] " & _
		"WHERE " & _
		"[medarbeider].[medid] = "  & Session("medarbID")

		set rsAvdelingKontor = GetFirehoseRS(StrSQL, Conn)

		if (HasRows(rsAvdelingKontor) = true) then
			lAvdelingID = rsAvdelingKontor("AvdelingskontorID").value
			rsAvdelingKontor.close
		end if
		Set rsAvdelingKontor  = nothing
	End If
	
	'STH FA changes 131106
	'check accepted temps
	if lngoppdragID > 0 then
	
		'strSQL = "SELECT TOP 1 V.EtterNavn,V.ForNavn,V.VikarId FROM OPPDRAG_VIKAR OV,VIKAR V " & _
		'	"WHERE OV.StatusId = 4 AND OV.OppdragId = " & lngoppdragID & _
		'	"AND OV.VikarId *= V.VikarId " & _
		'	"ORDER BY OV.OppdragVikarId DESC"
		
		strSQL = "SELECT TOP 1 V.EtterNavn,V.ForNavn,V.VikarId,VA.Ansattnummer FROM OPPDRAG_VIKAR OV,VIKAR V,VIKAR_ANSATTNUMMER VA " & _
			"WHERE OV.StatusId = 4 AND OV.OppdragId = " & lngoppdragID & _
			"AND OV.VikarId *= V.VikarId " & _
			"AND OV.VikarId *= VA.VikarId " &_
			"ORDER BY OV.OppdragVikarId DESC"
	
		 
		set rsVikarStatus = GetFirehoseRS(strSQL, Conn)
	
		if (HasRows(rsVikarStatus) = false) then
			strVikarFirstName = ""
			strVikarLastName = ""
			acceptedVikar = 0
			ansattNummer = 0
			hasAcceptedTemp = false
			rsVikarStatus.Close
			Set rsVikarStatus = Nothing
		Else
			strVikarFirstName = rsVikarStatus("EtterNavn")
			strVikarLastName = rsVikarStatus("ForNavn") & " " & rsVikarStatus("EtterNavn")
			ansattNummer = rsVikarStatus("Ansattnummer")
			hasAcceptedTemp = true
			acceptedVikar = rsVikarStatus("VikarId")
			rsVikarStatus.Close
			Set rsVikarStatus = Nothing
		End if
	
	End if
	

	' Get Customer Information ?
	' Do we have a Firm
	If clng(lngFirmaID) > 0 Then

		' Get Customer Info
		StrSQL = "Select Firma, SOCuID, isnull(CRMAccountGuid,'') AS CRMAccountGuid FROM FIRMA WHERE FirmaID = " & lngFirmaID
		set rsFirma = GetFirehoseRS(StrSQL, Conn)

		if (HasRows(rsFirma) = false) then
			rsFirma.Close
			Set rsFirma = Nothing
			AddErrorMessage("Kontakt ble ikke funnet!")
			call RenderErrorMessage()
		end if

	   ' Move from recordset to variables
	   strFirma = rsFirma("Firma")
	   SOCuID = rsFirma("SOCuID")
	   strCRMAccountGuid = rsFirma("CRMAccountGuid")

	   ' Close and release recordset
	   rsFirma.Close
	   Set rsFirma = Nothing
	
		
	   'Get work address from Superoffice if new task 
		if (SOCuID > 0 AND lngOppdragID = 0) then
			'Set cts = server.CreateObject("Integration.SuperOffice")
			'set rsAddress = cts.GetAddressByContactId(clng(SOcuID), 2) 'Get workadress
			'if HasRows(rsAddress) then
			'	strArbAdresse = rsAddress("address1")
			'end if
			'set rsAddress = nothing
			'set cts = nothing
			Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			aXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetAccountAddress&Socuid=" + Cstr(SOCuID) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
			aXmlHTTP.send ""

                	strArbAdresse = aXmlHTTP.responseText 


		end if
	End If

	' Create title and heading in page
	strHeading = "Nytt oppdrag"
	If (lngOppdragID > 0) Then
		if (ccAction) then
			strHeading = "Kopier oppdrag"
		else
			strHeading = "Endre oppdrag"
		end if	     
	End If

	if isNull(lFaId) then
		bChooseSelected = false
		lFaId = 0
	End If
	
	if isNull(ansattNummer) then
		ansattNummer = 0
	End If
	

	'Top Menu init
	dim strALinkStart
	dim strALinkEnd
	dim AToolbarAtrib(7,3) '0 = Enable/Disable, 1 = from/Link to activate, 2 = close link form string
	'Lagre Oppdrag
	AToolbarAtrib(0,0) = "1"
	'AToolbarAtrib(0,1) = "<a href='javascript:document.all.frmJobNew.submit();' title='Lagre oppdrag'>"
	If lngOppdragID > 0 Then
	AToolbarAtrib(0,1) = "<a href='javascript:SaveWithAcceptedVikar(""" & strVikarLastName & """," & ansattNummer & "," & lFaId & ");' title='Lagre oppdrag'>"
	Else
		AToolbarAtrib(0,1) = "<a href='javascript:document.all.frmJobNew.submit();' title='Lagre oppdrag'>"
	End If
	AToolbarAtrib(0,2) = "</a>"
	'Vise oppdrag
	if (ccAction) then
		AToolbarAtrib(1,0) = "0"
	else
		AToolbarAtrib(1,0) = "1"
	end if
	AToolbarAtrib(1,1) = "<a href='WebUI/OppdragView.aspx?OppdragID=" & lngOppdragID & "' title='Vis oppdrag'>"
	AToolbarAtrib(1,2) = "</a>"
	'til endre Konsulent
	AToolbarAtrib(2,0) = "0"
	AToolbarAtrib(2,1) = ""
	AToolbarAtrib(2,2) = ""
	'Tilknytte konsulent
	AToolbarAtrib(3,0) = "0"
	AToolbarAtrib(3,1) =  ""
	AToolbarAtrib(3,2) = ""
	'Aktiviteter for oppdrag
	if (ccAction) then
		AToolbarAtrib(4,0) = "0"
	else
		AToolbarAtrib(4,0) = "1"
	end if 
	AToolbarAtrib(4,1) = "<a href='AktivitetOppdrag.asp?OppdragID=" & lngOppdragID & "' title='Vis aktiviteter for vikaren'>"
	AToolbarAtrib(4,2) = "</a>"
	'Kalender
	if (ccAction) then
		AToolbarAtrib(5,0) = "0"
	else
		AToolbarAtrib(5,0) = "1"
	end if
	AToolbarAtrib(5,1) = "<form ACTION='Kalender.asp?OppdragID=" & lngOppdragID & "' METHOD='POST' name='frmJobCalendar'></form>" & _
	"<a href='javascript:document.all.frmJobCalendar.submit();' title='Vis kalender for oppdrag'>"
	AToolbarAtrib(5,2) = "</a></form>"
	
	AToolbarAtrib(6,0) = "0"
	AToolbarAtrib(6,1) = "<input NAME='hdnCopyCommission' TYPE='hidden' VALUE='Copy Commission'>" & _
	"<a href='javascript:document.all.formJobChange.submit();' title='Kopier oppdrag'>"
	AToolbarAtrib(6,2) = "</a>"
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading%></title>
		<script type="text/javascript" src="/xtra/Js/dateFuncs.js"></script>
		<script type="text/javascript" src="/xtra/Js/javascript.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script type="text/javascript" src='/xtra/js/menu.js' id='menuScripts'></script>
		<script type="text/javascript" src='/xtra/js/navigation.js' id='navigationScripts'></script>
		<script language='javascript' src='js/ecajax.js'></script>
		<%
		'dato for reg av oppdrag
		dim strDay
		dim strMonth
		dim strYear
		dim dto

		strYear = year(now())
		strMonth = month(now())
		strDay = day(now())

		if cint(strMonth)< 10 then
			strMonth = "0" & strMonth
		end if

		if cint(strDay)< 10 then
			strDay = "0" & strDay
		end if

		dto = replace(strYear & strMonth & strDay, " ", "")
		
                
		%>
		<script language="javaScript" type="text/javascript">
		//Funksjon for å advare hvis regdato på oppdrag ligger bakover i tid
		
		 
		var currentValue = -1 ;
		
		
		function SetTimeLoanField()  
		{
		   var hourlyPay = document.getElementById("tbxTimeloenn");
		   var enterValue = document.getElementById("txtHourlyRate").value ;
		   hourlyPay.value = enterValue ;
		}		
		
		function SetOppdragStatus()	{		
			var checkBox = document.getElementById("Radio2");
			if(checkBox.checked){
		  		var combobox = document.getElementById("Select3");
		  		var selectedTex = combobox.options[combobox.selectedIndex].text;
		    	if(selectedTex =="Fullstendig"){		     
		       		var answer =   confirm("Lagrer du med status 'Fullstendig' kan oppdragstatus ikke bli endret senere, og oppdraget blir fakturert. Ønsker du å fortsette?");		       
		       		if(!answer){
		          		combobox.value = currentValue ;		              
		       		}
		       		// change the Kommentor text area to be mandatory		       
		       		document.getElementById("tbxComment").className = "mandatory"; 		       
		     	}
		     	else{
		     		// change the Kommentor text area to be NON mandatory	
		     		var cls = "mandatory";
		     		var reg = new RegExp('(\\s|^)'+cls+'(\\s|$)');	       		
		       		document.getElementById("tbxComment").className = document.getElementById("tbxComment").className.replace(reg,'');

		     	}
			}
		}
		
		function SetKommentStatus(){		
			var checkBox = document.getElementById("Radio2");
			if(checkBox.checked){
		  		var combobox = document.getElementById("Select3");
		  		var selectedTex = combobox.options[combobox.selectedIndex].text;
		    	if(selectedTex =="Fullstendig"){
		       		// change the Kommentor text area to be mandatory		       
		       		document.getElementById("tbxComment").className = "mandatory"; 		       
		     	}
			}
		}		
		
		function GetCurrentValue(){
		    currentValue = document.getElementById("Select3").value ;		
		}
	
		function SetShowHide(div)
         {
           if(div ==1 )
           {
             document.getElementById("OppdragDiv").style.display = "block";
             document.getElementById("DirectRecDiv").style.display = "None";
             document.getElementById("TermsDiv").style.display = "block";
             document.getElementById("AmountDiv").style.display = "None";
             document.getElementById("BaseRateDiv").style.display = "None";
             document.getElementById("NoOfOpportunitiesDiv").style.display = "block";
             
             document.all.TermsCombo.selectedIndex =  0;
             
           }
           if(div ==2)
           {
             document.getElementById("DirectRecDiv").style.display = "block";
             document.getElementById("OppdragDiv").style.display = "None";
             document.getElementById("TermsDiv").style.display = "None";
             document.getElementById("AmountDiv").style.display = "block";
             document.getElementById("NoOfPersonsDiv").style.display = "block";
             document.getElementById("BaseRateDiv").style.display = "None";
             document.getElementById("NoOfOpportunitiesDiv").style.display = "None";
             
           }
                 
          }
          
          function SetDiv()
         {
            
         	 
           if( document.getElementById("Radio1").checked )
           {
             document.getElementById("OppdragDiv").style.display = "block";
             document.getElementById("DirectRecDiv").style.display = "None";
             document.getElementById("TermsDiv").style.display = "block";
              document.getElementById("AmountDiv").style.display = "None";
              document.getElementById("BaseRateDiv").style.display = "None";
              document.getElementById("NoOfOpportunitiesDiv").style.display = "block";
              
              
              SetSalryAmountTextBox();
           
             
           }
           if(document.getElementById("Radio2").checked )
           {
             document.getElementById("DirectRecDiv").style.display = "block";
             document.getElementById("OppdragDiv").style.display = "None";
             document.getElementById("TermsDiv").style.display = "None";
              document.getElementById("AmountDiv").style.display = "Block";
              document.getElementById("NoOfPersonsDiv").style.display = "Block";
              document.getElementById("BaseRateDiv").style.display = "None";
              document.getElementById("NoOfOpportunitiesDiv").style.display = "None";
              
           }
              
          }

		function imposeMaxLength(Object, MaxLen)
		{
  			return (Object.value.length <= MaxLen);
		}

            
		    function SetSalryAmountTextBox() 
		    {
		    
		
              if( document.all.TermsCombo.selectedIndex == 1 )
              {
               
               document.getElementById("AmountDiv").style.display = "Block";
                document.getElementById("BaseRateDiv").style.display = "Block";
               
//                document.all.txtAmount.disabled=false;
//                document.all.txtAmount.style.backgroundColor = '#ffff99';
              }
             else
             {
                 document.getElementById("AmountDiv").style.display = "None";
                 document.getElementById("BaseRateDiv").style.display = "None";
                 
//                 document.all.txtAmount.disabled=true;
//                 document.all.txtAmount.style.backgroundColor = 'LightGrey';
               }
          }



			var blnErrorOccured

			function fraDato()
			{
			
				var diff;
				var fraDato;
				var dato;
				var aar;

				dato = parseInt(<%=Trim(dto)%>);
				fraDato = document.forms[0].elements['tbxFraDato'].value;
				aar = (fraDato.substring(6,8));

				if (aar < '70')
				{
					aar = '20';
				}
				else
				{
					aar = '19';
				}

				fraDato = parseInt(aar + fraDato.substring(6, 8) + fraDato.substring(3, 5) + fraDato.substring(0, 2));
				diff = dato - fraDato;
				if(diff > 30)
				{
					alert("Fra dato på oppdraget er før denne måned! Du kan ikke lage oppdraget dersom det er fakturert og utbetalt lønn for måneden.");
				}

			}

			function dateNotPassed(tbxDate)
			{
				var diff;
				var dato;
				var aar;

				dato = parseInt(<%=Trim(dto)%>);
				aar = (tbxDate.substring(6,8));

				if (aar < '70')
				{
					aar = '20';
				}
				else
				{
					aar = '19';
				}

				tbxDate = parseInt(aar + tbxDate.substring(6, 8) + tbxDate.substring(3, 5) + tbxDate.substring(0, 2));
				diff = dato - tbxDate;

				if(diff > 0)
				{
					return false;
				}
				return true;
			}

			function SetAndPostAction(strAction)
			{
				document.all.hdnJobAction.value = strAction;
				document.all.frmJobNew.submit();
			}
			
			function ToggleReportingContact()
			{
				if(document.all.questback.checked)
				{
					document.all.dbxQuestbackP.disabled = true;
					document.all.dbxQuestbackP.selectedIndex = 0;
				}
				else
				{
					document.all.dbxQuestbackP.disabled = false;
					if(document.all.dbxSOKontaktP != null)
						document.all.dbxQuestbackP.selectedIndex = document.all.dbxSOKontaktP.selectedIndex
				}
			}
			
			function EnOrDeReportingContact()
			{
				<% if isNull(SOCuID) then %>					
					document.all.dbxQuestbackP.disabled = true;
					document.all.questback.disabled = true;
				<% elseif (bReportingOverrule) then %>
					document.all.dbxQuestbackP.disabled = true;
					document.all.dbxQuestbackP.selectedIndex = 0;
				<% else %>									
					document.all.dbxQuestbackP.disabled = false;
				<% end if %>

			}
			
			function UpdateQuestback()
			{							
					document.all.dbxQuestbackP.selectedIndex = document.all.dbxSOKontaktP.selectedIndex
			}
			
			function SaveWithAcceptedVikar(vikar,ansattNummer,iniFaId)
			{
				<% if (ccAction <> true) then %>
					var oppdragNo = document.all.dbxOppdragNr.value;
				<% end if %>
				var faId = document.all.dbxFAgreement.value;
				var description;
				var startDate;
				var endDate;
				var status;				
				
				<% If lngOppdragID > 0 Then %>
				       
				    <% If (ccAction) and (hasAcceptedTemp = true) Then %> 

						var cpaction = "";
						<%
						if (len(trim(vikarID)) > 0) Then						
						%>  
							cpaction = "copyassign";
						<%
						else
						%>
							cpaction = "copy";
						<%
						end if
						%>
						
						 if (iniFaId == faId) 
						 {
							document.all.hdnJobAction.value = cpaction;
							document.all.frmJobNew.submit();
						 }
						 else
						 {
						 	if (faId == 0)
						 	{
						 		document.all.hdnJobAction.value = cpaction;
								document.all.frmJobNew.submit();
								
							}
							else
							{
				                                 <%  If (ccAction <> true) Then %>

									var popup = window.open('FAgreementCategoryVikar.asp?vikarName='+ vikar + '&OppdragNr='+ oppdragNo + '&VikarId='+ ansattNummer + '&FaId='+ faId + '&cpaction=' + cpaction ,'CategoryAssigner'
									,'top = 150,location=no,height=180,width=750,menubar=no,resizable=no,toolbar=no,status=yes,scrollbars=no');
								
									popup.focus();

								<% else %>

								document.all.hdnJobAction.value = cpaction;
								document.all.frmJobNew.submit();

								<% end if %>
										
							}
						 }
						
					<% ElseIf (ccAction) Then 						
						if (len(trim(vikarID)) > 0) Then						
						%>  
							document.all.hdnJobAction.value = "copyassign";
						<%
						else
						%>
							document.all.hdnJobAction.value = "copy";
						<%
						end if
						%>
						document.all.frmJobNew.submit();
											
					<% ElseIf (hasAcceptedTemp = true) Then %>						
						 if (iniFaId == faId) 
						 {
							document.all.hdnJobAction.value = "lagre";
							document.all.frmJobNew.submit();
						 }
						 else
						 {
						 	if (faId == 0)
						 	{
						 		document.all.hdnJobAction.value = "lagre";
								document.all.frmJobNew.submit();
								
							}
							else
							{
								var popup = window.open('FAgreementCategoryVikar.asp?vikarName='+ vikar + '&OppdragNr='+ oppdragNo + '&VikarId='+ ansattNummer + '&FaId='+ faId + '&cpaction=lagre','CategoryAssigner'
								,'top = 150,location=no,height=180,width=750,menubar=no,resizable=no,toolbar=no,status=yes,scrollbars=no');
								
								popup.focus();
							}
						 }
					<% Else %>						
						document.all.hdnJobAction.value = "lagre";
						document.all.frmJobNew.submit();
					<% End If %>
				<% Else %>
					document.all.hdnJobAction.value = "lagre";
					document.all.frmJobNew.submit();
				<% End If %>
			}

			function shortKey(e)
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				if (modKey && modKey2 && keyChar == "S")
				{
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
			}

			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
			

		var handleCallback = function (result) 
		{
			if(result) 
			{
				var str = "<select name='dbxMedarbeider' id='dbxMedarbeider' onkeydown='typeAhead()'>";
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
		
		function LoadCategories()
		{
			var url;
			url = "Callback.asp?FuncName=GetOppdragCategories&TomID=";
			url = url.concat(document.all.dbxtjenesteomrade.value);
			asynchronousCall(url,handleCategories);

		}
		
		var handleCategories = function (result)
		{
			if(result) 
			{
				var str = "<select name='dbxCategory' id='dbxCategory' class='mandatory' >";
				str = str.concat(result,"</SELECT>");
				document.getElementById('divCommissionCategory').innerHTML = str;
			}
		};
		
		</script>
	 
	</head>
	<body onload="fokus(); EnOrDeReportingContact();SetDiv();SetKommentStatus();">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
				<!--#include file="Includes/Top_Menu_job2.asp"-->
			</div>
			<% if (ccAction) then %>
			<div class="content" style="height:20px; padding:10px; BACKGROUND-COLOR: #66CCFF !important; VERTICAL-ALIGN: middle; text-align:center; " >
			<font style="FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 11px;  FONT-WEIGHT: bold;">
				Du er nå i ?kopier oppdrag? modus. Nytt oppdrag blir ikke laget før du har lagret!				
			</font>
			</div>
			<% end if %>
			<div class="content">
				<form action="OppdragDB.asp" method="POST" name="frmJobNew" id="Form1">
					<input name="OppdragID" type="HIDDEN" value="<%=lngOppdragID%>" id="Hidden2">
					<input name="tbxStatus" type="HIDDEN" value="<%=lStatusID%>" id="Hidden1">
					<input name="tbxOppdragskode" type="HIDDEN" value="<%=strOppdragskode%>" id="Hidden4">
					<input name="hdnJobAction" type="HIDDEN" value="lagre" id="Hidden5">
					<input name="hdnVikarID" type="hidden" value="<%=vikarID%>" id="hiddencc1">
					<input name="hdnTerm" type="hidden" value="<%=terms%>" id="Hidden6">
					<input name="hdnNotRecuit" type="hidden" value="<%=isDisabledDirectNoCopy%>" id="Hidden7">				
					
					<table id="Table1">
						<col width="50%">
						<col width="50%">
						<tr>
							<td>
								<table id="Table2">
									<%
									If (clng(lngOppdragID) = 0) or (ccAction) Then
										%>
										<tr>
											<th>Oppdragnr:</th>
											<td>-&nbsp;</td>
										</tr>
										<tr>
											<th>Kontaktnr:</th>
											<td>
												<input id=lnk0 class="mandatory" name="FirmaID" type="TEXT" size=6 maxlength="6" value="<%if clng(lngFirmaID) > 0 then%><%=lngFirmaID%><% end if%>">
											</td>
											<td>
												&nbsp;
											</td>
										</tr>
										<%
									ElseIf (clng(lngOppdragID) > 0) Then
										%>
										<tr>
											<th>Oppdragnr:</th>
											<td><%=lngOppdragID%><input type="hidden" value="<%=lngOppdragID%>" name="dbxOppdragNr" id="dbxOppdragNr"></td>
										</tr>
										<tr>
											<th>Kontaktnr:</th>
											<td><%=lngFirmaID%><td>
												<input name="FirmaID" type="HIDDEN" value="<%=lngFirmaID%>" id="Hidden3">
											</td>
										</tr>
										<%
									End If
									%>
									<tr>
										<th>Kontakt:</th>
										<%
										linkurl = Application("CRMAccountLink") & strCRMAccountGuid & "%7d&pagetype=entityrecord"
										'response.write strCRMAccountGuid
										if (SOCuID > 0) then
											%>
											<td><a href='<%=linkurl%>' target='_blank'><%=strFirma%></a></td>
											<%
										else
											%>
											<td><%=strFirma%></td>
											<%
										end if
										%>
									</tr>
								</table>
							</td>
							<td>
								<table id="Table3">
									<tr>
										<th>Kontaktperson.:</th>
										<td>
											<%
											StrSQL = "SELECT COUNT(*) AS TSCount FROM DAGSLISTE_VIKAR WHERE OppdragID=" &lngOppdragID
											set rsTimesheetCount = GetFirehoseRS(StrSQL, Conn)
											Do Until rsTimesheetCount.EOF
												NoOfTimesheets = rsTimesheetCount("TSCount")													
												rsTimesheetCount.MoveNext
											Loop
											rsTimesheetCount.Close
											Set rsTimesheetCount = Nothing
											%>
											
											<%
											If (lngOppdragID > 0) Then
												if len(trim(tmpCCAction)) > 0 and tmpCCAction = "copy" then
													contactDisabled = ""													
												Else			
													If (NoOfTimesheets > 0) Then														
														contactDisabled = "disabled"
													Else														
														contactDisabled = ""
													End If													
												End If
											Else
												contactDisabled = ""												
											End If
											Set oXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
											oXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetContacts&Socuid=" + Cstr(SOCuID) + "&Sopeid=" +  Cstr(SOPeID), False , Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")

											
											'response.write "SOcuID " & SOCuID
											'response.write "@@@"												
											
											if (clng(SOPeID) > 0 OR clng(lBestilltAv) = 0) then	
												'response.write "aa"											
													
												If lngOppdragID > 0 Then											
													'personsHTML = cts.HTMLGetPersonsForContactAsDropDown(clng(SOcuID), clng(SOPeID), "dbxSOKontaktP", "(Ingen valgt)", "0", " class='mandatory' onchange='UpdateQuestback()' " & contactDisabled , true)
												else
													'personsHTML = cts.HTMLGetPersonsForContactAsDropDown(clng(SOcuID), clng(SOPeID), "dbxSOKontaktP", "(Ingen valgt)", "0", " class='mandatory' onchange='UpdateQuestback()' " & contactDisabled , false)
												end if
												'Response.Write personsHTML
												oXmlHTTP.send ""
												Response.Write "<select name='dbxSOKontaktP' class='mandatory' onchange='UpdateQuestback()' ><option value='0'>(Ingen valgt)</option>"
												Response.Write oXmlHTTP.responseText
												Response.Write "<input type='hidden' value='" & SOPeID & "' name='dbxSOPeID' id='dbxSOPeID'>"

			
											end if
											
											'response.write "SOpeID" & clng(SOPeID)	
											
											if (clng(lBestilltAv) > 0) then
												%>
												<%=strKontaktperson%><input type="hidden" value="<%=lBestilltAv%>" name="dbxKontaktP" id="dbxKontaktP">
												<input type="hidden" value="<%=strKontaktperson%>" name="dbxKontaktName" id="dbxKontaktName">
												<%
											end if
											%>
											
											<input type="hidden" value="<%=NoOfTimesheets%>" name="dbxNoOfTimesheets" id="dbxNoOfTimesheets">
										</td>
									</tr>								
									<tr>
										<th>
											Dato/tid:
										</th>
										<td>
											<input name="tbxBestDato" type="TEXT" size="8" maxlength="8" value="<%=strBestiltDato%>" onblur="dateCheck(this.form, this.name)" id="Text1">
											<input name="tbxBestKl" type="TEXT" size="4" maxlength="5" value="<%=strBestiltkl%>" onblur="timeCheck(this.form, this.name)" id="Text2">
										</td>
									</tr>
									<tr>
										<th>
											Referanse:
										</th>
										<td>
											<input name="tbxCustomerReference" type="TEXT" size="20" value="<%=strCustomerReference%>" id="customerreference">											
										</td>
									</tr>
								</Table>
							<td>
						</tr>
					</table>
				</div>
				<div class="contentHead">
					<h2>Oppdragsinformasjon</h2>
				</div>
				<div class="content">
					<table id="Table4">
						<col width="50%">
						<col width="50%">
						<tr>
							<td>
								<table id="Table6">
									<tr>
										<td>
											<table id="Table5">
												<tr>
													<th>Arb.adresse:</th>
													<td colspan="2"><input id="lnk2"  type="TEXT" name="tbxArbAdresse" size="40" maxlength="50" value="<%=strArbAdresse%>"></td>
												</tr>
												<tr>
													<th>Beskrivelse:</th>
													<td colspan="8"><input id="lnk4" class="mandatory"  name="tbxBeskrivelse" type="TEXT" size="70" maxlength="255" value="<%=strBeskrivelse%>"></td>
												</tr>
											</table>
										<td>
									</tr>
								</table>
							</td>
						</tr>
					</Table>
				</div>
				<div class="contentHead">
					<h2>Questback - undersøkelse</h2>
				</div>
				<div class="content">
					<table id="Table4">
						<col width="50%">
						<col width="50%">
						<tr>
							<td>
								<table id="Table6">
									<tr>
										<td>
											<table id="Table5">
												<tr>
												<th>Skal ikke motta:</th>
													<td colspan="8"><input class="checkbox" type="checkbox" id="questback" name="questback" disabled="true" onclick="ToggleReportingContact()" <%If (bReportingOverrule) Then Response.Write "checked"%>></td>
													
												</tr>
												<tr>
													<th>Kontaktperson:</th>
													<td colspan="2">
													<% 
														Set oXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
														
		
					
														If (SOCuID > 0) then
															If isNull(ReportingContactID) OR (ReportingContactID = 0) Then
																If clng(SOPeID) > 0 Then																
																	ReportingContactID = clng(SOPeID)
																Else
																	ReportingContactID = 0
																End If
															End If	
															oXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetContacts&Socuid=" + Cstr(SOCuID) + "&Sopeid=" +  Cstr(ReportingContactID) , False , Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")																													
															'personsHTML = cts.HTMLGetPersonsForContactAsDropDown(clng(SOcuID), clng(ReportingContactID), "dbxQuestbackP", "(Alternativ kontaktperson)", "0", " style='width:400' ", False)												
														Else
															oXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetContacts&Socuid=0"  + "&Sopeid=0"  , False , Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
															'personsHTML = cts.HTMLGetPersonsForContactAsDropDown(0, 0, "dbxQuestbackP", "(Alternativ kontaktperson)", "0", "", False)
															%>
															<input type="hidden" value="old" name="dbxKontaktPerson" id="dbxKontaktPerson">
															<%
														End If
														
														'Response.Write personsHTML
														oXmlHTTP.send ""
														Response.Write "<select name='dbxQuestbackP' style='width:400'><option value='0'>(Alternativ kontaktperson)</option>"
														Response.Write oXmlHTTP.responseText															
															
													%>
													</td>
												</tr>
											</table>
										<td>
									</tr>
								</table>
							</td>
						</tr>
					</Table>
				</div>
				<div class="contentHead">
					<h2>Diverse / Tilhørighet</h2>
				</div>
				<div class="content">
					<table id="Table7">
						<col width="35%">
						<col width="30%">
						<col width="35%">
						<tr>
							<td style="width:40%">
							
							
							
							
							<table id="Table10" width ="100%">
									<tr>
										<th>Type:</th>
										<td>
										 
											<input name="rbnKurskode" type="RADIO" onclick ="javascript:SetShowHide(1);"  class="radio" <%=isDisableOppdrag %> value="0" <%=strOppdrag%> id="Radio1">Oppdrag
											<input name="rbnKurskode" type="RADIO" onclick="javascript:SetShowHide(2);"  class="radio" <%=isDisabledDirectRec %> value="3" <%=strRecruitment%> id="Radio2">	Rekruttering
											 
										</td>
									</tr>
									 
										<tr id ="TermsDiv" style="display:none;">
										<th>Vilkår:</th>
										<td>
											<select name="TermsCombo" id="TermsCombo" onchange ="SetSalryAmountTextBox()" class="mandatory" <%=isDisabledDirectNoCopy %>>
											
										     	<%
										     	if(terms= 0) then
										     	   sel="SELECTED"
										     	else
										     	  sel=" "
										     	end if
										     	%>
										     	 <option value ="0" <%=sel%> >Timelønn</option>
										     	 <% 
										     	  if(terms= 1) then
										     	   sel="SELECTED"
										     	  else 
										     	   sel=" "
										     	  end if
										     	  %>
										     	 <option value="1" <%=sel%> >Fast lønn</option>
										   
											</select>
											 
										</td>
									</tr>									
									
									<tr id="AmountDiv" style="display:none;">
										<th>Beløp:</th>
										<td> 
											<input id="txtAmount" class="mandatory" name="txtAmount" type="TEXT" size="8" maxlength="12"  value="<%=amount%>"  >									 
										</td>
									</tr>
									
									<tr id="NoOfPersonsDiv" style="display:none;">
										<th>Antall personer:</th>
										<td> 
											<input id="txtNoOfPersons" class="mandatory" name="txtNoOfPersons" type="TEXT" size="8" maxlength="12"  value="<%=noofpersons%>"  >
										</td>
									</tr>
									
									<tr id="BaseRateDiv" style="display:none;">
										<th>Timesats:</th>
										<td>
 
											<input id="txtHourlyRate" class="mandatory" onchange ="SetTimeLoanField()" name="txtHourlyRate" type="TEXT" size="8" maxlength="12"  value="<%=hourlyRate%>"  >
									 
										</td>
									</tr>
									
									<tr>
										<th>Status:</th>
										<td>
										 	<select name="dbxStatus" id="Select3" <%=OppdragStatusEnable %> onchange  ="SetOppdragStatus()"  onfocus  ="GetCurrentValue()" >
					                           	<%
											StrSQL = "SELECT OppdragsStatusID, OppdragsStatus FROM h_oppdrag_status"
											set rsStatus = GetFirehoseRS(StrSQL, Conn)
											Do Until rsStatus.EOF
												If rsStatus("OppdragsStatusID") = lStatusID Then
													sel = "SELECTED"
												Else
													sel = ""
												End If
												%>
												<option value="<%=rsStatus("OppdragsStatusID")%>" <%=sel%>><%=rsStatus("Oppdragsstatus") %></option>
												<%
												rsStatus.MoveNext
											Loop
											rsStatus.Close
											Set rsStatus = Nothing
											%>
											</select>
										</td>
									</tr>
									<!-- This is NoOfPersons for normal recruitments -->
									<tr id="NoOfOpportunitiesDiv" style="display:none;">
										<th>Antall personer:</th>
										<td> 											
											
										<%
											if len(trim(tmpCCAction)) > 0 and tmpCCAction = "copy" then
										%>									
											<input id="txtNoOfOpportunities" class="mandatory" name="txtNoOfOpportunities" type="TEXT" size="8" maxlength="12" disabled="disabled" value="1"  >
										<%				
											Else
										%>
											<input id="txtNoOfOpportunities" class="mandatory" name="txtNoOfOpportunities" type="TEXT" size="8" maxlength="12"  value="<%=noofopportunities%>"  >
										
										<%
											End If									
										%>
										</td>
									</tr>	
								</table>
								
							</td>
							<td  style="width:29%">
							 <div id ="OppdragDiv"  style="display:none;width:"100%" > 
								<table id="Table9">
									<tr>
										<th>Dato:</th>
										<td>
											<input id="lnk8" class="mandatory"  name="tbxFraDato" type="TEXT" size="8" maxlength="8" value="<%=strFraDato%>" onblur="dateCheck(this.form, this.name), fraDato(this.form, this.name)">
											-
											<input id="lnk9" class="mandatory"  name="tbxTilDato" type="TEXT" size="8" maxlength="8" value="<%=strTilDato%>" onblur="dateCheck(this.form, this.name), dateInterval(this.form, this.name)">
										</td>
									</tr>
									<tr>
										<th>Klokken:</th>
										<td>
											<input id="lnk10"  name="tbxFraKl" type="TEXT" size="4" maxlength="5" value="<%=strFrakl%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name)">
											-
											<input id="lnk11"  name="tbxTilKl" type="TEXT" size="4" maxlength="5" value="<%=strTilkl%>" onchange="timeCheck(this.form, this.name), workTime(this.form, this.name)">
										</td>
									</tr>
									<tr>
										<th>Lunsj:</th>
										<td>
											<input class="mandatory" id="lnk12"  name="tbxLunsj" type="TEXT" size="5" maxlength="9" value="<%=Lunsj%>" onblur="timeCheck(this.form, this.name), workTime(this.form, this.name)">
										</td>
									</tr>
									<tr>
										<th>Timer pr.dag:</th>
										<td><input id="lnk15"  name="tbxTimerPrDag" type="TEXT" size="5" maxlength="9" value="<%=Timerprdag%>"></td>
									</tr>
									<tr>
										<th>Timepris:</th>
										<td><input class="mandatory" id="lnk16"  name="tbxTimePris" type="TEXT" size="5" maxlength="9" value="<% = Timepris %>"></td>
									</tr>
									<tr>
										<th>Timelønn:</th>
										<td><input class="mandatory" id="Text3"  name="tbxTimeloenn" type="TEXT" size="5" maxlength="9" value="<% = timeloenn %>"></td>
									</tr>
								</Table>
								</div>
								<div id ="DirectRecDiv"  style="display:none;width:"100%"   >
								<table id ="directTable">
								<tr>
								<th >
									
                            Rekruttering Dato :
								</th>
								<td>
								 <input class="mandatory" id="tbxRecruitmentDate" value="<%=reruitmentDate%>"  name="tbxRecruitmentDate" type="TEXT" size="8" maxlength="8" onblur="dateCheck(this.form, this.name)" >
								</td>
								</tr>
								<tr >
								<th colspan ="2" align="left">
								Kommentar på faktura : 
								</th>
								 
								</tr>
								<tr>
								<td colspan ="2">
								<textarea rows="15" cols="30" id="tbxComment" name="tbxComment"> <%=comments%></textarea> 
								  
								</td>
								</tr>
								</table>
								
								</div>
							</td>
							<td style="width:31%">
								
								 <table id="Table8" width ="100%">
									<tr>
										<th>Avdelingskontor:</th>
										<td>
											<select name="dbxAvdeling" id="Select1">
												<%
												StrSQL = "SELECT id, navn FROM avdelingskontor WHERE show_hide = 1 ORDER BY id"
												set rsAvdeling = GetFirehoseRS(StrSQL, Conn)
												Do Until rsAvdeling.EOF
													If rsAvdeling("ID") = lAvdelingID Then
														sel = " SELECTED"
													Else
														sel = ""
													End If
													Response.Write "<option VALUE='" & rsAvdeling("ID") & "' " & sel & " >" & rsAvdeling("navn")
													rsAvdeling.MoveNext
												Loop
												' Close and release recordset
												rsAvdeling.Close
												Set rsAvdeling = Nothing
												%>
											</select>
										</td>
									</tr>
									<tr>
										<th>Regnskapsavdeling:</th>
										<td>
											<%
											if strHeading = "Nytt oppdrag" then
												strSQL = "SELECT isnull(AVDELING.AVDELINGID, 0) as AVDELINGID " & _
													"FROM medarbeider as med " & _
													"INNER JOIN vikar on med.vikarid = vikar.vikarid " & _
													"INNER JOIN VIKAR_AVDELING on VIKAR_AVDELING.vikarid = vikar.vikarid " & _
													"INNER JOIN AVDELING on AVDELING.AvdelingID = VIKAR_AVDELING.AVDELINGID " & _
													"WHERE med.medid = " & Session("medarbID")

												set rsDefaultRA = GetFirehoseRS(StrSQL, Conn)
												if (HasRows(rsDefaultRA) = true) then
													lRegnskapID = rsDefaultRA("AVDELINGID")
												end if
												rsDefaultRA.close
												set rsDefaultRA = nothing
											end if
											%>
											<select name="dbxregnskap" id="lnk6" >
												<%
												StrSQL = "select avdelingid, avdeling,Show_Hide from avdeling order by avdeling"
												set rsRegnskapskontor = GetFirehoseRS(StrSQL, Conn)
												while not (rsRegnskapskontor.EOF)
													If rsRegnskapskontor("avdelingid") = lRegnskapID Then
														sel = " SELECTED"
													Else
														sel = ""
													End If
													If (sel <> "") or (rsRegnskapskontor("Show_Hide")= "0")Then
													      Response.Write "<option VALUE='" & rsRegnskapskontor("avdelingid") & "' " & sel & " >" & rsRegnskapskontor("avdeling") & "</option>"
													End If
													rsRegnskapskontor.MoveNext
												wend
												rsRegnskapskontor.Close
												Set rsRegnskapskontor = Nothing
												%>
											</select>
										</td>
									</tr>
									<tr>
										<th>Tjenesteomr&aring;de:</th>
										<%
										if (strHeading = "Nytt oppdrag") then
											strSQL = "SELECT isnull(TJENESTEOMRADE.TomID, 0) as TomID " & _
												"FROM medarbeider AS med  " & _
												"INNER JOIN vikar ON med.vikarid = vikar.vikarid " & _
												"INNER JOIN VIKAR_TJENESTEOMRADE ON VIKAR_TJENESTEOMRADE.vikarid = vikar.vikarid " & _
												"INNER JOIN TJENESTEOMRADE ON TJENESTEOMRADE.TomID = VIKAR_TJENESTEOMRADE.TomID " & _
												" WHERE med.medid =" & Session("medarbID")

											set rsDefaultTOM = GetFirehoseRS(StrSQL, Conn)
											if (HasRows(rsDefaultTOM) = true) then
												itomID = rsDefaultTOM("TomID")
											end if
											rsDefaultTOM.close
											set rsDefaultTOM = nothing
										end if
										%>
										<td>
											<select name="dbxtjenesteomrade" id="dbxtjenesteomrade" onchange='LoadCategories()' >
												<%
												StrSQL = "SELECT tomid, navn = navn + ' ' + isnull(beskrivelse,'') FROM tjenesteomrade ORDER BY tomid"
												set rsTjenesteOmrade = GetFirehoseRS(StrSQL, Conn)
												if (HasRows(rsTjenesteOmrade) = true) then
													while not (rsTjenesteOmrade.EOF)
														If rsTjenesteOmrade("TomID") = itomID Then
															sel = " SELECTED"
														Else
															sel = ""
														End If
														Response.Write "<option VALUE='" & rsTjenesteOmrade("tomID") & "' " & sel & " >" & rsTjenesteOmrade("navn") & "</option>"
														rsTjenesteOmrade.MoveNext
													wend
												end if
												rsTjenesteOmrade.Close
												Set rsTjenesteOmrade = Nothing
												%>
											</select>
										</td>
									</tr>
									<tr>
									<th>Kategori:</th>
									<td>
										<div id="divCommissionCategory">
										<select name="dbxCategory" id="dbxCategory" class="mandatory" >
										<%
											StrSQL = "SELECT OPPDRAG_CATEGORY.CategoryID, OPPDRAG_CATEGORY.Name FROM OPPDRAG_CATEGORY LEFT OUTER JOIN " & _
												"TJENESTEOMRADE ON OPPDRAG_CATEGORY.TomID = TJENESTEOMRADE.TomID " & _
												"WHERE     (TJENESTEOMRADE.TomID =" &  itomID & ")"
												set rsCategory = GetFirehoseRS(StrSQL, Conn)
												
											if(clng(icategoryID) = 0) then
												Response.Write "<option VALUE='0' SELECTED >Choose -></option>"
											else
												Response.Write "<option VALUE='0' >Choose -></option>"
											end if
											
											if (HasRows(rsCategory) = true) then
												while not (rsCategory.EOF)
													If rsCategory("CategoryID") = icategoryID Then
														sel = " SELECTED"
													Else
														sel = ""
													End If
													Response.Write "<option VALUE='" & rsCategory("CategoryID") & "' " & sel & " >" & rsCategory("Name") & "</option>"
													rsCategory.MoveNext
												wend
											end if
												rsCategory.Close
												Set rsCategory = Nothing
										%>
										</select>
										</div>
									</td>
									</tr>
									<tr>
										<th>Ansvarlig:</th>
										<td>
											<table id="tblResponsible" cellspacing="0">
												<tr>
													<td>
														<div id="divResponsible">
											<select name="dbxMedarbeider" id="dbxMedarbeider" onkeydown="typeAhead()">
												<%
     												Response.write GetCoWorkersAsOptionList(Clng(lAnsMedID))
												%>
											</select>
														</div>
																								</td>
													<td>
														<input id='chkShowAllRes' name='chkShowAllRes' type='checkbox' class='checkbox' Value='1' onClick="ShowAll(<%=Clng(lAnsMedID)%>);">
													</td>
													<td>
														Vis Alle
													</td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<th>Rammeavtale:</th>
										<td>
											<select name="dbxFAgreement" id="dbxFAgreement" onkeydown="typeAhead()" style="width: 205px" class="mandatory">
												<option value="-1">Velg -></option>
											
												<%
												if lFaId = 0 and bChooseSelected then
												%>
											<option value="0"> - DO NOT USE -</option>
												<%
     													response.write GetActiveAgreementList(0)
     												Elseif lFaId = 0  then
     												%>
     													<option value="0" selected > - DO NOT USE -</option>
     												<%
     													response.write GetActiveAgreementList(0)
     												Else
     												%>
     													<option value="0" > - DO NOT USE -</option>
     												<%
     													response.write GetActiveAgreementList(Clng(lFaId))
     												End If
												%>
											</select>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</div>				
				<div class="contentHead">
					<h2>Notat</h2>
				</div>
				<div class="content">
					<table width="96%" id="Table12">
						<col width="15%">
						<col width="80%">
						<tr>
							<td>Kommentar til ansvarlig:</td>
							<td><textarea id="lnk23"  name="tbxNotatAnsvarlig" rows="2" cols="70"><%=strNotatAnsvarlig%></textarea></td>
						</tr>
						<!--
						<tr>
							<td>Kommentar til økonomi:</td>
							<td>
							  <div>	
								<%
									'editor.Value = strNotatOkonomi
									'editor.Create "txtEditorEconomy"													
							    %>
							  </div>							  
							</td>
						</tr>
						-->
						<tr>
							<td>Kommentar til økonomi:</td>
							<td><textarea id="lnk24" name="tbxNotatOkonomi" rows="2" cols="70" wrap="SOFT"><%=strNotatOkonomi%></textarea></td>
						</tr>						
						<tr>
							<td>Fakturakommentar :</td>
							<td><textarea id="invoiceComments" name="tbxInvoiceComments" rows="2" cols="70" wrap="SOFT"><%=strInvoiceComments%></textarea></td>
						</tr>
					</table>
					
					<br/><br/>
				</div>
				</form>
			</div>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>