<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim strSQL
	dim lngBrukerID
	dim rsHotlist
	dim rsOppdragStatus
	dim oppdragVikarStatus
        

	dim strTimeloenn
	dim strFACode
	dim lngFirmaID
	dim SOCuID
	dim SOPeID
	dim lFaId
	dim strKontaktperson
	dim rsOppdragVikar
	dim conIMP
	dim rsUserVikar
	dim iIMPId
	dim Conn
	dim hasConsultants : hasConsultants = false
	dim questbackEnabled
	dim QuestbackContackID

	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType
	dim blnShowHotList

	blnShowHotList = false
	strAddToHotlistLink = ""
	strHotlistType = "oppdrag"

	lngBrukerID = Session("BrukerID")

	' Move parameter OppdragsID to variable
	If lenb(Request("OppdragID")) > 0 Then
	   lngOppdragID = CLng(Request("OppdragID"))
	Else
		AddErrorMessage("Feil:Parameter for oppdragid mangler!")
		call RenderErrorMessage()
	End If

	' Move parameter FirmaID to variable
	If Request.QueryString("FirmaID") <> "" Then
	   lngFirmaID = Request.QueryString("FirmaID")
	Else
	   lngFirmaID = ""
	End If

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	'Determine if there are any consultants attached to this task
	strSQL = "SELECT DISTINCT COUNT([vikarID]) AS [AntallVikarer] FROM [Oppdrag_Vikar] WHERE [OppdragID] = " & lngOppdragID
	set rsOppdragVikar = GetFirehoseRS(strSQL, Conn)
	
	if (rsOppdragVikar("AntallVikarer") > 0) then
		hasConsultants = true
	end if
	rsOppdragVikar.close
	set rsOppdragVikar = nothing

	' Get Oppdrag Information ?
	strSql = "SELECT O.OppdragID, O.StatusID, O.FirmaID, O.ArbAdresse,  " &_
		"O.FraDato, O.FraKl, O.TilDato, O. TilKl, O.Beskrivelse, O.Oppdragskode, " &_
		"O.Bestiltdato, O.bestiltklokken, O.timepris, O.timerprdag, O.Timeloenn, O.deltagere, " &_
		"O.notatansvarlig, O.notatokonomi, O.Lunch, O.SOPeID, O.FaID, O.ReportingOverrule, O.ReportingContactID, " &_
		"F.Firma, Kontaktperson = K.Fornavn + ' ' + K.Etternavn, F.SOCuID, " &_
		"Ansvarlig = M.Fornavn + ' ' +M.Etternavn, " &_
		"S.Oppdragsstatus, A.navn Avdelingskontor, R.avdeling Regnskapskontor, " &_
		"T.navn Tjenesteomrade, O.notatkunde, " &_
		"O.webPub, O.webSted, O.webBeskrivelse, O.WebOverskrift, FA.FaCode ,FA.FaName " &_
		"FROM Oppdrag O, FIRMA F, KONTAKT K, MEDARBEIDER M, Avdeling R, " &_
		"H_OPPDRAG_STATUS S, tjenesteomrade T, avdelingskontor A,FrameworkAgreement FA " &_
		"WHERE OppdragID = " & lngOppdragID & " AND " & _
		"O.FirmaID *= F.FirmaID AND " &_
		"O.Bestilltav *= K.KontaktID and " &_
		"O.StatusID *= S.OppdragsStatusID AND " &_
		"O.FaID *= FA.FaID AND " &_
		"O.AvdelingID = R.AvdelingID AND " &_
		"A.ID =* O.AvdelingskontorID AND " &_
		"T.tomID =* O.tomID AND " &_
		"O.AnsMedID *= M.MedID "

	set rsOppdrag = GetFirehoseRS(StrSQL, Conn)

	' Move FROM recordset to variables
	if(not isnull(rsOppdrag("ReportingContactID"))) then
		QuestbackContackID = clng(rsOppdrag("ReportingContactID"))
	elseif (not isnull(rsOppdrag("SOPeID"))) then
		QuestbackContackID = clng(rsOppdrag("SOPeID")) 
	end if
	
	if(QuestbackContackID > 0 ) then
		Dim objSO 
		dim sopersonRs 		
		Set objSO = server.CreateObject("Integration.SuperOffice")

	set sopersonRs = objSO.GetPersonSnapshotById(QuestbackContackID)
		if (not sopersonRs.EOF) then
			if (isnull(sopersonRs("middlename"))) then
				strQuestbackContact  = sopersonRs("firstname") & " " & sopersonRs("lastname")
			else
				strQuestbackContact  = sopersonRs("firstname") & " " & sopersonRs("middlename") & " " & sopersonRs("lastname")
			end if
		end if
		set sopersonRs = nothing	
	end if

	questbackEnabled	= rsOppdrag("ReportingOverrule")
	lngOppdragID		= rsOppdrag("OppdragID")
	lStatusID			= rsOppdrag("StatusID")
	lngFirmaID			= rsOppdrag("FirmaID")
	SOCuID				= rsOppdrag("SOCuID")
	SOPeID				= rsOppdrag("SOPeID")
	strFirma			= rsOppdrag("Firma")
	strKontaktperson	= rsOppdrag("Kontaktperson")
	strFACode		= rsOppdrag("FaCode")
	strFaName		= rsOppdrag("FaName")
	lFaId			= rsOppdrag("FaID")
	
	if(questbackEnabled = false) then
		if(not isnull(SOCuID)) then
			strQuestbackEnabled = "Ja"
		else
			strQuestbackEnabled = "-"
		end if
	else
		strQuestbackEnabled = "Nei"
	end if
	if(isnull(strKontaktperson) and not isnull(SOPeID)) then
		Dim cts 
		dim personRs 		
		Set cts = server.CreateObject("Integration.SuperOffice")

	set personRs = cts.GetPersonSnapshotById(clng(SOPeID))
		if (not personRs.EOF) then
			if (isnull(personRs("middlename"))) then
				strKontaktperson = personRs("firstname") & " " & personRs("lastname")
			else
				strKontaktperson = personRs("firstname") & " " & personRs("middlename") & " " & personRs("lastname")
			end if
		end if
		set personRs = nothing
		set cts = nothing	
	end if
	strAnsvarlig		= rsOppdrag("Ansvarlig")
	strStatus			= rsOppdrag("OppdragsStatus")
	strAvdeling			= rsOppdrag("Avdelingskontor")
	strRegnskap			= rsOppdrag("Regnskapskontor")
	strTom				= rsOppdrag("Tjenesteomrade")
	strFraDato			= rsOppdrag("FraDato")
	strTilDato			= rsOppdrag("TilDato")
	strFraKl			= FormatDateTime( rsOppdrag("FraKl"), 4 )
	strTilKl			= FormatDateTime( rsOppdrag("TilKl"), 4 )
	strTimerprdag		= rsOppdrag("Timerprdag")
	strBeskrivelse		= rsOppdrag("Beskrivelse")
	strBestiltdato		= rsOppdrag("BestiltDato")
	strBestiltkl		= FormatDateTime( rsOppdrag("Bestiltklokken") , 4)
	strAdresse			= rsOppdrag("ArbAdresse")
	strNotatAnsvarlig	= rsOppdrag("NotatAnsvarlig")
	strNotatOkonomi		= rsOppdrag("NotatOkonomi")
	strNotatKunde		= rsOppdrag("NotatKunde")
	strTimepris			= rsOppdrag("Timepris")
	strTimeloenn		= rsOppdrag("Timeloenn")
	lunch				= FormatDateTime( rsOppdrag("Lunch") , 4)
	lOppdragskode		= CLng( rsOppdrag("Oppdragskode") )

	'Webrelaterte felter for Ekstern Web site
	webPub = cint(rsOppdrag("webPub"))
	strSted = rsOppdrag("webSted")
	strNotatWeb = rsOppdrag("webBeskrivelse")
	strWebOverskrift = rsOppdrag("WebOverskrift")

	' Relaese recordset
	rsOppdrag.Close
	Set rsOppdrag = Nothing

	' Is this KURSOPPDRAG ?
	If lOppdragskode = 1 Then

		' Get oppdrags kursdata
		strSQL = "SELECT O.kurskode, O.deltagere, Program=KO.KTittel, Kompniva=KL.KLevel, D.Dokumentasjon, T.KursType "&_
				" FROM OPPDRAG O, H_KOMP_TITTEL KO, H_KOMP_LEVEL KL," &_
				"H_KURS_DOK D, H_KURS_TYPE T" &_
				" WHERE OppdragID = " & lngOppdragID & " AND " & _
				"O.ProgramID *= KO.K_TittelID and KO.K_TypeID = 3 AND " &_
				"O.Kompniva *= KL.K_LevelID AND " &_
				"O.DokID *= D.OppdragdokID AND " &_
				"O.TypeID *= T.KurstypeID "

		Set rsOppdragKurs = GetFirehoseRS(strSQL, Conn)

		strdeltagere		= rsOppdragKurs("Deltagere")
		strProgram			= rsOppdragKurs("Program")
		strKompNiva			= rsOppdragKurs("KompNiva")
		strDokumentasjon	= rsOppdragKurs("Dokumentasjon")
		strType				= rsOppdragKurs("KursType")
		strKurskode			= rsOppdragKurs("Kurskode")

		' set radiobutton on kurskode
		If strKurskode = "1" Then
		   strDagskurs = "CHECKED"
		ElseIf  strKurskode  = "2" Then
		   strKveldskurs = "CHECKED"
		End If

		' Close Oppdrag
		rsOppdragKurs.Close
		Set rsOppdragKurs = Nothing
	End If

	' Get no of VIKARER who is ready to create Timelister
	strSQL = "SELECT NoOfOppdragVikar = COUNT( oppdragvikarid ) FROM Oppdrag_Vikar WHERE OppdragID = " & lngOppdragID & " AND StatusID = 4 AND Timeliste = 0"
	Set rsNoOfOppdragVikar = GetFirehoseRS(strSQL, Conn)
	lNoOfOppdragVikar = CLng( rsNoOfOppdragVikar( "NoOfOppdragVikar" ) )

	' Close Oppdrag
	rsNoOfOppdragVikar.Close
	Set rsNoOfOppdragVikar = Nothing

	' Title and heading in page
	strHeading = "Viser oppdrag for " & strFirma

	If Request.QueryString("slett")="ja" Then
		strID = Request.QueryString("ID")
		strSQL = "DELETE FROM OPPDRAG_KOMPETANSE WHERE k_oppdrID = " & strID
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Feil oppstod under sletting av kompetanse.")
			call RenderErrorMessage()
		End if
		Response.redirect "OppdragVis.asp?OppdragID=" & lngOppdragID
	end if

	strSQL = "SELECT * FROM HOTLIST WHERE status= 1 And BrukerID=" & lngBrukerID & " And oppdragID = " & lngOppdragID
	Set rsHotlist = GetFirehoseRS(strSQL, Conn)

	if (HasRows(rsHotlist) = false) then
		blnShowHotList = true
		strAddToHotlistLink = "AddHotlist.asp?kode=1&oppdragID=" & lngOppdragID & "&kundeNavn=" & server.URLEncode(strFirma) & "&kundeNr=" & lngFirmaID
		strHotlistType = "oppdrag"
	end if
	rsHotlist.close
	set rsHotlist = nothing
	
	'''STH bug 
	strSQL = "select  count(*) as count , max(case isnull(statusid,0) " & _ 
		 "when 4 then 1 " & _
		 "else 0 end) as status " & _
		 "from OPPDRAG_VIKAR where oppdragid=" & lngOppdragID

		 
	set rsOppdragStatus = GetFirehoseRS(strSQL, Conn)
	
	if (cint(rsOppdragStatus("count"))=0) then
		oppdragVikarStatus = 0	
	else
		oppdragVikarStatus = cint(rsOppdragStatus("status"))
	end if


	'-- Check Add has been already saved 
	dim sSql
	dim Conn2
	dim rsComAdd
	dim newOrUpdateAdvert
	dim status
    dim unpublish
	dim stepstoneIcoUrl
	dim xtraIcoUrl
	
	unpublish="NO"
	set Conn2 = GetConnection(GetConnectionstring(XIS, ""))
	
    sSql="SELECT     Count(CommissionId) as Recs   FROM        CommissionAds  WHERE     (CommissionId = " &  lngOppdragID  & ")"
	set rsComAdd = GetFirehoseRS(sSql, Conn2)
	
	
	
	if (rsComAdd("Recs") = 0 ) then
	   newOrUpdateAdvert = "new" 
	   status = "Ny"
    else
      newOrUpdateAdvert = "update" 
      status = "Oppdatèr"
    end If
    
    rsComAdd.Close
    set rsComAdd = Nothing
    
    
    
	set rsComAdd = GetFirehoseRS(sSql, Conn2)

    
    '-- End Check Add has been already saved 
    
    '-- Check the publisher's publish status
    
     sSql="SELECT    PublishedStatus,PubId FROM        CommPublishers  WHERE     (CommId = " &  lngOppdragID  & " AND PublishedStatus != 'U')"
     xtraIcoUrl =  "images/Un_published.jpg" 
     stepstoneIcoUrl = "images/Un_published.jpg" 
     finnNoIcoUrl = "images/Un_published.jpg" 
 
     set rsComAdd = GetFirehoseRS(sSql, Conn2)
     
 
     WHILE ( not rsComAdd.EOF)
           unpublish="YES"

           if(rsComAdd("PubId") =0) then
             
              If (UCase(rsComAdd("PublishedStatus")) = "P" Or UCase(rsComAdd("PublishedStatus")) = "R") then
                  xtraIcoUrl =  "/xtra/images/published.jpg" 
                  xtraAction = "unpublish"
              End If
             
           elseif (rsComAdd("PubId") =1) then
                  If (UCase(rsComAdd("PublishedStatus")) = "P" Or UCase(rsComAdd("PublishedStatus")) = "R") then
                    stepstoneIcoUrl=  "/xtra/images/published.jpg" 
                    stepStoneAction="unpublish"
                 End If
             
            elseif (rsComAdd("PubId") =2) then
                  If (UCase(rsComAdd("PublishedStatus")) = "P" Or UCase(rsComAdd("PublishedStatus")) = "R") then
                    finnNoIcoUrl=  "/xtra/images/published.jpg" 
                    finnNoAction="unpublish"
                  End If
           
           
          end if
           
              
           rsComAdd.MoveNext
     WEND
     
    rsComAdd.Close
    set rsComAdd = Nothing

    ' End Check the publisher's publish status

   

	'Top Menu init
	dim strALinkStart
	dim strALinkEnd
	
      
               

	dim AToolbarAtrib(9,3) '0 = Enable/Disable, 1 = from/Link to activate, 2 = close link form string

	'Lagre Oppdrag
	AToolbarAtrib(0,0) = "0"
	AToolbarAtrib(0,1) = ""
	AToolbarAtrib(0,2) = ""
	'Vise oppdrag
	AToolbarAtrib(1,0) = "0"
	AToolbarAtrib(1,1) = ""
	AToolbarAtrib(1,2) = ""
	'Til Endre oppdrag
	AToolbarAtrib(2,0) = "1"
	AToolbarAtrib(2,1) = "<input NAME='pbnDataAction' TYPE='hidden' VALUE='Endre Oppdrag'>" & _
	"<a href='javascript:document.all.formJobChange.submit();' title='Endre oppdrag'>"
	AToolbarAtrib(2,2) = "</a>"
	'Tilknytte vikar
		AToolbarAtrib(3,0) = "1"
		AToolbarAtrib(3,1) =  "<form action='VikarSoek.asp?kurskode=" & strKurskode & "' METHOD='POST'  name='frmAddConsultant'>" & _
		"<input name='hdnPosted' TYPE='hidden' VALUE='1'>" & _
		"<input name='tbxOppdragID' TYPE='hidden' VALUE='" & lngOppdragID & "'>" & _
		"<input name='tbxFirmaID' TYPE='hidden' VALUE='" & lngFirmaID & "'>" & _
		"<a href='javascript:document.all.frmAddConsultant.submit();' title='Tilknytt vikar'>"
		AToolbarAtrib(3,2) = "</a></form>"
	'Aktiviteter for oppdrag
	AToolbarAtrib(4,0) = "1"
	AToolbarAtrib(4,1) = "<a href='AktivitetOppdrag.asp?OppdragID=" & lngOppdragID & "' title='Vis aktiviteter på oppdraget'>"
	AToolbarAtrib(4,2) = "</a>"
	'Kalender
	AToolbarAtrib(5,0) = "1"
	AToolbarAtrib(5,1) = "<form action='Kalender.asp?OppdragID=" & lngOppdragID & "' METHOD='POST' name='frmJobCalendar'>" & _
	"<a href='javascript:document.all.frmJobCalendar.submit();' title='Vis kalender for oppdrag'>"
	AToolbarAtrib(5,2) = "</a></form>"

	if(lStatusID <>1) then
	  	  AToolbarAtrib(6,0) = "0"
    else
	  AToolbarAtrib(6,0) = "1"
	end if  
	AToolbarAtrib(6,1) = "<a href='WebUI\WizardAP.aspx?CommId=" & lngOppdragID & "&unpublish=" & "" & "' title='" & status  &  "'>"
        
                
  
	AToolbarAtrib(6,2) = "</a>"
    
  	AToolbarAtrib(7,0) = "1"
	AToolbarAtrib(7,1) = "<a href='WebUI\WizardAP.aspx?CommId=" & lngOppdragID & "&unpublish=YES ' title='" &  "Fjern ann"  &  "'>"

      
        

	AToolbarAtrib(7,2) = "</a>"

	AToolbarAtrib(8,0) = "1"
	AToolbarAtrib(8,1) = "<input NAME='hdnCopyCommission' TYPE='hidden' VALUE='Kopier oppdrag'>" & _
	"<a href='javascript:document.all.frmCopyCommission.submit();' title='Kopier oppdrag'>"
	AToolbarAtrib(8,2) = "</a>"
	
	
	AToolbarAtrib(9,0) = "1"
	AToolbarAtrib(9,1) = "<a target='_blank' href='WebUI\PreviewAdd.aspx?CommId=" & lngOppdragID & "' title='" & "Forhåndsvis annonse"  &  "'>"
	AToolbarAtrib(9,2) = "</a>"
	

	rsOppdragStatus.Close
	set rsOppdragStatus = Nothing
	
	
	

    
    %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><% =strHeading %></title>
		<script type="text/javascript" src="/xtra/Js/javascript.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>
                <script language='javascript' src='/xtra/js/advertPublishFunction.js' id='advertNavigationScripts'></script>
 
		<script language="javaScript" type="text/javascript">
			function shortKey(e)
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				if (modKey && modKey2 && keyChar=="S")
				{
					parent.frames[funcFrameIndex].location=("/xtra/OppdragSoek.asp");
				}
			}
			//her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	<body onLoad="fokus();">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
				<!--#include file="Includes/Top_Menu_job.asp"-->
			</div>
			<form name="formJobChange" ACTION="OppdragNy.asp?OppdragID=<%=lngOppdragID%>" METHOD="POST" ID="Form1">
				</form>
				<form name="frmCopyCommission" action="OppdragNy.asp?OppdragID=<%=lngOppdragID%>&OppdragAct=copy&OVID=" method="post" id="formcc">								
				</form>					
				<div class="content">
					<input NAME="tbxOppdragID" TYPE="HIDDEN" VALUE="<%=lngOppdragID%>" ID="Hidden1">
					<table width="96%" ID="Table1">
						<col width="33%">
						<col width="20%">
						<col width="25%">
						<col width="23%"
						<tr>
							<td>
								<table ID="Table2">
									<tr>
										<th>Oppdragnr:</th>
										<td><%=lngOppdragID%></td>
									</tr>
									<tr>
										<th>Beskrivelse:</th>
										<td><%=strBeskrivelse%></td>
									</tr>
									<tr>
										<th>Arb.adresse:</th>
										<td><%=strAdresse%></td>
									</tr>
									<tr>
										<th>Dato:</th>
										<td><%=strBestiltdato%>, kl:<%=strBestiltkl%></td>
									</tr>
									<tr>
										<th>Status:</th>
										<td><%=strStatus %></td>
									</tr>
									<tr>
										<th>Dato:</th>
										<td><%=strFraDato%> - <%=strTilDato%></td>
									</tr>
									<tr>
										<th>Tidsrom:</th>
										<td><%=strFraKl%> - <%=strTilKl%> </td>
									</tr>
									<tr>
										<th>Lunsj:</th>
										<td><%=lunch %></td>
									</tr>
									<%
									If lOppdragskode = 1 Then
										%>
										<tr>
											<th>&nbsp;</th>
											<td>
												<input class="radio" NAME="rbnKurskode" TYPE="RADIO" VALUE="1" <%=strDagskurs%> ID="Radio1">Dagskurs
												<input class="radio" NAME="rbnKurskode" TYPE="RADIO" VALUE="2" <%=strKveldskurs%> ID="Radio2">KveldKurs
											</td>
										</tr>
										<%
									End If
									%>
									<tr>
										<th>Timer pr.dag:</th>
										<td><%=strTimerprdag %></td>
									</tr>
									<tr>
										<th>Timepris:</th>
										<td><% =strTimepris %></td>
									</tr>
									<tr>
										<th>Timelønn:</th>
										<td><% =strTimeloenn %></td>
									</tr>
									
								</table>
							</td>
							<td>
								<table ID="Table3">
									<tr>
										<th>Kontaktnr:</th>
										<td><%=lngFirmaID%></td>
									</tr>
									<tr>
										<th>Kontakt:</th>
										<td><%=CreateSONavigationLink(SUPEROFFICE_PANEL_CONTACT_URL, SUPEROFFICE_PANEL_CONTACT_URL, SOCuID, strFirma, "Vis Kontakt '" & strFirma & "'")%></a></td>
									</tr>
									<tr>
										<th>Kontaktperson:</th>
										<td><%=strKontaktperson%></td>
									</tr>									
								</table>
							</td>
							<td>
								<table ID="Table4">
									<tr>
										<th>Avdelingskontor:</th>
										<td><%=strAvdeling%></td>
									</tr>
									<tr>
										<th>Regnskapsavdeling:</th>
										<td><%=strRegnskap%></td>
									</tr>
									<tr>
										<th>Tjenesteomr&aring;de:</th>
										<td><%=strtom%></td>
									</tr>
									<tr>
										<th>Ansvarlig:</th>
										<td><%=strAnsvarlig%></td>
									</tr>
									<tr>
										<th>Rammeavtale:</th>
										<td><% =strFACode %> - <% =strFaName%></td>
									</tr>
								</table>
							</td>
							<td>
							<table ID="questbackTable">
								<tr>
										<th>Skal motta evaluering:</th>
										<td><%= strQuestbackEnabled %></td>
									</tr>
									<tr style="visibility:<% if(questbackEnabled = false and not isnull(SOCuID)) then response.write("visible") else response.write("hidden") end if %>">
										<th>Alternativ kontakt (evaluering):</th>
										<td><%= strQuestbackContact %></td>
									</tr>
							</table>
							</td>
						</tr>
					</table>
					<br>
					<br>
					<table cellpadding="0" cellspacing="0" border="0" ID="Table5">
						<%
						If lOppdragskode = 1 Then
							%>
							<tr>
								<th>Program:</th>
								<td><%=strProgram%></td>
							</tr>
							<tr>
								<th>Kursnivå:</th>
								<td><% =strKompNiva%></td>
							</tr>
							<tr>
								<th>Type:</th>
								<td><%=strType%></td>
							</tr>
							<tr>
								<th>Versjon:</th>
								<td><%=strDeltagere%></td>
							</tr>
							<tr>
								<th>Dokumentasjon:</th>
								<td><%=strDokumentasjon%></td>
							</tr>
							<%
						End If
						If  Trim( strNotatAnsvarlig ) <> "" Then
								Response.Write "<tr><th>Kommentar til ansvarlig:&nbsp</th>"
								Response.Write "<td>" & strNotatAnsvarlig & "</td></tr>"
						End If
						If  Trim( strNotatOkonomi ) <> "" Then
								Response.Write "<tr> <th>Kommentar til økonomi:&nbsp</th>"
								Response.Write "<td>" & strNotatOkonomi & "</td></tr>"
						End If
						%>
					</table>&nbsp;
				</div>
				<div class="contentHead" style="width: 963; height: 144"><h2>Oppdragpublisering på internett</h2>
				</div>
				 <div class="content">
                  <table cellspacing="1" width="50%" style="padding: 0" >
                    <tr>
                      <td width="30%" colspan="3">
                      .......................................</td>
                    </tr>
                    <tr>
                      <td width="12%">xtra.no </td>
                      <td width="4%">
                     
                      <img border="0" src="<%=xtraIcoUrl%>"></a></td>
                      <td width="38%">
                     
                      &nbsp;</td>
                    </tr>
                    <tr>
                       <td width="30%" colspan="3">
                      .......................................</td>
                    </tr>
                    <tr>
                      <td width="12%">StepStone.no</td>
                      <td width="4%">
                      <img border="0" src="<%=stepstoneIcoUrl%>"></a></td>
                      <td width="38%">
                      &nbsp;</td>
                    </tr>
                    <tr>
                      <td width="30%" colspan="3">
                      .......................................</td>
                    </tr>

                    <tr>
                      <td width="12%">Finn.no</td>
                      <td width="4%">
                      <img border="0" src="<%=finnNoIcoUrl%>"></a></td>
                      <td width="38%">
                      &nbsp;</td>
                    </tr>
                    <tr>
                      <td width="30%" colspan="3">
                      .......................................</td>
                    </tr>
                  </table>
                </div>
							
				<%				
				if (hasConsultants) then
					%>
					<div class="contentHead"><h2>Rettigheter på web</h2></div>
					<div class="content">
						<table ID="Table6">
							<%
							if lStatusID = "5" then
								dim objCons
								strSQL = "SELECT v.vikarid, v.fornavn, v.etternavn FROM Oppdrag o, Vikar v, Oppdrag_Vikar ov " &_
										"WHERE o.OppdragID = " & lngOppdragID & " and ov.OppdragID = o.OppdragID and v.VikarID = ov.VikarID "&_
										" and ov.statusid = 4"

								set rsVikar = GetFirehoseRS(StrSQL, Conn)

								if HasRows(rsVikar) then
									vikarId = rsVikar("VikarID")
									strVikar = rsVikar("Fornavn") & " " & rsVikar("Etternavn")
								end if
								rsVikar.close
								set rsVikar = nothing

								if trim(vikarId) <> "" then
									'Initialize ADO objects


									set objCons = Server.CreateObject("XtraWeb.Consultant")
									objCons.XtraConString = Application("Xtra_intern_ConnectionString")
									objCons.XtraDataShapeConString = Application("ConXtraShape")

									call objCons.GetConsultant(vikarId)

									set objRightsCons = ObjCons.GetWebRights
									objRightsCons.GetTaskRights(lngOppdragID)
									%>
									<tr>
										<td><strong>Rettigheter på web for vikar</strong><br>
											<%
											set objRight = objRightsCons.Item(2)
											%>
											<input type="checkbox" class="checkbox" disabled <%=objright.datavalues("checked").value%> value="1" id=checkbox1 name="consBox2"><%=objright.datavalues("intraNavn").Value%><br>
											<%
											set objRight = objRightsCons.Item(1)
											%>
											<input type="checkbox" class="checkbox" disabled <%=objright.datavalues("checked").value%> value="1" id="Checkbox2" name="consBox1"><%=objright.datavalues("intraNavn").Value%><br>
											<%
											if not lOppdragskode = 1 then
												set objRight = objRightsCons.Item(3)
												%>
												<input type="checkbox" class="checkbox" disabled <%=objright.datavalues("checked").value%> value="1" id="Checkbox3" name="consBox3"><%=objright.datavalues("intraNavn").Value%><br>
												<%
											else
												set objRight = objRightsCons.Item(4)
												%>
												<input type="checkbox" class="checkbox" disabled <%=objright.datavalues("checked").value%> value="1" id="Checkbox4" name="consBox4"><%=objright.datavalues("intraNavn").Value%><br>
												<%
											end if
											%>
										</td>
									</tr>
									<%
									set objRightsCons = nothing
									set objRightsCust = nothing
									objCons.cleanup
									set objCons = nothing
									set objCust = nothing
								end if
							end if
							%>
						</table>
					</form>
					<%
					if (lStatusID = "5") and (trim(vikarId) <> "" ) then
						%>
						<form ACTION="RettigheterWeb.asp" METHOD="POST" name="frmWebRights" ID="Form2">
							<input NAME="oppdragID" TYPE="HIDDEN" VALUE="<%=lngOppdragID%>" ID="Hidden2">
							<span class="menuInside" title="Rettigheter på web"><a href="#" onClick="javascript:document.all.frmWebRights.submit()">&nbsp;Rettigheter på web</a></span><br>&nbsp;
						</form>
						<%
					End If
					%>
				</div>
				<%
			end if
			' Get kompetansekrav
			strKompetanse = "SELECT K_OppdrId, KType, O.K_TypeID, KTittel, KLevel, Beskrivelse " &_
				"from OPPDRAG_KOMPETANSE O, H_KOMP_TITTEL T, H_KOMP_LEVEL L, H_KOMP_TYPE TY " &_
				"WHERE O.OppdragID = " & lngOppdragID  &_
				" and O.K_TypeID = T.K_TypeID " &_
				" and O.K_TittelID = T.K_TittelID " &_
				" and O.K_TypeID = TY.K_TypeID " &_
				" and O.K_LevelID *= L.K_LevelID " &_
				" order by O.K_TypeID, O.K_TittelID"

			set rsKompetanse = GetFirehoseRS(strKompetanse, Conn)
			if HasRows(rsKompetanse) then
				%>
				<table ID="Table7">
					<%
					Do Until rsKompetanse.EOF
						%>
						<tr>
							<th><%=rsKompetanse("KType")%>:</th>
							<td>
								<%
								If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) Then
									%>
									<a href="OppdragKomp.asp?OppdragID=<%=lngOppdragID %>&amp;TypeID=<% =rsKompetanse("K_TypeID") %>&amp;KompetanseID=<% =rsKompetanse("K_OppdrID") %>"><% =rsKompetanse("KTittel") %></a>
									<%
								Else
									%>
									<%=rsKompetanse("KTittel")%>
								<%
								End If
								%>
							</td>
							<th>Nivå:</th>
							<td><%=rsKompetanse("KLevel")%></td>
							<th>Kommentar:</th>
							<td><%=rsKompetanse("Beskrivelse")%></td>
							<td class="center"><a href="OppdragVis.asp?OppdragID=<%=lngOppdragID%>&amp;Slett=ja&amp;ID=<%=rsKompetanse("k_oppdrID")%>">fjern</a></td>
						</tr>
						<%
						rsKompetanse.MoveNext
					Loop
					%>
				</table>
				<%
			end if
			' Close Oppdrag
			rsKompetanse.Close
			Set rsKompetanse = Nothing

			if (hasConsultants) then
				'Get assigned consultants for tasks
				' Changed 26.11.2001 E.L
				strSQL = "SELECT " & _
					"OPPDRAG_VIKAR.Faktor, " & _
					"OPPDRAG_VIKAR.direkteTelefon, " & _
					"OPPDRAG_VIKAR.jobbEpost, " & _
					"OPPDRAG_VIKAR.OppdragVikarID, " & _
					"OPPDRAG_VIKAR.VikarID, " & _
					"OPPDRAG_VIKAR.Timeloenn, " & _
					"OPPDRAG_VIKAR.Timepris, " & _
					"OPPDRAG_VIKAR.Notat, " & _
					"OPPDRAG_VIKAR.Fradato, " & _
					"OPPDRAG_VIKAR.Tildato, " & _
					"OPPDRAG_VIKAR.Timeliste, " & _
					"OPPDRAG_VIKAR.StatusID, " & _
					"OPPDRAG_VIKAR.CategoryID, " & _
					"FrameworkCategory.CategoryCode, " & _
					"VIKAR_ANSATTNUMMER.ansattnummer, " & _
					"VIKAR.Fornavn, " & _
					"VIKAR.Etternavn, " & _
					"H_OPPDRAG_VIKAR_STATUS.Status " & _
					"FROM OPPDRAG_VIKAR " & _
					"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _
					"LEFT OUTER JOIN FrameworkCategory ON OPPDRAG_VIKAR.CategoryID = FrameworkCategory.CategoryID " & _
					"LEFT OUTER JOIN H_OPPDRAG_VIKAR_STATUS ON OPPDRAG_VIKAR.Statusid = H_OPPDRAG_VIKAR_STATUS.OppdragVikarStatusID " & _
					"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
					"WHERE OPPDRAG_VIKAR.Oppdragid = '" & lngOppdragID & "' " & _
					"ORDER BY fradato DESC, OPPDRAG_VIKAR.Statusid DESC "

				set rsOppdragVikar = GetFirehoseRS(strSQL, Conn)
				if not rsOppdragVikar.EOF then
					%>
					<div class="contentHead">
						<h2>Tilknyttede vikarer</h2>
					</div>
					<div class="content">

					<%
					' Status "Bekreftet" ?
					If  lNoOfOppdragVikar > 0 And HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true Then
						' Add function to create Timeliste
						%>
							<br>
							<form ACTION='OppdragDB.asp' NAME="frmTimeSheet" METHOD="POST" ID="Form3">
								<INPUT NAME="OppdragID" TYPE="HIDDEN" VALUE="<%=lngOppdragID%>" ID="Hidden3">
								<INPUT NAME="frakode" TYPE="HIDDEN" VALUE="1" ID="Hidden4">
								<INPUT NAME="Kurskode" TYPE="HIDDEN" VALUE="<%=strDagskurs%>" ID="Hidden5">
								<INPUT NAME="Kurs" TYPE="hidden" VALUE="<%=strKurskode%>" ID="Hidden6">
								<INPUT NAME="hdnJobAction" TYPE="hidden" VALUE='Lag Timeliste' ID="Hidden7">
							</form>
							<span class="menuInside" title="Lag Timeliste"><a href="#" onClick="javascript:document.all.frmTimeSheet.submit()">&nbsp;Lag Timeliste</a></span><br><br>
						<%
					End If
					%>
					<div class="listing">
						<table cellpadding="0" cellspacing="1" width="97%" ID="Table8">
							<tr>
								<th>Vikar</th>
								<th>Fradato</th>
								<th>Tildato</th>
								<th>Status</th>
								<th>Timelønn</th>
								<th>Timepris</th>
								<th colspan="3">Timeliste</th>
								<th Colspan="2">Bekreftelser</th>
								<th class="center">Akt</th>
								<%
								if (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) then
									%>
									<th class="center">Slett</th>
									<th class="center">Kopier</th>
									<%
								end if
								%>
							</tr>
							<%
							Do Until rsOppdragVikar.EOF
								strName = rsOppdragVikar("Fornavn") & " " & rsOppdragVikar("Etternavn")
								%>
								<tr>
									<td>
										<a href="OppdragVikarNy.asp?OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID") %>&amp;FaId=<%=lFaId%>"><img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="Redigere kontaktinformajon" align="absmiddle"></a>
										<%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & rsOppdragVikar( "VikarID" ), strName, "Vis vikar " & strName )%><br />
										<%
										if len(rsOppdragVikar("jobbepost"))> 0 then
											%>
											<a href="mailto:<% =rsOppdragVikar("jobbepost")%>"><% =rsOppdragVikar("jobbepost")%></a>&nbsp;
											<%
										end if
										if len(rsOppdragVikar("direkteTelefon"))> 0 then
											response.write "tlf:" & rsOppdragVikar("direkteTelefon")
										end if
										
										if len(rsOppdragVikar("CategoryId"))> 0 then
											response.write " Kat:" & rsOppdragVikar("CategoryCode")
										end if
										%>
									</td>
									<td><% =rsOppdragVikar("Fradato")%>&nbsp;</td>
									<td><% =rsOppdragVikar("Tildato")%>&nbsp;</td>
									<%
									If  HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true Then 'adminrettighet
										%>
											<td><a Href="oppdragVikarny.asp?OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID") %>&amp;FaId=<%=lFaId%>"><% =rsOppdragVikar("status") %></a>&nbsp;</td>
										<%
									Else
										%>
										<td><% =rsOppdragVikar("status") %>&nbsp;</td>
										<%
									End If
									%>
									<td><% =rsOppdragVikar("Timeloenn") %>&nbsp;</td>
									<td><% =rsOppdragVikar("Timepris") %>&nbsp;</td>
									<%
									If rsOppdragVikar("statusID") = 4 and rsOppdragVikar("Timeliste") = 1 Then 'aksept og timeliste er laget
										%>
										<td><a Href="datafiler/Vikar_timeliste_vis3.asp?VikarID=<%=rsOppdragVikar("VikarID") %>&amp;OppdragID=<% =lngOppdragID %>&amp;FirmaID=<% =lngFirmaID %>&amp;frakode=1">timeliste</a>&nbsp;</td>
										<td><a Href="OppdragVikarNy.asp?OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID")%>&amp;FaId=<%=lFaId%>&amp;AKSJON=UTVID">Utvid</a>&nbsp;</td>
										<td><a Href="OppdragVikarNy.asp?OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID")%>&amp;FaId=<%=lFaId%>&amp;AKSJON=FORKORT">Forkort</a>&nbsp;</td>
										<%
										If lOppdragskode = 0 Then 'ikke kursoppdrag
											%>
											<td><a Href="Oppdrag-bekr-kunde.asp?OppdragID=<%=lngOppdragID%>&amp;OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID")%>">Kontakt</a>&nbsp;</td>
											<td><a Href="Oppdrag-bekr-Vikar.asp?OppdragID=<%=lngOppdragID%>&amp;OppdragVikarID=<%=rsOppdragVikar("OppdragVikarID")%>">Vikar</a>&nbsp;</td>
											<%
										Else
											%>
											<td>&nbsp;</td>
											<td>&nbsp;</td>
											<%
										End If 'ikke kursoppdrag
									Else
										%>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<%
									End If 'akksept og timeliste er laget
									%>
									<td class="center"><a Href="AktivitetOppdrag.asp?VikarID=<%=rsOppdragVikar("VikarID") %>&amp;OppdragID=<% =lngOppdragID %>&amp;FirmaID=<% =lngFirmaID %>&amp;frakode=1"><img src="Images/icon_activities.gif" alt="Aktivitet på oppdrag og vikar"></a></td>
									<%
									if (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) then
										param = "OppdragVikarSlettTimeliste.asp" &_
											"?fradato=" & rsOppdragVikar("Fradato") &_
											"&tildato=" & rsOppdragVikar("Tildato") &_
											"&VikarID=" & rsOppdragVikar("VikarID") &_
											"&OppdragID=" & lngOppdragID &_
											"&OppdragvikarID=" & rsOppdragVikar("OppdragVikarID")
										ccparam = "OppdragNy.asp?OppdragID=" & lngOppdragID & "&OppdragAct=copy&vikarID=" & rsOppdragVikar("VikarID")& "&OVID=" & rsOppdragVikar("OppdragVikarID") '"javascript:document.all.frmCopyCommissionVikar.submit();"
										%>
										<form name="frmCopyCommissionVikar" action="OppdragNy.asp?OppdragID=<%=lngOppdragID%>&OppdragAct=copy&vikarID=<%=rsOppdragVikar("VikarID")%>" method="post" id="formcc2">										
										</form>									
										<td class="center"><a href="<% =param %>"><img src="Images/icon_delete.gif" alt="Slette"></a></td>
										<% if rsOppdragVikar("statusID") <> 4 then %>										
											<td class="center"><a href="<%=ccparam %>"><img src="Images/icon_copyCommission.gif" alt="Flytt og tilknytt til kopiert oppdrag"></a></td>
										<% else %>
											<td class="center">&nbsp;</td>
										<% end if 
									end if
									%>
								</tr>
								<%
								if (len(rsOppdragVikar("Notat")) > 0) then
									%>
									<tr>
										<td colspan="14"><% =rsOppdragVikar("Notat") %>&nbsp;</td>
									</tr>
									<%
								end if
								rsOppdragVikar.MoveNext
							Loop
						end if
						rsOppdragVikar.Close: Set rsOppdragVikar = Nothing
						%>
						</table>
					</div>
				</div>
				<%
			end if
			%>
		</div>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>