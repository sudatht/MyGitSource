<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	brukerID = Session("BrukerID")

	' Move parameters to local variables
	lVikarID = Request.Querystring("suspectID")
	strRegistrertVikar = request.querystring("vikarID")

	if strRegistrertVikar = "" then
		dim exp
		dim cvfil				'as string
		dim ObjProdgroup		'as xtraweb.productgroup
		dim rsQualifications	'as adodb.recordset
		dim	ObjJobgroup			'xtraweb.jobgroup
		dim RsJobWishes 		'as adodb.recordset
		dim	StrJobGroup 		'as string
		dim fulltime
		dim beskrivelse
		dim interestedJobs
		dim strHTTPAdress
		dim visJobgr
		dim forerkort
		dim bildisp
		dim oppsigelse
		dim strOppsigelse
		
		dim susp
		dim adr
		dim jobApp
		dim jobPl
		dim edu
		dim ref
		dim uplFile


		set susp	= Server.CreateObject("XtraWeb.Suspect")
		susp.XtraConString = Application("Xtra_intern_ConnectionString")
		susp.XtraDataShapeConString = Application("ConXtraShape")
		set adr		= Server.CreateObject("XtraWeb.Addresses")
		set jobApp	= Server.CreateObject("XtraWeb.JobApplication")
		set jobPl	= Server.CreateObject("XtraWeb.JobPlace")
		set edu		= Server.CreateObject("XtraWeb.Education")
		set exp		= Server.CreateObject("XtraWeb.Experience")
		set ref		= Server.CreateObject("XtraWeb.Reference")
		set uplFile = Server.CreateObject("XtraWeb.UploadFile")

		' Check VikarID
		If lVikarID = "" Then
			AddErrorMessage("Error: SuspectId mangler!")
			call RenderErrorMessage()
		End If

		if not susp.GetSuspect(lVikarID) then
			AddErrorMessage("Error: Suspect ikke funnet!")
			call RenderErrorMessage()		
		end if

		if susp.DataValues("overfort") then
			AddErrorMessage("Suspect allerede overført!")
			call RenderErrorMessage()		
		end if

		if susp.DataValues("foererkort") = "1" then
			forerkort = true
		else 
			forerkort = false
		end if

		if cbool(susp.DataValues("bil")) then
			bildisp = true
		else
			 bildisp = false
		end if
		oppsigelse = susp.DataValues("oppsigelsestid")

		if oppsigelse = 0 then
			strOppsigelse = "Ingen"
			elseif oppsigelse = 1 then strOppsigelse = "14 dager"
			elseif oppsigelse = 2 then strOppsigelse = "1 måned"
			elseif oppsigelse = 3 then strOppsigelse = "2 måneder"
			elseif oppsigelse = 4 then strOppsigelse = "3 måneder"
			elseif oppsigelse = 5 then strOppsigelse = "over 3 måneder"
		end if

		set jobApp		= susp.JobApplication
		jobApp.Refresh()
		set jobPl		= jobApp.JobPlaces
		set CV			= susp.CV
		CV.XtraConString = Application("Xtra_intern_ConnectionString")
		CV.XtraDataShapeConString = Application("ConXtraShape")
		CV.refresh
		
		'Creates a DB connection to retrieve the CV record
		Set cvConn = GetClientConnection(GetConnectionstring(XIS, ""))	
		set rsCV = GetFirehoseRS("SELECT max(CVId) AS CVId FROM CV WHERE ConsultantId =" & lVikarID, cvConn)	
		vikarcvid = rsCV("CVId")	
		rsCV.Close		
	
		'Creates a DB connection to retrieve the course list
    	Set listConn = GetClientConnection(GetConnectionstring(XIS, ""))
		Set rsCourseList	= GetFirehoseRS("SELECT DataId,CVId,Place,Country,Description,Title,FieldType,FromMonth,FromYear,ToMonth,ToYear FROM CV_Data WHERE FieldType = 'COU' and CVId =" & vikarcvid & " ORDER BY FromYear desc, FromMonth desc", listConn)	

		set allJobgr	= cv.JobGroups

		if jobApp.DataValues.Count > 1 then
			beskrivelse = jobApp.DataValues("Description")
			interestedJobs = jobApp.DataValues("InterestedJobs")
		end if

		set allEdu = cv.Educations
		set allExp = cv.Experiences
		set allRef = cv.References
		set allFile = cv.UplaodFiles 'SKA

		visJobgr = false

		if cv.Datavalues.Count > 1 then
			if not isNull(cv.DataValues("Filename")) then
				cvfil = cv.DataValues("Filename")
			else
				cvfil = ""
			end if
		else
			cvfil = ""
		end if

	end if

	' Open database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))	

	strSQL = "SELECT V.regdato, V.Etternavn,V.Fornavn,V.foedselsdato,V.WorkType,  " &_
			"A.Adresse, A.Postnr, A.PostSted, V.AnsMedID,"&_
			"V.Telefon, V.MobilTlf, V.Fax, V.Epost, V.Hjemmeside, V.personsoek," &_
			"V.notat,Link1URL,Link2URL,Link3URL " &_
			"FROM V_SUSPECT V, V_SUSPECT_ADRESSE A" &_
			" where V.suspectID = " & lVikarID  &_
			" and V.suspectID = A.adresseRelID" & _
			" and A.AdresseType = 2"
			
	set rsVikar = GetFirehoseRS(strSQL, Conn)

	strRegDato		= rsVikar("regdato")
	strEtternavn	= rsVikar("Etternavn")
	strFornavn		= rsVikar("Fornavn")
	strFoedselsdato	= rsVikar("Foedselsdato")
	strNotat		= rsVikar("Notat")
	strTelefon		= rsVikar("Telefon")
	strPersonsoek	= rsVikar("Personsoek")
	strFax			= rsVikar("Fax")
	strMobilTlf		= rsVikar("MobilTlf")
	strEPost		= rsVikar("EPost")
	strAnsMed		= rsVikar("AnsMedID")
	strAdresse		= rsVikar("Adresse")
	strWorkType		= rsVikar("WorkType")
	strLink1URL		= rsVikar("Link1URL")
	strLink2URL		= rsVikar("Link2URL")
	strLink3URL		= rsVikar("Link3URL")
	
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
	strPoststed = rsVikar("Postnr") & " " & rsVikar("PostSted")
	strHeading = "Viser suspect " & rsVikar("Fornavn") & " " &rsVikar("Etternavn")
	vikNavn = rsVikar("Fornavn") & " " &rsVikar("Etternavn")

	' Close and release recordset
	rsVikar.Close
	Set rsVikar = Nothing

	'sletter enkeltlinjer i kompetansen.
	If Request.QueryString("slett") = "ja" Then
		strID = Request.QueryString("ID")
		strSQL = "Delete from V_SUSPECT_KOMPETANSE where kompetanseID = " &strID
		call ExecuteCRUDSQL("DELETE FROM users where userID=" & iIMPId, conn)		
		Response.redirect "Vuspectvis.asp?suspectID=" & lVikarID
	end if
%>
<html>
	<head>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<title><%=strHeading %></title>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>		
		<script language="javaScript" type="text/javascript">
			function gaaTil()
			{
				var bokmrk="<%=Request("sted")%>";
				if (bokmrk!=""){

					document.location=(bokmrk);
				}
			}
			
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
			// her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;			
		</script>
	</head>
	<body OnLoad="gaaTil()">	
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
			</div>
			<div class="content">		
				<form Name="suspekt" action="SuspectDB.asp?oppdat=ja&suspectID=<%=lVikarID%>" METHOD="post" ID="Form1">				
					<table ID="Table1">
						<tr>
							<th>Suspectnr:</th>
							<td><%=lVikarID %></td>
						</tr>
						<tr>
							<th>Sett ansvarlig:</th>
							<td>
								<select NAME="dbxMedarbeider" ID="dbxMedarbeider" onkeydown="typeAhead()">									
									<%
									' Get ansvarlig medarbeider
									response.write GetCoWorkersAsOptionList(strAnsMed)
									%>
								</select>&nbsp;
								<INPUT TYPE="submit" VALUE="Overfør til ansvarlig" ID="Submit1" NAME="Submit1">
							</td>
						</tr>
						<tr>
							<th>Fornavn:</th>
							<td><%=strFornavn %></td>
						</tr>
						<tr>
							<th>Etternavn:</th>
							<td><%=strEtternavn %></td>
						</tr>
						<tr>
							<th>Telefon:</th>
							<td><%=strTelefon %></td>
						</tr>
						<tr>
							<th>Mobil:</th>
							<td><%=strMobilTlf %></td>
						</tr>
						<tr>
							<th>E-Post:</th>
							<td><a HREF="mailto:<%=strEPost%>"><%=strEPost%></td>
						</tr>
						<tr>
							<th>Reg.dato:</th>
							<td><%=strRegDato%></TD>
						</tr>
						<tr>
							<th>Hjemmeadresse:</th>
							<td><%=strAdresse%></td>
						</tr>
						<tr>
							<th>Poststed:</th>
							<td><%=strPoststed %></td>
						</tr>
						<tr>
							<th>Fødselsdato:</th>
							<td><%=susp.DataValues("Foedselsdato")%></td>
						</tr>
						<tr>
							<th>Kjønn:</th>
							<td><%if susp.DataValues("Kjoenn") = "m" then
									Response.Write "Mann"
								elseif susp.DataValues("Kjoenn") = "k" then
									Response.Write "Kvinne"
								end if%>
							</td>
						</tr>
						<tr>
							<th>Har førerkort:</th>
							<td><%if forerkort then Response.Write "Ja" else Response.Write "Nei" end if%></td>
						</tr>
						<tr>
							<th>Disponerer bil:</th>
							<td><%if bildisp then Response.Write "Ja" else Response.Write "Nei" end if%></td>
						</tr>
						<tr>
							<th>Oppsigelsestid:</th>
							<td><%=strOppsigelse%></td>
						</tr>
						<tr>
							<th>Stillingsbrøk:</th>
							<td><%=strWorkType%></td>
						</tr>
						<tr>
							<th>Ønsket arbeidssted:</th>
							<td>
								<%for each place in jobPl%>
									<%=place.DataValues("PlaceName")%>&nbsp;
								<%next%>
							</td>
						</tr>
						<tr>
							<th>Beskrivelse:</th>
							<td colspan="3"><%=beskrivelse%></td>
						</tr>
						<tr>
							<th>Ønskede-stillinger:</th>
							<td colspan="3"><%=interestedJobs%></td>
						</tr>
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
				</form>
			</div>
			<%
			if cv.Datavalues.Count > 1 then 
			%>	
			<div class="contentHead"><h2>Produkt- og fagkompetanse</h2></div>
			<div class="content">
				<table ID="Table2">
					<tr>
						<td>
							<table class="listing" ID="Table3">
								<tr>
									<td colspan="2"><caption>Fagkompetanse:</caption></td>
								</tr>
								<tr>
									<th>Fagområde</th>
									<th>Jobbtype</th>
								</tr>
								<%
								set ObjJobgroup = server.CreateObject("xtraweb.jobgroup")
								set RsJobWishes = ObjJobgroup.GetAllApproved(Cv.XtraConString, cv.DataValues("cvid").value)
								StrJobGroup = ""
								while not RsJobWishes.EOF
									if RsJobWishes.fields("fagomrade").value <> StrJobGroup then
										if StrJobGroup <> "" then
											Response.Write "</table></td></tr>"
										end if
										StrJobGroup = RsJobWishes.fields("fagomrade").value
										Response.Write "<tr><th valign='top'>" & StrJobGroup & "</th>"
										Response.Write "<td><table cellpadding=""0"" cellspacing=""0"">"
									end if
									Response.Write "<tr><td>" & RsJobWishes.fields("KTittel").value & "</td></tr>"
									RsJobWishes.movenext
								wend
								if StrJobGroup <> "" then
									Response.Write "</table></td></tr>"
								end if

								set ObjJobgroup	= nothing
								set RsJobWishes = nothing
								%>
							</table>
						</td>
						<td valign="top">
							<table class="listing" ID="Table4">
								<tr>
									<td colspan="2"><caption>Produktkompetanse:</caption></td>
								</tr>
								<tr>
									<th>Produktområde</th>
									<th>Produkt</th>
								</tr>
								<tr>
								<%
								set ObjProdgroup = server.CreateObject("xtraweb.productgroup")
								set rsQualifications = ObjProdgroup.GetAllApproved(CV.XtraConString, Cv.datavalues("cvid").value)
								strProduktomrade = ""
								while not rsQualifications.EOF
									if rsQualifications("Produktomrade").value <> strProduktomrade then
										if strProduktomrade <> "" then
											Response.Write "</table></td></tr>"
										end if
										strProduktomrade = rsQualifications("Produktomrade").value
										Response.Write "<tr><th valign='top'>" & StrProduktomrade & "</th>"
										Response.Write "<td><table cellpadding=""0"" cellspacing=""0"">"
									end if
									Response.Write "<tr><td>" & rsQualifications("KTittel").value & "</td></tr>"
									rsQualifications.movenext
								wend
								if strProduktomrade <> "" then
									Response.Write "</table></td></tr>"
								end if

								set ObjProdgroup		= nothing
								set rsQualifications	= nothing
								%>
							</table>
						</td>
					</tr>
				</table>
			</div>
			<%
			end if
			%>
			
						
			<%			
			if allEdu.Count > 0 then
			%>
			<div class="contentHead"><h2>Utdannelse og praksis</h2></div>
			<div class="content">
				<table class="listing" ID="Table5">
				
					<!-- Education list -->
					<tr>
						<td>
							<table width="100%" ID="Table6">
								<caption>Utdannelse</caption>
								<tr>
									<th>Periode</th>
									<th>Sted</th>
									<th>Linje</th>
									<th>Kommentar</th>
								</tr>
								<%
								for each edu in allEdu								
									
									toMonthYear = ""
									if isNull(edu.DataValues("ToMonth")) and isNull(edu.DataValues("ToYear")) then									
										toMonthYear = "d.d"
									else		
										if edu.DataValues("ToMonth") < 10 then
											toMonthYear = "0" & edu.DataValues("ToMonth") & "/" & edu.DataValues("ToYear")
										else
											toMonthYear = edu.DataValues("ToMonth") & "/" & edu.DataValues("ToYear")
										end if
									end if
								
									%>
									
									<tr>
										<td><%if edu.DataValues("FromMonth") < 10 then%>0<%end if%>
										      <%=edu.DataValues("FromMonth")%>/<%=edu.DataValues("FromYear")%> -
										    <%=toMonthYear%></td>
										<td><%=edu.DataValues("Place")%></td>
										<td><%=edu.DataValues("Title")%></td>
										<td><%=edu.DataValues("Description")%></td>
									</tr>
									<%
								next
								%>
							</table>
						</td>
					</tr>
					<%
					end if
					%>
										
					<!-- Courses list -->
					
					<%
					if(HasRows(rsCourseList)) then		
					%>
					<tr>
						<td>
							<table width="100%" ID="Table6">
								<caption>Kurs</caption>
								<tr>
									<th>Periode</th>
									<th>Sted</th>
									<th>Linje</th>
									<th>Kommentar</th>
								</tr>
								<%		
		
								Do Until (rsCourseList.EOF)
									strDataId      = rsCourseList("DataId").Value
									strPlace       = rsCourseList("Place").Value															
									strDescription = rsCourseList("Description").Value
									strTitle       = rsCourseList("Title").Value
									strFromMonth   = rsCourseList("FromMonth").Value
									strFromYear    = rsCourseList("FromYear").Value
									strToMonth     = rsCourseList("ToMonth").Value
									strToYear      = rsCourseList("ToYear").Value
								
									if strFromMonth < 10 then
										strFromMonth = "0" & strFromMonth
									end if
									if (strToMonth < 10) and (strToMonth >= 1) then
										strToMonth = "0" & strToMonth
									end if
						
									if (strToMonth>0) or (strToYear>0) then
										StrPeriod = strFromMonth & "/" & strFromYear & " - " & strToMonth & "/" & strToYear
									else
										StrPeriod = strFromMonth & "/" & strFromYear & " - d.d."				
									end if
									%>
									
									<tr>
										<td><%=StrPeriod%></td>
										<td><%=strPlace%></td>
										<td><%=strTitle%></td>
										<td><%=strDescription%></td>
									</tr>
									<%
									rsCourseList.MoveNext
									Loop
									rsCourseList.Close
									%>
							</table>
						</td>
					</tr>
					<%
					end if
					%>
					
					<!-- Experience list -->
					
					<%
					if allExp.Count > 0 then
					%>					
					<tr>
						<td valign="top">
							<table width="100%" ID="Table7">
								<caption>Praksis</caption>
								<tr>
									<th>Periode</th>
									<th>Arbeidsgiver</th>
									<th>Tittel</th>
									<th>Kommentar</th>
								</tr>
								<%
								for each exp in allExp
								
								toMonthYear = ""
								if isNull(exp.DataValues("ToMonth")) and isNull(exp.DataValues("ToYear")) then									
									toMonthYear = "d.d"
								else		
									if exp.DataValues("ToMonth") < 10 then
										toMonthYear = "0" & exp.DataValues("ToMonth") & "/" & exp.DataValues("ToYear")
									else
										toMonthYear = exp.DataValues("ToMonth") & "/" & exp.DataValues("ToYear")
									end if
								end if
									
								%>
								<tr>
									<td ><%if exp.DataValues("FromMonth") < 10 then%>0<%end if%>
										<%=exp.DataValues("FromMonth")%>/<%=exp.DataValues("FromYear")%> -
										<%=toMonthYear%></td>
									<td ><%=exp.DataValues("Place")%></td>
									<td ><%=exp.DataValues("Title")%></td>
									<td ><%=exp.DataValues("Description")%></td>
								</tr>
								<%
								next
								%>
							</table>
						</td>
					</tr>
					<%
					end if
					if allRef.Count > 0 then
					%>
					<tr>
						<td valign="top">
							<table width="100%" ID="Table8">
								<caption>Referanser</caption>
								<tr>
									<th>Navn</th>
									<th>Tittel</th>
									<th>Telefon</th>
									<th>Kommentar</th>
								</tr>
								<%
								for each ref in allRef
								%>
								<tr>
									<td><%=ref.DataValues("Name")%></td>
									<td><%=ref.DataValues("Title")%></td>
									<td><%=ref.DataValues("Tel")%></td>
									<td><%=ref.DataValues("Comment")%></td>
								</tr>
								<%
								next
								%>
							</table>
						</td>
					</tr>
					<%
					end if					
					%>
				</table>
			</div>				
			<%
			if cvfil <> "" then
				%>
				<div class="contentHead"><h2>CV-fil</h2></div>
				<div class="content">			
					<p>
						<strong>Søkeren har lastet opp eget CV-dokument!</strong>
					</p>
					<p>
						<strong>CV:</strong><br>
						
						<a href="http://<%=Application("HTTPadress")%>/Xtra/CVUpload/CVdok/<%=lVikarID %>/<%=cvfil%>" target="_blank"><%=cvfil%></a>
					</p>
				</div>
				<%
			end if
			
			if allFile.Count > 0 then
					%>
						<div class="contentHead"><h2>CV-fil</h2></div>
						<div class="content">	
						<table width="100%" id="Table11">
								<caption>Cvs:</caption>
								<%
								for i=1 to allFile.Count
								test = "/Xtra/CVUpload/CVdok/" & lVikarID & "/"
								%>
								<tr>
									
									<td><a href="http://<%=Application("HTTPadress")%><%=test%><%=allFile.Item(i).DataValues("FileName")%>" target="_blank"><%=allFile.Item(i).DataValues("FileName")%></a></td>
								</tr>
								<%
								next
								%>
						</table>
						
					<%
			end if			
			
			If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = true) Then
			%>
			<table ID="Table9">
				<tr>
					<td>
						<form action='/Xtra/WebUI/DataDeletion/OnTheFlySuspectDelete.aspx?Admin=true' onsubmit="return confirm('Are you sure you want to delete this suspect?');" method='Post' ID="Form3">
						<!-- SuspectDB.asp?slett=ja&suspectID=<%=lVikarID%> -->
						    <input type="hidden" name="delete" value="true" id="delete" />
						    <input type="hidden" name="suspectID" value="<%=lVikarID%>" id="suspectID" />
							<input type='submit' value='Slett søker' ID="Submit3" NAME="Submit3">
						</form>
					</td>
					<td>
						<form action='IsSuspectDuplicate.asp?suspectID=<%=lVikarID%>' method='Post' ID="Form4">
							<input type='submit' value='Overfør suspect' ID="Submit4" NAME="Submit4">
						</form>
					</td>
				</tr>
			</table>
			<%
			end if
			set ref		= nothing
			set exp		= nothing
			set edu		= nothing
			set uplFile = nothing
			set allEdu	= nothing
			set allExp	= nothing
			set allRef	= nothing
			set allFile = nothing
			set cv		= nothing
			set prod	= nothing
			set prodgr	= nothing
			set jobty	= nothing
			set jobgr	= nothing
			set jobPl	= nothing
			jobApp.cleanup
			set jobApp	= nothing
			susp.CV.cleanup
			set adr		= nothing
			susp.cleanup
			set susp	= nothing
			%>
			
			<br/>
			<br/>
		</div>	
		
		
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>