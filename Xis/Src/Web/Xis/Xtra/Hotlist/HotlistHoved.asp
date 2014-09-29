<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="..\includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="..\includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.SuperOffice.Integration.inc"-->
<!--#INCLUDE FILE="..\includes\xis.rights.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="..\includes\Xis.HTML.Renderfunctions.inc"-->
<%
	dim brukerID
	dim valg
	dim strID
	dim strManglendeSK
	dim strMangledeCV
	dim strManglendeKS
	dim strManglendeKM
	dim strSokereWeb
	dim rsManglendeSK
	dim rsManglendeKS
	dim rsManglendeKM
	dim rsSokereWeb
	dim rsTimeSheets
	dim strHotlistPrefix
	dim rsAnsvarlig
	dim rsResponsiblePersonsList
	'DataDeletion : PRO@EC
	dim rsSuspectCount
	dim rsTempROneCount
	dim rsTempRTwoCount
	dim rsOppdragCount
	dim rsExpiredOppdragCount
	dim objResource
	dim lLoggedInResponsibleID
	dim lSelectedResponsibleID
		
	


	set objResource = Server.CreateObject("Localizer.ResourceManager")
	
	'Session("medarbID") = 226

	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	brukerID	= Session("BrukerID")
	valg		= Request.QueryString("valg")
	SlettAction	= trim(Request.QueryString("Slett"))
	strID		= Request.QueryString("ID")

	if (Session("medarbID") = 0) then
		'user must be logged on, abort!
		AddErrorMessage(objResource.GetText("MsgNotLogged"))
		call RenderErrorMessage()
	else
		set rsAnsvarlig = GetFirehoseRS("exec [dbo].[GetConsultantManagerByID] " & Session("medarbID"), Conn)
		if (HasRows(rsAnsvarlig)) then
			dim navn
			fornavn = rsAnsvarlig.fields("fornavn")
			strHotlistPrefix = fornavn
			if(mid(fornavn, len(fornavn) - 1, 1) <> "s") then
				strHotlistPrefix = rsAnsvarlig.fields("fornavn") & "s"
			else
				strHotlistPrefix = rsAnsvarlig.fields("fornavn") & " sin "
			end if
			
		end if
		rsAnsvarlig.close
		set rsAnsvarlig = nothing
				
		lLoggedInResponsibleID = Session("medarbID")		
		if(Request.Form("ReRun") = "1") then
			lSelectedResponsibleID = Request.Form("SelectedResponsibleID")			
		else
			lSelectedResponsibleID = lLoggedInResponsibleID
		end if
		
	end if

	
	
	'response.write lLoggedInResponsibleID
	'response.write "**"
	'response.write lSelectedResponsibleID
	
	
	
	
	
	'Get all the responsible persons
	strResponsiblePersonsList = "SELECT MedID AS ResponsibleID, (Fornavn + ' ' + Etternavn) AS Name FROM MEDARBEIDER ORDER BY Fornavn"	
	set rsResponsiblePersonsList = GetFirehoseRS(strResponsiblePersonsList, Conn)	
	

	'Get data for CVhotlist
	If Session("medarbID") > 0 Then
	'If user is a consultantleader, retrieve shortcut list

'Get the count for Contract Not Sent
	


			strManglendeSK = _
			"SELECT count( distinct vikar.vikarid)" & _
			"FROM VIKAR " & _
			"LEFT OUTER JOIN ADRESSE ON VIKAR.Vikarid = ADRESSE.adresseRelID " & _
			"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
			"INNER JOIN VIKAR_UKELISTE ON VIKAR.Vikarid = VIKAR_UKELISTE.Vikarid " & _
			"WHERE VIKAR.Statusid = '3' " & _
			"AND VIKAR.typeID = 1" & _
			"AND ADRESSE.AdresseRelasjon = '2' " & _
			"AND ADRESSE.AdresseType = '1' " & _
			"AND VIKAR.AnsMedID = " & Session("medarbID") & _
			"AND ( " & _
			"( VIKAR.MottattSkattekort IS NULL ) " & _
			"OR ( VIKAR.MottattSkattekort = '-1' ) " & _
			"OR ( VIKAR.VikarId NOT IN (  " & _
				"SELECT " & _
				"VIKAR.Vikarid " & _
				"FROM VIKAR " & _
				"WHERE VIKAR.MottattSkattekort = YEAR(GETDATE()) " & _
				"OR " & _
				"( VIKAR.MottattSkattekort = (YEAR(GETDATE()) - 1) AND (MONTH(GETDATE()) = 1) ) " & _
			"))) "

		set rsManglendeSK = GetFirehoseRS(strManglendeSK, Conn)


	'-----------------------------------
		
		'Selected ALL
		If(CInt(lSelectedResponsibleID) = 0) Then

		strManglendeKS = "SELECT count(vikarid) " & _
		"FROM VIKAR  " & _
		"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
		"WHERE VIKAR.KontraktSendt Is Null AND VIKAR.Kontraktmottatt is null AND VIKAR.statusID = 3 "

'Selected a responsible person from the list or default which is the logged in user
		Else

strManglendeKS = "SELECT count(vikarid) " & _
		"FROM VIKAR  " & _
		"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
		"WHERE VIKAR.KontraktSendt Is Null AND VIKAR.Kontraktmottatt is null AND VIKAR.statusID = 3 and AnsMedID = " & CInt(lSelectedResponsibleID)
               End If

		set rsManglendeKS = GetFirehoseRS(strManglendeKS, Conn)


'Get the count for Contract Not Received		
		'-----------------------------------
		
		'Selected ALL
		If(CInt(lSelectedResponsibleID) = 0) Then

		strManglendeKM = "SELECT count(vikarid) " & _
		"FROM VIKAR  " & _
		"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
		"WHERE VIKAR.KontraktSendt Is not Null AND VIKAR.Kontraktmottatt is null " & _
		"AND ADRESSE.AdresseRelasjon = '2' " & _
		"AND ADRESSE.AdresseType = '1' " & _
		"AND VIKAR.statusID = 3"

'Selected a responsible person from the list or default which is the logged in user
		Else
		strManglendeKM = "SELECT count(vikarid) " & _
		"FROM VIKAR  " & _
		"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
		"WHERE VIKAR.KontraktSendt Is not Null AND VIKAR.Kontraktmottatt is null " & _
		"AND ADRESSE.AdresseRelasjon = '2' " & _
		"AND ADRESSE.AdresseType = '1' " & _
		"AND VIKAR.statusID = 3 and AnsMedID = " & CInt(lSelectedResponsibleID)

End If


		set rsManglendeKM = GetFirehoseRS(strManglendeKM, Conn)

		if CInt(lSelectedResponsibleID) > 0 then 		
			strSokereWeb = "Select count(suspectid) as NOF from V_SUSPECT where ansmedid = " & lSelectedResponsibleID & " " & _
		" and slettes  = 0 "&_
		" and overfort = 0 "
		if CInt(lSelectedResponsibleID) > 0 then 		
			strSokereWeb = "Select count(suspectid) as NOF from V_SUSPECT where ansmedid = " & lSelectedResponsibleID & " " & _
			" and slettes  = 0 "&_
			" and overfort = 0 "
		else
			strSokereWeb = "Select count(suspectid) as NOF from V_SUSPECT where slettes  = 0 and overfort = 0 "
		end if
		else
			strSokereWeb = "Select count(suspectid) as NOF from V_SUSPECT where slettes  = 0 and overfort = 0 "
		end if

		set rsSokereWeb = GetFirehoseRS(strSokereWeb, Conn)

		set rsTimeSheets = GetFirehoseRS("EXEC [dbo].[GetNonApprovedtimesheetCountForConsultantLeader]"  & lSelectedResponsibleID, Conn)
		
		'DataDeletion : PRO@EC
	    set rsSuspectCount = GetFirehoseRS("EXEC [dbo].[spECGetSuspectRuleCount] " & CInt(lSelectedResponsibleID), Conn)
	    	    
	    set rsTempROneCount = GetFirehoseRS("EXEC [dbo].[spECGetTempSixMonthsRuleCount] " & CInt(lSelectedResponsibleID), Conn)
	    	       
	    set rsTempRTwoCount = GetFirehoseRS("EXEC [dbo].[spECGetTempThreeYearsRuleCount] " & CInt(lSelectedResponsibleID), Conn)
	    
	    set rsOppdragCount =  GetFirehoseRS("EXEC [dbo].[spECGetCommissionsCount] " & CInt(lSelectedResponsibleID), Conn)
	    
	    set rsExpiredOppdragCount = GetFirehoseRS("EXEC [dbo].[EC_GetExpiredAndExpiringAddCount] " & lSelectedResponsibleID, Conn)
	    
	    
	End If
	
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>Hotlist</title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script language='javascript' src='../js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='../js/navigation.js' id='navigationScripts'></script>
		<script language="javaScript" type="text/javascript">
			//lager felles variabler
			function shortKey(e)
			{
				var keyChar = String.fromCharCode(event.keyCode);
				var modKey  = event.ctrlKey;
				var modKey2 = event.shiftKey;

				<%
				If HasUserRight(ACCESS_TASK, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="P")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistOppdrag.asp");
					}
					<%
				end if
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					%>
					if (modKey && modKey2 && keyChar=="V")
					{
						parent.frames[funcFrameIndex].location=("/xtra/hotlist/hotlistVikar.asp");
					}
					<%
				end if
				%>
			}

			function RefreshPage(){
				//alert("oo");
				var combobox = document.getElementById("dbxResponsible");
		  		var selectedID = combobox.options[combobox.selectedIndex].value;
		  		//alert(selectedID);
		  		document.CVForm.SelectedResponsibleID.value = selectedID;
		  		document.CVForm.ReRun.value = "1";
				document.CVForm.submit();
			}

			//Her catches eventen som trigger shortcut'en
			document.onkeydown = shortKey;
		</script>
	</head>
	
	<body onLoad="fokus()">	

<form name="CVForm" action="HotlistHoved.asp" method="post">        
	  	<input type="hidden" id="SelectedResponsibleID" name="SelectedResponsibleID">
	  	<input type="hidden" id="ReRun" name="ReRun">		        
		<div class="pageContainer" id="pageContainer">
			<a id="Top"></a>
<!-- Main Heading -->
			<div class="contentHead1">
				    <!-- easterEgg goes here -->
				    <table style="width:98%;" cellpadding="0" cellspacing="0">
				        <tr>
				            <td style="width:90%; text-align:left;"><h1><%objResource.WriteText("List")%></h1></td>
				            <td style="width:9%; text-align:right;">
				                <div onclick="window.open('/xtra/babuscha/game/mem.html')" style="display:block; width:90px; height:26px; position:relative;"></div>
				            </td>
				        </tr>
				    </table>
				
			</div>

<!-- Sub Heading with drop down list -->
<div class="content">				    
				    <table style="width:98%;" cellpadding="0" cellspacing="0" border="0">
				        <tr>
				            <td colspan="2" style="text-align:left;">&nbsp;Huskeliste for
				            <select name="dbxResponsible" id="dbxResponsible" onchange="RefreshPage()">
				            		<option value="0" id="0">All</option>
	                           		<%
												response.write GetCoWorkersAsOptionList(lLoggedInResponsibleID)
									%>
							</select>
				            
				            </td>
				            
				        </tr>
				    </table>
				
			</div>

			<div class="content">
				<div class="listing followUp">
					<table ID="Table1">
						<tr>
							<th colspan="2"><%objResource.WriteText("Followup")%></th>
							<th><%objResource.WriteText("Number")%></th>
						</tr>
						<%
						'skattekort til oppfølging
						if (rsManglendeSK.fields(0) > 0) then
							%>
							<tr>
								<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Skattekort</td>
								<td><a href='hotlistSkattekort.asp'><%=rsManglendeSK.fields(0)%></a></td>
							</tr>
							<%
						else
							%>
							<tr>
								<td>&nbsp;</td>
								<td>Skattekort</td>
								<td>ingen</td>
							</tr>
							<%
						end if
						rsManglendeSK.close
						set rsManglendeSK = nothing
						'Timelister til oppfølging
						if (rsTimeSheets.fields(0) > 0) then
							%>
							<tr>
								<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Manglende timelister fra vikarer</td>
								<td><a href='hotlistTimelister.asp?MedId=<%= lSelectedResponsibleID %>'><%=rsTimeSheets.fields(0)%></a></td>
							</tr>
							<%
						else
							%>
							<tr>
								<td>&nbsp;</td>
								<td>Manglende timelister fra vikarer</td>
								<td>ingen</td>
							</tr>
							<%
						end if
						rsTimeSheets.close
						set rsTimeSheets = nothing
						if (rsSokereWeb.fields("NOF") > 0) then
							%>
							<tr>
								<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Søkere fra web til oppfølging</td>
								<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "/xtra/jobb/SuspectList.asp?MedId=" & lSelectedResponsibleID, rsSokereWeb.fields("NOF"), "Vis søkere fra web til oppfølging" )%>	</td>
								
							</tr>
							<%
						else
							%>
							<tr>
								<td>&nbsp;</td>
								<td>Søkere fra web til oppfølging</td>
								<td>ingen</td>
							</tr>
							<%
						end if
						rsSokereWeb.close
						set rsSokereWeb = nothing
						if (rsManglendeKS.fields(0).value > 0) then
							%>
							<tr>
								<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td><%objResource.WriteText("ContractSent")%></td>
								<td><A href='#SkattekortSendt'><%=rsManglendeKS.fields(0)%></a></td>
							</tr>
							<%

'Get the count for Contract Not Sent
							'-----------------------------------
		
							'Selected ALL
							If(CInt(lSelectedResponsibleID) = 0) Then
							strManglendeKS = "SELECT DISTINCT VIKAR.VikarID, VIKAR_ANSATTNUMMER.ansattnummer, [Vikar].[Fornavn] + ' ' + [Vikar].[Etternavn] AS Navn, VIKAR.KontraktSendt, VIKAR.Epost " & _
							"FROM VIKAR  " & _
							"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
							"INNER JOIN VIKAR_ANSATTNUMMER ON VIKAR_ANSATTNUMMER.VikarID = VIKAR.VikarID " & _
							"WHERE VIKAR.KontraktSendt Is Null AND VIKAR.Kontraktmottatt is null AND VIKAR.statusID = 3"

'Selected a responsible person from the list or default which is the logged in user					
							Else
							strManglendeKS = "SELECT DISTINCT VIKAR.VikarID, VIKAR_ANSATTNUMMER.ansattnummer, [Vikar].[Fornavn] + ' ' + [Vikar].[Etternavn] AS Navn, VIKAR.KontraktSendt, VIKAR.Epost " & _
							"FROM VIKAR  " & _
							"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
							"INNER JOIN VIKAR_ANSATTNUMMER ON VIKAR_ANSATTNUMMER.VikarID = VIKAR.VikarID " & _
							"WHERE VIKAR.KontraktSendt Is Null AND VIKAR.Kontraktmottatt is null AND VIKAR.statusID = 3 and AnsMedID = " & CInt(lSelectedResponsibleID)

End If


							set rsManglendeKS = GetFirehoseRS(strManglendeKS, Conn)
						else
							%>
							<tr>
								<td>&nbsp;</td>
								<td><%objResource.WriteText("ContractSent")%></td>
								<td>ingen</td>
							</tr>
							<%
						end if
						if (rsManglendeKM.fields(0) > 0) then
							%>
							<tr>
								<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Manglende kontrakt mottatt</td>
								<td><A href='#SkattekortMottatt'><%=rsManglendeKM.fields(0)%></a></td>
							</tr>
							<%
'Get the count for Contract Not Received
							'---------------------------------------
		
							'Selected ALL
							If(CInt(lSelectedResponsibleID) = 0) Then
							strManglendeKM = "SELECT VIKAR.VikarID, VIKAR_ANSATTNUMMER.ansattnummer, [Vikar].[Fornavn] + ' ' + [Vikar].[Etternavn] AS Navn, VIKAR.KontraktSendt, VIKAR.Epost, VIKAR.KontraktSendt " & _
							"FROM VIKAR  " & _
							"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
							"INNER JOIN VIKAR_ANSATTNUMMER ON VIKAR_ANSATTNUMMER.VikarID = VIKAR.VikarID " & _
							"WHERE VIKAR.KontraktSendt Is not Null AND VIKAR.Kontraktmottatt is null " & _
							"AND ADRESSE.AdresseRelasjon = '2' " & _
							"AND ADRESSE.AdresseType = '1' " & _
							"AND VIKAR.statusID = 3 "

'Selected a responsible person from the list or default which is the logged in user					
							Else
							strManglendeKM = "SELECT VIKAR.VikarID, VIKAR_ANSATTNUMMER.ansattnummer, [Vikar].[Fornavn] + ' ' + [Vikar].[Etternavn] AS Navn, VIKAR.KontraktSendt, VIKAR.Epost, VIKAR.KontraktSendt " & _
							"FROM VIKAR  " & _
							"INNER JOIN ADRESSE ON ADRESSE.AdresseRelID = VIKAR.VikarID " & _
							"INNER JOIN VIKAR_ANSATTNUMMER ON VIKAR_ANSATTNUMMER.VikarID = VIKAR.VikarID " & _
							"WHERE VIKAR.KontraktSendt Is not Null AND VIKAR.Kontraktmottatt is null " & _
							"AND ADRESSE.AdresseRelasjon = '2' " & _
							"AND ADRESSE.AdresseType = '1' " & _
							"AND VIKAR.statusID = 3 and AnsMedID = " & CInt(lSelectedResponsibleID)

End If


							set rsManglendeKM = GetFirehoseRS(strManglendeKM, Conn)
						else
							%>
							<tr>
								<td>&nbsp;</td>
								<td>Manglende kontrakt mottatt</td>
								<td>ingen</td>
							</tr>
							<%
						end if
						'DataDeletion Links 
						'Added: PRO@EC
						if(rsSuspectCount.fields(0)) then
						    %>
						    <tr>
						        <td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Søkere fra web ? eldre enn 6 måneder</td>
								<td><A href='/xtra/WebUI/DataDeletion/SuspectDelete.aspx'><%=rsSuspectCount.fields(0)%></a></td>
						    </tr>
						    <%
						else
						    %>
						    <tr>
								<td>&nbsp;</td>
								<td>Søkere fra web ? eldre enn 6 måneder</td>
								<td><A href='/xtra/WebUI/DataDeletion/SuspectDelete.aspx'>ingen</a></td>
							</tr>
						    <%
						end if
						rsSuspectCount.close
						set rsSuspectCount = nothing
						
						if(rsTempROneCount.fields(0)) then
						    %>
						    <tr>
						        <td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Vikarer med status ?søker? uten oppfølging på 6 måneder</td>
								<td><A href='/xtra/WebUI/DataDeletion/TempDelete.aspx?Rule=1'><%=rsTempROneCount.fields(0)%></a></td>
						    </tr>
						    <%
						else
						    %>
						    <tr>
								<td>&nbsp;</td>
								<td>Vikarer med status ?søker? uten oppfølging på 6 måneder</td>
								<td><A href='/xtra/WebUI/DataDeletion/TempDelete.aspx?Rule=1'>ingen</a></td>
							</tr>
						    <%
						end if
						rsTempROneCount.close
						set rsTempROneCount = nothing
						
						if(rsTempRTwoCount.fields(0)) then
						    %>
						    <tr>
						        <td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Vikarer uten oppfølging siste 3 år</td>
								<td><A href='/xtra/WebUI/DataDeletion/TempDelete.aspx?Rule=2'><%=rsTempRTwoCount.fields(0)%></a></td>
						    </tr>
						    <%
						else
						    %>
						    <tr>
								<td>&nbsp;</td>
								<td>Vikarer uten oppfølging siste 3 år</td>
								<td><A href='/xtra/WebUI/DataDeletion/TempDelete.aspx?Rule=2'>ingen</a></td>
							</tr>
						    <%
						end if
						rsTempRTwoCount.close
						set rsTempRTwoCount = nothing
						
						if(rsOppdragCount.Fields(0)) then
						%>
						 	<tr>
						        	<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Mine oppdrag</td>
								<td><A href='/xtra/WebUI/MyCommissions.aspx?AnsmedId=<%= Session("medarbID") %>'><%= rsOppdragCount.Fields(0) %></a></td>
						    	</tr>
						<%
						else
						%>
							<tr>
								<td>&nbsp;</td>
								<td>Mine oppdrag</td>
								<td><A href='/xtra/WebUI/MyCommissions.aspx?AnsmedId=<%= Session("medarbID") %>'>ingen</a></td>
							</tr>
						<%
						end if
						rsOppdragCount.close
						set rsOppdragCount = nothing
						
						' //DataDeletion
						
						if(rsExpiredOppdragCount.Fields(0)) then
						%>
						 <tr>
						        	<td><img src="/xtra/images/icon_FollowUp.gif" width="14" height="14"></td>
								<td>Expired and near to Expire Adds</td>
								<td><A href='/xtra/WebUI/ClosingAddList.aspx?AnsmedId=<%= lSelectedResponsibleID %>'> <%= rsExpiredOppdragCount.Fields(0) %></a></td>
						    	</tr>
						<%
						else
						 %>    	
						<tr>
								<td>&nbsp;</td>
								<td>Expired and near to Expire Adds</td>
								<td><A href='/xtra/WebUI/ClosingAddList.aspx?AnsmedId=<%= Session("medarbID") %>'>ingen</a></td>
							</tr>
						
						<%
						
						end if
						rsExpiredOppdragCount.close
						set rsExpiredOppdragCount = nothing
						%>
					</table>
				</div>
				<%
				If HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) Then
					if (rsManglendeKS.fields.count > 1) then
						%>
						<div class="contentHead">
							<a id="SkattekortSendt"></a>
							<h2><%objResource.WriteText("ContractSent")%></h2>
						</div>
						<div class="content">
							<div class="listing">
								<table cellspacing="1" cellpadding="1" ID="Table2">
									<tr>
										<th><%objResource.WriteText("Vikar")%></th>
										<th><%objResource.WriteText("Epost")%></th>
									</tr>
									<%
									i = 0
									while (not rsManglendeKS.EOF)
										nummer = nummer + 1
										%>
										<tr>
											<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsManglendeKS( "VikarID" ), rsManglendeKS( "Navn"), "Vis vikar " & rsManglendeKS( "Navn") )%></td>
											<td><A HREF="mailto:<%=rsManglendeKS("Epost").Value%>" ><font class="groenn"><%=rsManglendeKS("epost").Value%></a></td>
										</tr>
										<%
										i = i + 1
										rsManglendeKS.MoveNext
									wend
									rsManglendeKS.close
									set rsManglendeKS = nothing
									%>
								</table>
								<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="<%objResource.WriteText("Top")%>"><%objResource.WriteText("Top")%><a>
							</div>
						</div>
						<%
					end if
					if (rsManglendeKM.fields.count > 1) then
						%>
						<div class="contentHead">
							<h2>Manglende kontrakt mottatt</h2>
						</div>
						<div class="content">
							<div class="listing">
							<A id="SkattekortMottatt"></a>
								<table cellspacing="1" cellpadding="1" ID="Table3">
									<tr>
										<th>Vikar</th>
										<th>epost</th>
										<th>Sendt dato</th>
									</tr>
									<%
									i = 0
									while not rsManglendeKM.EOF
										nummer = nummer + 1
										%>
										<tr>
											<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "vikarvis.asp?VikarID=" & rsManglendeKM( "VikarID" ), rsManglendeKM( "Navn"), "Vis vikar " & rsManglendeKM( "Navn") )%></td>
											<td><A HREF="mailto:<%=rsManglendeKM("Epost").Value%>" ><font class="groenn"><%=rsManglendeKM("epost").Value%></a></td>
											<td><%=rsManglendeKM("KontraktSendt").Value%></td>
										</tr>
										<%
										i = i + 1
										rsManglendeKM.MoveNext
									wend
									rsManglendeKM.close
									set rsManglendeKM = nothing
									%>
								</table>
							</div>
						<a href="#top"><img src="../Images/icon_GoToTop.gif" alt="Til toppen">Til toppen<a>
						<%
					end if
				end if
				%>
			</div>
	</body>
</html>
<%
set objResource = Nothing
CloseConnection(Conn)
set Conn = nothing
%>
