<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim customerApprovalDefaultValue
	dim customerApprovalValue
	dim oppdragid
	dim Conn
	dim strCon
	dim ConIMP 'Connection
	dim rsUser 'Recordset
	dim iIMPId 'integer, IMP userid
	dim strUsername 'string
	dim vikarId
	dim kontaktId
	dim objRightsCons
	dim objCons
	dim iTeller
	dim dataval 'datavalue
	dim objRight 'as webright
	dim strSQL 'as string
	dim rsVikar 'as recordset
	dim rsKunde 'as recordset
	dim rsUserVikar
	dim rsUserKunde
	dim strBrukernavnVikar
	dim strBrukernavnKunde
	dim strVikar 'as string
	dim strFirmaID 'as string
	dim strFirma
	dim strKontaktperson
	dim strBeskrivelse
	dim sjekk_test
	dim rsOppdrag
	dim iOppdrag
	dim strAnsattnummer
	dim postedOppId

	dim blnShowHotList
	
	Dim objUserProxy  'Web service proxy for the DNN user service
	Dim sUserServiceURL 'Url of the user web service
	Dim objUserDom
	Dim iApp
	Dim sUserXml
	
	iApp = Cint(Application("Application"))
	sUserServiceURL = Application("DNNUserServiceURL")

	iIMPId = 0
	vikarId = ""
	kontaktId = ""
	sjekk_test = false

	' Get a database connection
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))

	if len(trim(Request("OppdragID"))) > 0 then
		postedOppId = trim(Request("OppdragID"))
	elseif Request.QueryString("OppdragId") <> "" then
		postedOppId = Request.QueryString("OppdragId")
	else
		postedOppId = 0
	end if
	
	oppdragid = postedOppId
	if (clng(oppdragid)= 0 or  trim(oppdragid)="") then
		AddErrorMessage("Feil:Parameter for oppdragid mangler!")
		call RenderErrorMessage()
	end if

customerApprovalValue = 0

	StrSQL = "SELECT Oppdragskode, Beskrivelse,IsCustomerApproval FROM Oppdrag where OppdragID = " & oppdragid
	set rsOppdrag = GetFirehoseRS(StrSQL, Conn)
	if(HasRows(rsOppdrag) = true) then
		iOppdrag = rsOppdrag("Oppdragskode")
		strBeskrivelse  = rsOppdrag("Beskrivelse")
		customerApprovalValue  = rsOppdrag("IsCustomerApproval")		
		rsOppdrag.close
	end if
	set rsOppdrag = nothing

'response.write customerApprovalValue
	
	
	
	'default value
	customerApprovalDefaultValue = Application("CustomerApprovalDefaultValue")
	
	'check if a specific value has been saved for the assignment, if so the saved value should override the dafault value
	if(isnull(customerApprovalValue)) then
		customerApprovalValue = customerApprovalDefaultValue
	else
		if(customerApprovalValue = true) then
			customerApprovalValue = 1
		else
			customerApprovalValue = 0
		end if
	end if
	
	'response.write customerApprovalValue


	strSQL = "SELECT DISTINCT " & _
		"VIKAR.vikarId, " & _
		"VIKAR.Fornavn, " & _
		"VIKAR.Etternavn, " & _
		"OPPDRAG.Oppdragskode, " & _
		"VIKAR_ANSATTNUMMER.ansattnummer " & _
		"FROM OPPDRAG " & _
		"LEFT OUTER JOIN OPPDRAG_VIKAR ON OPPDRAG.Oppdragid = OPPDRAG_VIKAR.Oppdragid " & _
		"LEFT OUTER JOIN VIKAR ON OPPDRAG_VIKAR.Vikarid = VIKAR.Vikarid " & _
		"LEFT OUTER JOIN VIKAR_ANSATTNUMMER ON VIKAR.Vikarid = VIKAR_ANSATTNUMMER.Vikarid " & _
		"WHERE OPPDRAG.Oppdragid = '" & oppdragid & "' " & _
		"AND OPPDRAG_VIKAR.Statusid = '4' "

	'Get Customer/Contact information
	set rsVikar = GetFirehoseRS(StrSQL, Conn)
	if not rsVikar.EOF and not rsVikar.BOF then
		vikarId = rsVikar("VikarID")
		strVikar = rsVikar("Fornavn") & " " & rsVikar("Etternavn")
		strAnsattnummer = rsVikar("ansattnummer").Value
	end if
	rsVikar.close
	set rsVikar = nothing
	
	if lenb(vikarId) <> 0  then
		'Set ConIMP = GetConnection(GetConnectionstring(EFP, ""))

		'Return IMP userNames from mappingtable
		'strBrukernavnVikar = ""
		'StrSQL = "SELECT u.username FROM imp_xtra_users ixu, users u WHERE xtra_id = " & vikarId & " AND type='ANSATT' AND ixu.userid = u.userid"
		'set rsUserVikar = GetFirehoseRS(StrSQL, ConIMP)
		'if (HasRows(rsUserVikar) = true) then
			'Store username in strusername for later use
		'	strBrukernavnVikar = rsUserVikar("username")
		'end if
		'rsUserVikar.close
		'set rsUserVikar = nothing

		'CloseConnection(conIMP)
		'set conIMP = nothing
		
		Set objUserProxy = Server.CreateObject("ECDNNUserProxy.UserServiceProxy")
		objUserProxy.Url = sUserServiceURL
		sUserXml = objUserProxy.GetUser(iApp, vikarId,"V")

		if sUserXml <> "" then
			Set objUserDom = Server.CreateObject("Microsoft.XMLDOM")
			objUserDom.LoadXml sUserXml
			
			strBrukernavnVikar = objUserDom.selectSingleNode("/user/userName").Text 
		end if	
	end if

	'Gets the consultant and the consultants Web rights..
	set objCons = Server.CreateObject("XtraWeb.Consultant")
	objCons.XtraConString = Application("Xtra_intern_ConnectionString")
	objCons.XtraDataShapeConString = Application("ConXtraShape")

	if not objCons.GetConsultant(vikarId) then
		AddErrorMessage("Feil:Kunne ikke finne vikaren!")
		call RenderErrorMessage()
	end if

	set objRightsCons = ObjCons.GetWebRights
	objRightsCons.GetTaskRights(oppdragid)

	'Rerun = save checked values
	if Request.Form("rerun") = "1" then
		'Save the Customer approval of timesheets value to db
		
		' Update oppdrag_vikar med Timeliste created
		If (Request.Form("consBox5") <> "") Then
			approvalcheck = 1		
		Else
			approvalcheck = 0
		end if
		
   		strSQL = "Update [oppdrag] set [IsCustomerApproval] = " & approvalcheck & " WHERE oppdragid =" & oppdragid
   		
   		'response.write strSQL
   		
   		'dddd
   		   		
		If ExecuteCRUDSQL(strSQL, Conn) = false then
			Conn.RollbackTrans
			CloseConnection(Conn)
			set Conn = nothing
			AddErrorMessage("Error Updating Timesheet")				
			call RenderErrorMessage()
		End if
				
		
		'BUG i save rutina..
		for iTeller = 1 to objRightsCons.Count
			set objright = objRightsCons.Item(iTeller)
			objright.datavalues("vikarID") = vikarId
			objright.datavalues("oppdragID") = oppdragid
			if Request.Form("consBox" & iTeller) = "1" then
				objright.datavalues("checked") = "CHECKED"
			else
				objright.datavalues("checked") = ""
			end if
		next
		if not objRightsCons.Save(Application("Xtra_intern_ConnectionString")) then
			AddErrorMessage("Feil:Vikarens rettigheter ble ikke lagret!")
			call RenderErrorMessage()			
		end if
		set objRightsCons = nothing
		objCons.cleanup
		set objCons = nothing		
		Response.Redirect("WebUI/OppdragView.aspx?OppdragID=" & oppdragid)
	end if
	
	CloseConnection(Conn)
	set Conn = nothing
	
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script language='javascript' src='/xtra/js/menu.js' id='menuScripts'></script>
		<script language='javascript' src='/xtra/js/navigation.js' id='navigationScripts'></script>		
		<script language="javaScript" type="text/javascript">
			function enable(check) 
			{
				if (check == 0) 
				{
					//alert("Sjekket");
					document.rettigheter.sjekket.value = 1;
					//document.rettigheter.custBox2.checked = false;
					//document.rettigheter.custBox2.disabled = false;
				} 
				else if (check == 1) 
				{
					//alert("Ikke Sjekket");
					document.rettigheter.sjekket.value = 0;
					//document.rettigheter.custBox2.checked = false;
					//document.rettigheter.custBox2.disabled = true;
				} 
				return true;
			}
			
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
	<body>
		<form id="frmRettigheter" method="post" action="RettigheterWeb.asp">
			<div class="pageContainer" id="pageContainer">
				<div class="contentHead1">
					<h1>Rettigheter på web</h1>
					<div class="contentMenu">
						<input TYPE="HIDDEN" NAME="oppdragID" VALUE="<%=oppdragid%>">
						<input type="hidden" name="rerun" value="1">
						<table width="96%">
							<tr>
								<td>
									<table>
										<tr>
											<td class="menu" id="menu1" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);">
												<a href="javascript:document.all.frmRettigheter.submit();" title="Lagre">
													<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
												</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="right"><!--#include file="Includes/contentToolsMenu.asp"--></td>
							</tr>
						</table>
					</div>
				</div>
				<div class="content">
					<%
					set objRight = objRightsCons.Item(1)
					%>
					<input type="hidden" name="sjekket" value="<%if objright.Datavalues("checked") = "CHECKED" then%> 1 <%else%> 0 <%end if%>">
					<table width="98%">
						<col width="33%">
						<col width="33%">
						<tr>
							<td>
								<table>
									<tr>
										<th>Oppdragsnr:</th>
										<td><a href="WebUI/OppdragView.aspx?OppdragID=<%=oppdragid%>" title="Tilbake til oppdraget"><%=oppdragid%></a></td>
									</tr>
									<tr>
										<th>Beskrivelse:</th>
										<td><%=strBeskrivelse%></td>
									</tr>
								</table>
							</td>
							<td>
								<table>
									<tr>
										<th>Ansattnummer:</th>
										<td><%=strAnsattnummer%></td>
									</tr>
									<tr>
										<th>Navn:</th>
										<td><%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & vikarId, strVikar, "Vis vikar " & strVikar )%></td>
									</tr>
									<tr>
										<th>Brukernavn:</th>
										<td><%=strBrukernavnVikar%></td>
									</tr>
									<tr>
										<th colspan="2">Rettigheter på web for vikar:</th>
									</tr>
									<tr>
										<td colspan="2">
											<%
											set objRight = objRightsCons.Item(2)
											%>
											<input class="checkbox" type="checkbox" <%if strBrukernavnVikar = "" then%> disabled <%end if%> <%=objright.datavalues("checked").value%> value="1" id=checkbox1 name="consBox2">
												<%=objright.datavalues("intraNavn").Value%><br>
												<%
												set objRight = objRightsCons.Item(1)
												if objright.Datavalues("checked") = "CHECKED" then
													sjekk_test = true
												end if
												%>
												<input class="checkbox" type="checkbox" <%if strBrukernavnVikar = "" then%> disabled <%end if%> <%=objright.datavalues("checked").value%> value="1" id=checkbox1 name="consBox1"
												<%if strBrukernavnKunde <> "" then%>onclick="enable(document.rettigheter.sjekket.value)"<%end if%>>
												<%=objright.datavalues("intraNavn").Value%><br>
												<%
											if iOppdrag = 0 then
												set objRight = objRightsCons.Item(3)
												%>
												<input class="checkbox" type="checkbox" <%if strBrukernavnVikar = "" then%> disabled <%end if%> <%=objright.datavalues("checked").value%> value="1" id=checkbox1 name="consBox3">
												<%=objright.datavalues("intraNavn").Value%><br>
												<%
											elseif iOppdrag = 1 then
												set objRight = objRightsCons.Item(4)%>
												<input class="checkbox" type="checkbox" <%if strBrukernavnVikar = "" then%> disabled <%end if%> <%=objright.datavalues("checked").value%> value="1" id=checkbox1 name="consBox4">
												<%=objright.datavalues("intraNavn").Value%><br>
												<%
											end if
											%>
											<input class="checkbox" type="checkbox" <%if strBrukernavnVikar = "" then%> disabled <%end if%> <%if customerApprovalValue = "1" then%> checked <%end if%> id=consBox5 name="consBox5">
											Aktiver kundegodkjenning
											<br>
										</td>
									</tr>								
								</table>
							</td>
						</tr>
					</table>
					<%
					set objRightsCons = nothing
					objCons.cleanup
					set objCons = nothing
					%>
				</div>
			</div>
		</form>
	</body>
</html>