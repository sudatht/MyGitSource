<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	'File purpose:		Displays the approved CV for a given consultant.
	'Created by:		Monica Johansen@electricfarm.no
	'Changed by:		Fred.myklebust@electricfarm.no
	'Changes:			13.12.2002
	'					Total rewrite - this file now lets a consultantleader generate CVs.

	dim strSQL
	dim StrCvfil			'as string
	dim ObjConsultant		'as xtraweb.consultant
	dim objAdr
	dim ObjCv				'as xtraweb.cv
	dim ObjEducation		'as xtraweb.Education
	dim ObjEducations		'as xtraweb.Educations
	dim rsQualifications	'as adodb.recordset
	dim rsAreas
	dim lngAreaID
	dim StrFromMonth 		'as string
	dim iFromYear 			'as string
	dim iToMonth 			'as string
	dim iToYear 			'as string
	dim strPeriod			'as string
	dim ObjCon				'as ADODB.Connection
	dim strVikarSQL			'as string
	dim rsVikar				'as adodb.recordset
	dim strAnsattnummer		'as string
	dim rcApproved
	dim strNivaa
	dim strProduct
	dim strBodyOnload
	dim strURL
	dim blnIsPosted 
	dim strProductAreas
		
	'Variables used to maintain page state
	dim strCheckName
	dim StrCheckAdress
	dim strCheckHomeAdress
	dim strCheckPhone
	dim strCheckEmail
	dim strCheckCountry
	dim strCheckBirthDate
	dim strCheckEducation
	dim strCheckCourses
	dim strCheckExperience
	dim strCheckKeyQualifcations
	dim strCheckOtherInformation
	dim strCheckReferences
	dim strCheckProductAreas
	dim strRdoTarget
	dim strChecked

	'Global variables - used in tool menu
	dim blnShowHotList
	dim lngVikarID
	
	'Initialize variables
	blnShowHotList = false	
	
	strAnsattnummer = "-"

	if Request("VikarID") <> "" then
		lngVikarID = Request("VikarID")
	end if

	set ObjConsultant = Server.CreateObject("XtraWeb.Consultant")
	ObjConsultant.XtraConString = Application("Xtra_intern_ConnectionString")
	ObjConsultant.GetConsultant(lngVikarID)
	if (ObjConsultant.Addresses.count>0) then
		set objAdr	= Server.CreateObject("XtraWeb.Address")	
		set objAdr	= ObjConsultant.Addresses(1)
	end if

	set ObjCv = ObjConsultant.CV
	ObjCv.XtraConString = Application("Xtra_intern_ConnectionString")
	ObjCv.XtraDataShapeConString = Application("ConXtraShape")
	ObjCv.Refresh

	if (ObjCv.Datavalues.Count > 1) then
		if not isNull(ObjCv.DataValues("Filename")) then
			StrCvfil = ObjCv.DataValues("Filename")
		else
			StrCvfil = ""
		end if
	else
		StrCvfil = ""
		ObjCv.save
	end if

	' Open database connection
	Set objCon = GetConnection(GetConnectionstring(XIS, ""))		

	strVikarSQL = "SELECT VIKAR_ANSATTNUMMER.ansattnummer " & _
	"FROM VIKAR_ANSATTNUMMER " & _
	"WHERE VIKAR_ANSATTNUMMER.Vikarid = '" & lngVikarID & "' "

	set rsVikar = GetFirehoseRS(strVikarSQL, objCon)	

	if HasRows(rsVikar) = true then
		strAnsattnummer = rsVikar("ansattnummer").Value
	end if
	rsVikar.close
	set rsVikar = nothing

	if (request("hdnPosted")="1")  then
		blnIsPosted = true
		strProductAreas = trim(request("chkProductAreas"))
		if (len(strProductAreas)>0) then
			strProductAreas = strProductAreas & ","
		end if
		strBodyOnload = "onLoad='javascript:Popwindow();' "
		strRdoTarget = request("chkCVTarget")
		strURL = "'vikarCVGenererHTML" & ".asp?dest=" & strRdoTarget & "&" & Request.form & "'"

		'Initialize to selection criterias to not selected
		strCheckName = ""
		StrCheckAdress = ""
		strCheckHomeAdress = ""
		strCheckPhone = ""
		strCheckEmail = ""
		strCheckCountry = ""
		strCheckBirthDate = ""
		strCheckEducation = ""
		strCheckCourses = ""
		strCheckExperience = ""
		strCheckKeyQualifcations = ""
		strCheckOtherInformation = ""
		strCheckReferences = ""
		strCheckProductAreas = ""				

		if request("chkName")="1" then
			strCheckName = "checked"
		end if
		if request("chkAdress")="1" then
			StrCheckAdress = "checked"		
		end if			
		if request("chkHomeAdress")="1" then
			strCheckHomeAdress = "checked"		
		end if			
		if request("chkEmail")="1" then
			strCheckEmail = "checked"
		end if
		if request("chkCountry")="1" then
			strCheckCountry = "checked"
		end if
		if request("chkPhone")="1" then
			strCheckPhone = "checked"
		end if
		if request("chkBirthDate")="1" then
			strCheckBirthDate = "checked"		
		end if
		if request("chkEducation")="1" then
			strCheckEducation = "checked"		
		end if
		if request("chkCourses")="1" then
			strCheckCourses = "checked"		
		end if
		if request("chkExperience")="1" then
			strCheckExperience = "checked"		
		end if
		if request("chkKeyQualifcations")="1" then
			strCheckKeyQualifcations = "checked"		
		end if
		if request("chkOtherInformation")="1" then
			strCheckOtherInformation = "checked"		
		end if
		if request("chkReferences")="1" then
			strCheckReferences = "checked"		
		end if
	else
		strBodyOnload = ""
		strURL = "'vikarCVGenererHTML.asp" & "'"
		'Default selection criteria
		strCheckName = "checked"
		StrCheckAdress = ""
		strCheckHomeAdress = ""
		strCheckPhone = ""
		strCheckEmail = ""
		strCheckCountry = ""
		strCheckBirthDate = "checked"
		strCheckEducation = "checked"
		strCheckCourses = "checked"
		strCheckExperience = "checked"
		strCheckKeyQualifcations = ""
		strCheckOtherInformation = ""
		strCheckReferences = ""
		strCheckProductAreas = "checked"
		strProductAreas = " "
	end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<meta name="generator" content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>	
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
		<script language="jscript">		

			function Popwindow()
			{
				window.open(<%=strURL%>, 'CV', 'menubar=yes,toolbar=yes,status=yes');
			}

			function DeSelect(objAll)
			{

				if (objAll.checked==true)
				{
					selectAll()
				}else{
					deSelectAll()
				}
			}

			function selectAll()
			{
				for(i=0;i<document.frmGenererCV.chkProductAreas.length;i++)
					{
						document.frmGenererCV.chkProductAreas[i].checked=true;
					}


				document.all.chkProductAreas.checked = true;
			};




			function deSelectAll()
			{
				for(i=0;i<document.frmGenererCV.chkProductAreas.length;i++)
					{
						document.frmGenererCV.chkProductAreas[i].checked=false;
					}
			};		

			function unselectAllCheckBox(oCheck)
			{
				if (oCheck.checked==false)
				{
					document.frmGenererCV.chkProductAreaAll.checked = false;
				}
			};		
			
		</script>		
		<title>&nbsp;Generere CV</title>
	</head>
	<body <%=strBodyOnload%>>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Generere CV for <%=ObjConsultant.DataValues("Fornavn") & " " & ObjConsultant.DataValues("Etternavn")%></a></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%">
						<tr>
							<td><!--#include file="vikarCVnyToolbar.asp"--></td>
							<td class="right">
							<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>
						</tr>
					</table>
				</div>
			</div>
			<div class="contentMenu2">
				<span class="menu2" id="1" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPersonalia.asp?VikarID=<%=lngVikarID%>">Personalia</a></span>
				<span class="menu2" id="2" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyJobbonsker.asp?VikarID=<%=lngVikarID%>">Fagkompetanse</a></span>
				<span class="menu2" id="3" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Produktkompetanse</a></span>
				<span class="menu2" id="4" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyNokkelKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Kandidatpresentasjon</a></span>
				<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=lngVikarID%>">Utdannelse</a></span>
				<span class="menu2" id="6" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>">Kurs</a></span>
				<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>">Yrkeserfaring</a></span>
				<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
				<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>">Referanser</a></span>
				<span class="menu2 active" id="10"><img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle"><strong>&nbsp;Generere CV</strong></span>			
			</div>		
			<div class="content">
				<form ACTION="VikarCVVis.asp" METHOD="POST" id="frmGenererCV" name="frmGenererCV">
					<input type="hidden" id="VikarID" name="VikarID" value="<%=lngVikarID%>">
					<input type="hidden" id="hdnPosted" name="hdnPosted" value="1">
					<h2>Velg CV elementer</h2>		
					<table ID="Table3">
					<col width="33%">
					<col width="34%">
					<col width="33%">						
						<tr>
							<td>
								<table>
									<tr>
										<td><input type="checkbox" <%=strCheckName%> class="checkbox" value="1" NAME="chkName" ID="chkName">&nbsp;Navn</td>
									</tr>
									<tr>								
										<td><input type="checkbox" <%=StrCheckAdress%> class="checkbox" value="1" NAME="chkAdress" ID="chkAdress">&nbsp;Adresse</td>
									</tr>
									<tr>								
										<td><input type="checkbox" <%=strCheckHomeAdress%> class="checkbox" value="1" NAME="chkHomeAdress" ID="chkHomeAdress">&nbsp;Poststed</td>								
									</tr>
									<tr>								
										<td><input type="checkbox" <%=strCheckPhone%> class="checkbox" value="1" NAME="chkPhone" ID="chkPhone">&nbsp;Telefonnummer</td>
									</tr>
									<tr>								
										<td><input type="checkbox" <%=strCheckEmail%> class="checkbox" value="1" NAME="chkEmail" ID="chkEmail">&nbsp;E-postadresse</td>
									</tr>								
									<tr>								
										<td><input type="checkbox" <%=strCheckCountry%> class="checkbox" value="1" NAME="chkCountry" ID="chkCountry">&nbsp;Nasjonalitet</td>
									</tr>								
									<tr>								
										<td><input type="checkbox" <%=strCheckBirthDate%> class="checkbox" value="1" NAME="chkBirthDate" ID="chkBirthDate">&nbsp;F&oslash;dselsdato</td>								
									</tr>							
									
								</table>
							</td>
							<td>
								<table>
									<tr>		
										<td><input type="checkbox" <%=strCheckEducation%> class="checkbox" value="1" NAME="chkEducation" ID="chkEducation">Utdannelse</td>
									</tr>
									<tr>		
										<td><input type="checkbox" <%=strCheckCourses%> class="checkbox" value="1" NAME="chkCourses" ID="chkCourses">Kurs</td>
									</tr>
									<tr>		
										<td><input type="checkbox" <%=strCheckExperience%> class="checkbox" value="1" NAME="chkExperience" ID="chkExperience">Erfaring</td>
									</tr>
									<tr>		
										<td><input type="checkbox" <%=strCheckKeyQualifcations%> class="checkbox" value="1" NAME="chkKeyQualifcations" ID="chkKeyQualifcations">&nbsp;Kandidatpresentasjon</td>
									</tr>
									<tr>		
										<td><input type="checkbox" <%=strCheckOtherInformation%> class="checkbox" value="1" NAME="chkOtherInformation" ID="chkOtherInformation">&nbsp;Kjernekompetanse</td>
									</tr>								
									<tr>		
										<td><input type="checkbox" <%=strCheckReferences%> class="checkbox" value="1" NAME="chkReferences" ID="chkReferences">&nbsp;Referanser</td>
									</tr>								
								</table>
							</td>								
							<td>
								<table id="Table5">
									<tr>
										<td>
											<strong>Vis produktkompetanse:</strong>
											<DIV class="container">
											<%
											'Get all available product areas
											strSQL = "select [ProdOmradeID],[Produktomrade]  from [h_komp_area]"
											set rsAreas = GetFirehoseRS(strSQL, ObjCon)	
											%>
											<INPUT class="checkbox" TYPE="CHECKBOX" <%if ubound(split(strProductAreas,","))=rsAreas.recordcount then%>checked<%end if%> ID="chkProductAreaAll" NAME="chkProductAreaAll" onClick="javascript:DeSelect(this)" VALUE="">Alle<br>										
											<%
											Do Until (rsAreas.EOF)
												lngAreaID = rsAreas("ProdOmradeID")	
												if blnIsPosted then
													if instr(1, strProductAreas, lngAreaID) then
														strCheckProductAreas = " checked "
													else
														strCheckProductAreas = ""												
													end if
												end if
												Response.Write "<INPUT class=""checkbox"" " & strCheckProductAreas & " TYPE=""CHECKBOX"" ID=""chkProductAreas"" NAME=""chkProductAreas"" onClick=""javascript:unselectAllCheckBox(this);"" VALUE=""" & lngAreaID & """>" & rsAreas("Produktomrade") & "<br>"
												rsAreas.MoveNext
											Loop
											rsAreas.close
											Set rsAreas = Nothing
											%>
											</DIV>
										</td>
									</tr>
								</table>
							</td>								
						</tr>
					</table>
					<h2>Generer CV i</h2>						
					<table>
						<tr>
							<%
							if (strRdoTarget="Application") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																								
							<td><input type="radio" class="radio" <%=strChecked%> value="Application" NAME="chkCVTarget" ID="chkCVTarget">Word<span class="warning">&nbsp;(Åpnes som HTML, må lagres i annet format)</span></td>
						</tr>					
						<tr>
							<%
							if (strRdoTarget="Browser") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																													
							<td><input type="radio" class="radio" <%=strChecked%> checked value="Browser" NAME="chkCVTarget" ID="chkCVTarget">Nettleser</td>							
						</tr>					
						<tr>																								
							<%
							if (strRdoTarget="Disk") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																													
							<td><input type="radio" class="radio" <%=strChecked%> value="Disk" NAME="chkCVTarget" ID="chkCVTarget">disk</td>
						</tr>									
						<tr>
							<td>&nbsp;</td>
						</tr>
					</table>
					<span class="menuInside" title="Generer CV"><a href="#" onClick="javascript:document.all.frmGenererCV.submit()"><img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generer CV</a></span><br>
					&nbsp;
				</form>
			</div>
		</div>
	</body>
</html>
<%
' sletter alle CV objekter..
set ObjCv		= nothing
ObjConsultant.CV.cleanup
ObjConsultant.cleanup
set ObjConsultant	= nothing
CloseConnection(objCon)
set objCon = nothing
%>

