<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<!-- #include file = "cuteeditor_files/include_CuteEditor.asp" --> 
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim cvfil 'as string
	dim blnLocked 			'as boolean
	dim fulltime
	dim beskrivelse
	dim courseid
	dim feilmelding
	dim dette_aar
	dim strAction
	

	'Global variables - used in tool menu
	dim blnShowHotList
	dim lngVikarID

   // Text Editor Code - CuteEditor	
	Dim editor	
	Set editor = New CuteEditor
	editor.ID = "Editor1"
	editor.AutoConfigure = "Simple"			
	editor.focus = true
	
	'Initialize variables
	blnShowHotList = false

	set cons	= Server.CreateObject("XtraWeb.Consultant")

	lVikarID    = Request.Querystring("VikarID")
	lngVikarID	= CStr(lVikarID)
	cons.XtraConString = Application("XtraWebConnection")
	cons.GetConsultant(lVikarID)

	dette_aar = Year(date)
	
	'Response.Write "loading of page "

	set cv	= cons.CV
	cv.XtraConString = Application("Xtra_intern_ConnectionString")
	cv.XtraDataShapeConString = Application("ConXtraShape")
	cv.Refresh
	
	'Creates a DB connection to retrieve the CV record
	Set cvConn = GetClientConnection(GetConnectionstring(XIS, ""))	
	set rsCV = GetFirehoseRS("SELECT max(CVId) AS CVId FROM CV WHERE ConsultantId =" & lngVikarID, cvConn)	
	vikarcvid = rsCV("CVId")	
	rsCV.Close	
	
	'Creates a DB connection to retrieve the course list
    Set listConn = GetClientConnection(GetConnectionstring(XIS, ""))
	Set rsCourseList	= GetFirehoseRS("SELECT DataId,CVId,Place,Country,Description,Title,FieldType,FromMonth,FromYear,ToMonth,ToYear FROM CV_Data WHERE FieldType = 'COU' and CVId =" & vikarcvid & " ORDER BY FromYear desc, FromMonth desc", listConn)	

	'If Consultant doesn't have CV, create one
	if cons.CV.DataValues.Count = 0 then
		cons.CV.Save
	end if

	blnLocked = cv.islocked
	courseid = 0

	if Request.Form("rerun") = "1" then
		strAction  = lcase(trim(request("pbnDataAction")))
		courseid = lcase(trim(request("course")))		
		
		if (isnull(courseid) = true) then
			courseid = 0
		end if		

		' Inserts a new Course or Updates a changed course
		if (strAction = "lagre") then				
			utdFromMonth = CStr(Request.Form("frommonth"))			
			utdFromYear	 = CStr(Request.Form("fromyear"))
			utdToMonth	 = CStr(Request.Form("tomonth"))
			utdToYear	 = CStr(Request.Form("toyear"))
			utdPlace	 = CStr(Request.Form("place"))
			utdTitle	 = CStr(Request.Form("title"))			
			
			if len(trim(Request.Form("Editor1"))) > 0 then
				utdComment	 = CStr(Request.Form("Editor1"))		
			else
			   utdComment = " "
			  end if				
			
			
			' Insert the new course
			if (courseid = 0) then
				strSQL = "Insert into CV_Data(CVId,Place,Description,Title,FieldType,FromMonth,FromYear,ToMonth,ToYear)" & _
			         " Values(" & vikarcvid & _
			         ",'" & utdPlace & "'" & _			         
			         ",'" & utdComment & "'" & _ 
			         ",'" & utdTitle & "'" & _
			         ",'COU'" & _
			         "," & utdFromMonth & _
			         "," & utdFromYear
			         
			    if (utdToMonth = "" and utdToYear = "") then
			        strSQL = strSQL &  ",NULL,NULL)"
			    else
			     	strSQL = strSQL & "," & utdToMonth & _
			      	                "," & utdToYear & ") "
			    end if
			         
			    Response.Write(strSQL)
			
				' Creates a DB connection to Insert data
				Set insertConn = GetClientConnection(GetConnectionstring(XIS, ""))	
				If ExecuteCRUDSQL(strSQL, insertConn) = false then
					Response.Write "Error generated"
					CloseConnection(insertConn)
					set insertConn = nothing
					AddErrorMessage("Feil oppstod under overføring til ansvarlig....")
					call RenderErrorMessage()
				End if
			
			' Update the selected course with new information
			else
				strSQL = "Update CV_Data Set " & _
			         "Place='" & utdPlace  & "'," & _
			         "Description='" & utdComment  & "'," & _
			         "Title='" & utdTitle  + "'," & _
			         "FromMonth=" & utdFromMonth  & "," & _
			         "FromYear=" & utdFromYear  & "," 			         
			         
			    if (utdToMonth = "" and utdToYear = "") then
			        strSQL = strSQL & "ToMonth=NULL,ToYear=NULL Where DataId=" & courseid
			    else
			     	strSQL = strSQL & "ToMonth=" & utdToMonth & "," & _
			      	                  "ToYear=" & utdToYear & "Where DataId=" & courseid
			    end if
			        
			    'response.Write(strSQL)
			
				' Creates a DB connection to Update data				
				Set updateConn = GetClientConnection(GetConnectionstring(XIS, ""))	
				If ExecuteCRUDSQL(strSQL, updateConn) = false then
					CloseConnection(updateConn)
					set updateConn = nothing
					AddErrorMessage("Feil oppstod under overføring til ansvarlig.")
					call RenderErrorMessage()
				End if
			end if
			
			set utdFromMonth = nothing
			set utdFromYear	 = nothing
			set utdToMonth	 = nothing
			set utdToYear	 = nothing
			set utdPlace	 = nothing
			set utdTitle	 = nothing
			set utdComment	 = nothing			
			
			courseid = 0
			Response.Redirect("vikarCVnyCourses.asp?type=4&VikarID=" & lngVikarID)
		end if

		' Removes a selected course
		if strAction = "slett" then
			courseid = clng(Request.Form("course"))
			if courseid > 0 then
				strSQL = "Delete from CV_Data where DataId = " & courseid
				Set deleteConn = GetClientConnection(GetConnectionstring(XIS, ""))	
				call ExecuteCRUDSQL(strSQL, deleteConn)				
				courseid = 0
				Response.Redirect("vikarCVnyCourses.asp?type=4&VikarID=" & lngVikarID)
			end if
		end if
		
		' Display the contents of the selected course
		if strAction = "show" then
			if (courseid <> 0) then
				' Creates a DB connection to retrieve selected Course data 
    			Set selectConn = GetClientConnection(GetConnectionstring(XIS, ""))	
    			set rsCourseObj = GetFirehoseRS("SELECT DataId,CVId,Place,Country,Description,Title,FieldType,FromMonth,FromYear,ToMonth,ToYear FROM CV_Data WHERE DataId =" + courseid, selectConn)
				objDataId      = rsCourseObj("DataId").Value
				objPlace       = rsCourseObj("Place").Value															
				objDescription = rsCourseObj("Description").Value
				objDescription = Replace(objDescription,"<P>","")
				objDescription = Replace(objDescription,"</P>","")
				
				objTitle       = rsCourseObj("Title").Value
				objFromMonth   = rsCourseObj("FromMonth").Value
				objFromYear    = rsCourseObj("FromYear").Value
				objToMonth     = rsCourseObj("ToMonth").Value
				objToYear      = rsCourseObj("ToYear").Value
				rsCourseObj.Close
			end if
		end if
		
	end if
	utdComment = "<link type='text/css' rel='stylesheet' href='http://" & Application("HTTPadress") & "/xtra/css/CV.css' title='default style'>" & utdComment
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
		
		<script language="javaScript" type="text/javascript">			
			function lagre_data(id){
				// Make sure all mandatory fields are filled
				if (trim(document.forms["CVForm"].frommonth.value) == ""){
    				alert("Fra-måned er ikke utfyllt!");	// From month not filled!
    				return false;
  				}   
				if (trim(document.forms["CVForm"].fromyear.value) == ""){
    				alert("Fra-år er ikke utfyllt!");		// From year not filled!
    				return false;
  				}					
  				if (trim(document.forms["CVForm"].toyear.value) != ""){
  					if (trim(document.forms["CVForm"].fromyear.value) > trim(document.forms["CVForm"].toyear.value)){
    					alert("Fra år bør være før til år!");	// From year should be less than to year!
    					return false;
  					}
  				}			
  				if (trim(document.forms["CVForm"].fromyear.value) == trim(document.forms["CVForm"].toyear.value)){
  					if (trim(document.forms["CVForm"].frommonth.value) > trim(document.forms["CVForm"].tomonth.value)){
    					alert("Fra måned bør være før til måned!");		// From month should be less than to month!
    					return false;
    				}
  				}
  				if (trim(document.forms["CVForm"].toyear.value) != ""){
  					if (trim(document.forms["CVForm"].tomonth.value) == ""){
    					alert("Til-måned er ikke utfyllt!");			// To month is not filled!
    					return false;
  					}
  				}
  				if (trim(document.forms["CVForm"].tomonth.value) != ""){
  					if (trim(document.forms["CVForm"].toyear.value) == ""){
    					alert("Til-år er ikke utfyllt!");				// To year is not filled!
    					return false;
  					}
  				}
  				if (trim(document.forms["CVForm"].place.value) == ""){
    				alert("Sted er ikke utfyllt!");						// Place is not filled!
    				return false;
  				}   
  				if (trim(document.forms["CVForm"].title.value) == ""){
    				alert("Tittel er ikke utfyllt!");						// Title is not filled!
    				return false;
  				}
				document.all.pbnDataAction.value='lagre';
				document.CVForm.course.value = id;
				document.CVForm.submit();
			}
			
			function slett_data(id){
				document.all.pbnDataAction.value='slett';
				document.CVForm.course.value = id;
				document.CVForm.submit();
			}
			
			function edit_data(id){
				document.all.pbnDataAction.value='show';
				document.CVForm.course.value = id;
				document.CVForm.submit();
			}
			
			function ValidatePeriod(field, type){
				var value = field.value;
				type = type.toLowerCase();
				var errorMsgType = "";

				if (value != null && value != ""){
					if (type == 'mm'){
						matchPattern = /^(0[1-9])|(1[0-2])$/;
						errorMsgType = "måneds feltet."
					}
					else if (type == 'yy'){
						matchPattern = /^(19\d{2})|(20\d{2})$/;
						errorMsgType = "års feltet."
					}
					if (!matchPattern.test(value)){
						alert('Ugyldig verdi i ' + errorMsgType + "\nVennligst korriger og prøv på nytt!");
						field.focus();
					}
				}
			}
			
			/* Method - Removes leading and trailing spaces and mutiple consective spaces and replaces with one space in the string  */
			function trim(inputString){
				if (typeof inputString != "string") {
					return inputString;
			    }
			    var retValue = inputString;
			    var ch       = retValue.substring(0, 1);
			    while (ch == " ") { 
			    	retValue = retValue.substring(1, retValue.length);
			    	ch = retValue.substring(0, 1);
			    }
			    ch = retValue.substring(retValue.length-1, retValue.length);
			    while (ch == " ") { 
			    	retValue = retValue.substring(0, retValue.length-1);
			    	ch = retValue.substring(retValue.length-1, retValue.length);
			    }
			    while (retValue.indexOf("  ") != -1) {
			    	retValue = retValue.substring(0, retValue.indexOf("  ")) + retValue.substring(retValue.indexOf("  ")+1, retValue.length); 
			    }
			    return retValue; 
			}
			
		</script>
		<script type="text/javascript" src="/xtra/js/CVFunctions.js"></script>
		<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>		
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<H1 style="font-size:160%; font-weight:bold; font-color:#000000;">CV redigering - <%=CreateSOLink(SUPEROFFICE_XIS_CONSULTANT_URL, "", "VikarVis.asp?VikarID=" & lngVikarID, Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn"), "Vis vikar " & Cons.DataValues("Fornavn") & " " & Cons.DataValues("Etternavn") )%></h1>
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
				<!--
				<span class="menu2" id="4" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyNokkelKvalifikasjoner.asp?VikarID=<%=lngVikarID%>">Kandidatpresentasjon</a></span>
				-->
				<span class="menu2" id="5" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyUtdannelse.asp?VikarID=<%=lngVikarID%>">Utdannelse</a></span>
				<span class="menu2 active" id="6"><strong>Kurs</strong></span>				
				<span class="menu2" id="7" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyPraksis.asp?VikarID=<%=lngVikarID%>">Yrkeserfaring</a></span>
				<span class="menu2" id="8" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyAndreOppl.asp?VikarID=<%=lngVikarID%>">Kjernekompetanse</a></span>
				<span class="menu2" id="9" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVnyReferanser.asp?VikarID=<%=lngVikarID%>">Referanser</a></span>
				<span class="menu2" id="10" onMouseOver="menu2Over(this.id);" onMouseOut="menu2Out(this.id);"><a href="vikarCVvis.asp?VikarID=<%=lngVikarID%>">&nbsp;<img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Generere CV</a></span>
			</div>
			<div class="content">
				<form name="CVForm" action="vikarCVnyCourses.asp?VikarID=<%=lngVikarID%>" method="post">
					<input type="hidden" id="rerun" name="rerun" value="1">
					<input Type="hidden" Name="pbnDataAction" Value="lagre">
					<input type="hidden" id="course" name="course">					
					<table width="96%" class="layout" cellpadding="0" cellspacing="0">
						<col width="80%">
						<col width="20%">
					<tr>
						<td>
						<table width="100%" border="0" cellpadding='0' cellspacing='0'>
							<tr>
								<td width="100">Fra (mm/åååå):</td>
								<td width="30"><input type="text" name="frommonth" size="2" maxlength="2"
									 value="<%=objFromMonth%>" class="mandatory"></td>
								<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
								<td width="60"><input type="text" name="fromyear" size="4" maxlength="4"
									value="<%=objFromYear%>" class="mandatory">
								</td>
								<td width="100">Til (mm/åååå):</td>
								<td width="30"><input type="text" name="tomonth" size="2" maxlength="2"
									value="<%=objToMonth%>" class="mandatory">
								</td>
								<td width="10">&nbsp;&nbsp;/&nbsp;&nbsp;</td>
								<td width="60"><input type="text" name="toyear" size="4" maxlength="4"
									value="<%=objToYear%>" class="mandatory">
								</td>
								<td width="200">&nbsp;</td>
							</tr>
							<tr>
								<td width="100">Sted:</td>
								<td colspan="8"><input type="text" name="place" size="40" maxlength="50"
									<%if courseid <> 0 then%> value="<%=objPlace%>"
									<%else%> value="<%=utdPlace%>" <%end if%> class="mandatory">
								</td>
							</tr>
							<tr>
								<td width="100">Linje:</td>
								<td colspan="8"><input type="text" name="title" size="40" maxlength="50"
									<%if courseid <> 0 then%> value="<%=objTitle%>"
									<%else%> value="<%=utdTitle%>" <%end if%> class="mandatory">
								</td>
							</tr>
							<tr>
								<td width="100">&nbsp;</td>
								<td colspan="8">&nbsp;</td>
							</tr>
							<tr>							
								<td width="100">Kommentarer:</td>
								<td colspan="8">
									<%
										if(courseid<>0)then 
									%>
									<%						
						
						
						editor.Text = objDescription						
						
						%>									
							
								    <%else%>
								    <%
								    	editor.Text = utdComment
								    %>
								    <%end if%>
								    <%
									editor.Draw()								
				    			%>
								</td>
							</tr>
							
					</table>
					</td>
				<td>
					<%if feilmelding <> "" then%>
						<p class="warning"><%=feilmelding%></p>
					<%end if%>&nbsp;
				</td>
				</tr>
				</table>
				<br/><br/>
				<span class="menuInside" style="margin-left:104px;" title="Lagre informasjonen"><a href="#" onClick="javascript:lagre_data(<%=courseid%>);"><img src="images/icon_save.gif" width="18" height="15" alt="" border="0" align="absmiddle">Lagre</a></span>
				<div class="listing">
				<table cellpadding='0' cellspacing='1'>
					<tr>
						<th colspan="7">Registrerte kurs</th>
					</tr>
					<tr>
						<th>Fra</th>
						<th>Til</th>
						<th>Sted</th>
						<th>Linje</th>
						<th>Kommentarer</th>
						<th class="center">Endre</th>
						<th class="center">Slette</th>
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

							if (isnull(strToMonth)) then
								strToMonth 	= cint(0)						
							end if	
	
							if (isnull(strToYear)) then
								strToYear 	=  cint(0)						
							end if
	
							if (strFromMonth < 10) then
								strFromMonth = "0" & strFromMonth
							end if
							strFromPeriod = strFromMonth & "/" & strFromYear
	
							if (strToMonth < 10) and (strToMonth >= 1) then
								strToMonth = "0" & strToMonth
							end if
	
							if (strToMonth = 0) and (strToYear = 0) then
								strToPeriod = "d.d."
							else
								if(strToYear = 0) then
									strToPeriod = strToMonth & "/????"
								else
									strToPeriod = strToMonth & "/" & strToYear
								end if
							end if
						%>
						<tr>
							<td><%=strFromPeriod%></td>
							<td><%=strToPeriod%></td>
							<td><%=strPlace%></td>
							<td><%=strTitle%></td>
							<td><%=strDescription%>&nbsp;</td>							
							<td class="center"><a href="javascript:edit_data(<%=strDataId%>)"><img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="Endre denne oppføringen" border="0"></a></td>
							<td class="center"><a href="javascript:slett_data(<%=strDataId%>)"><img src="/xtra/images/icon_delete.gif" width="14" height="14" alt="Slette denne oppføringen" border="0"></a></td>
						</tr>
						<%
							rsCourseList.MoveNext
							Loop
							rsCourseList.Close
						%>
					</table>
				</form>
				</div>				
				<%
				' sletter alle CV objekter...
				set cv		= nothing
				cons.CV.cleanup
				cons.cleanup
				set cons	= nothing
				%>
			</div>
		</div>
	</body>
</html>