<%@ Language=VBScript %>
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Renderfunctions.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	dim exp
	dim strKeyQualifications
	dim fulltime
	dim beskrivelse
	dim lngVikarID
	dim rcApproved
	dim ObjEducation		'as xtraweb.education
	dim ObjEducations		'as xtraweb.educations
	dim ObjExperience		'as xtraweb.experience
	dim ObjExperiences		'as xtraweb.experiences
	dim ObjCv				'as xtraweb.cv
	dim strNivaa
	dim strProduct
	dim strAndreOpp
	dim strPrevNivaa
	dim strKommentar
	dim strKtittel
	dim strDisplayNivaa

	const C_HOMEADRESS_TYPE = 1
		
	if Request.QueryString("VikarID") <> "" then
		lngVikarID = Request.QueryString("VikarID")
	end if

	set cons = Server.CreateObject("XtraWeb.Consultant")
	cons.XtraConString = Application("Xtra_intern_ConnectionString")
	cons.GetConsultant(lngVikarID)

	set ObjCv = cons.CV
	ObjCv.XtraConString = Application("Xtra_intern_ConnectionString")
	ObjCv.XtraDataShapeConString = Application("ConXtraShape")
	ObjCv.Refresh
	
	'Creates a DB connection to retrieve the CV record
	Set cvConn = GetClientConnection(GetConnectionstring(XIS, ""))	
	set rsCV = GetFirehoseRS("SELECT max(CVId) AS CVId FROM CV WHERE ConsultantId =" & lngVikarID, cvConn)	
	vikarcvid = rsCV("CVId")	
	rsCV.Close			

	'Creates a DB connection to retrieve the course list
	Set listConn = GetClientConnection(GetConnectionstring(XIS, ""))
	Set rsCourseList	= GetFirehoseRS("SELECT DataId,CVId,Place,Country,Description,Title,FieldType,FromMonth,FromYear,ToMonth,ToYear FROM CV_Data WHERE FieldType = 'COU' and CVId =" & vikarcvid & " ORDER BY FromYear desc, FromMonth desc", listConn)	
	
	'Creates a DB connection to retrieve the kundepresentasjon list
	Set kundepresentasjonConn = GetClientConnection(GetConnectionstring(XIS, ""))
	Set rskundepresentasjon	= GetFirehoseRS("SELECT convert(varchar(8000),kundepresentasjon) as kundepresentasjon from vikar where vikarid =" & lngVikarID, kundepresentasjonConn)	
	kundepresentasjon = rskundepresentasjon("kundepresentasjon")
	rskundepresentasjon.Close
			
	'Creates a DB connection to retrieve the Vikar Country
	Set countryConn = GetClientConnection(GetConnectionstring(XIS, ""))	
	strCountry = "select printablename = case when c.printablename is null  then 'Ingen valgt' else c.printablename end from vikar v left outer join country c on c.countryid = v.country where v.vikarid = " & lngVikarID
	'Response.Write(strCountry)
	set rsCountry = GetFirehoseRS(strCountry, countryConn)	
	countryname = rsCountry("PrintableName")	
	rsCountry.Close			
	
	if (ObjCv.Datavalues.Count > 1) then
		if (not isnull(ObjCv.datavalues("Key_Qualifications"))) then
			strKeyQualifications = ObjCv.datavalues("Key_Qualifications").value
		else
			strKeyQualifications = ""
		end if

		if (not isnull(ObjCv.datavalues("Other_information"))) then
			strAndreOpp = ObjCv.datavalues("Other_information").value
		else
			strAndreOpp = ""
		end if
	end if

	' Open database connection
	Set objCon = Server.CreateObject("ADODB.Connection")
	objCon.ConnectionTimeOut = Session("Xtra_ConnectionTimeout")
	objCon.CommandTimeOut = Session("Xtra_CommandTimeout")
	objCon.Open Session("Xtra_ConnectionString"), Session("Xtra_RuntimeUserName"), Session("Xtra_RuntimePassword")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Vikarprofil</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<meta http-equiv="Content-Style-Type" content="text/css">
	<meta http-equiv="Content-Script-Type" content="text/javascript">
	<link type="text/css" rel="stylesheet" href="/xtra/css/CV.css" title="default style">
	<link type="text/css" rel="stylesheet" href="/xtracss/print.css" title="default print style" media="print">
	
	<style type="text/css">
	    H1, H3	{margin:0 0 0 0; color:#666666; background:transparent;}
		H1			{font-size:2.4em; font-weight:normal;}
		.showCV TD, .showCV TH	{border-bottom:1px solid #cccccc;}
		.showCV TH  {font-size: 1.1em;}
		.showCV TD  {font-size: 1.1em;}
		.showCV TABLE	{width:100%;}
		.showCV TR	{vertical-align:top;}
	</style>
</head>
<body class="newWindow">

	<table border="0" width="100%">
		<tr>
		<td><h1 style="color: Green; font-weight: bold">CV</h1></td>
			<td align="right"><img align="right" src="http://intern.xtra.no/portals/0/site_images/xtra_logo_cv.png" width="150px" height="44px" alt="Xtra logo"></td>
		</tr>
	</table>
	
	<div><img src="http://intern.xtra.no/portals/0/site_images/header.png" alt="" height="25" width="100%"></div>
	<br />
	<br />
	<br />
	<br />
	<div class="showCV">
		
		<table>
			<tr>
				<td width="20%"></td>
				<td width="80%"></td>
			</tr>
			<tr>
				<th align="left">Navn:</th>
				<td><%=cons.DataValues("Fornavn") & " " & cons.DataValues("Etternavn")%></td>
			</tr>
			<%
			set rcAddress = objCon.execute("exec [dbo].[GetAddressForConsultantByType] " & lngVikarID & "," & "'" & C_HOMEADRESS_TYPE & "'")
			If objCon.Errors.Count > 0 Then
				' Error message
				Call SqlError()
			End if
			if (not rcAddress.EOF ) then

				if ( len(rcAddress.fields("Adresse"))>0) then
				%>
				<tr>
					<th align="left">Adresse:</th>
					<td><%=rcAddress.fields("Adresse")%></td>
				</tr>
				<%
				end if


				if ( len(trim(rcAddress.fields("Postnr")))>0 and len(trim(rcAddress.fields("Poststed")))>0 ) then
				%>
					<tr>
						<th align="left">Postadresse:</th>
						<td><%=rcAddress.fields("Postnr") & " " & rcAddress.fields("Poststed")%></td>
					</tr>
				<%
				end if

			end if
			set rcAddress = nothing
			if (  len(trim(cons.DataValues("Telefon")))>0 ) then
			%>
			<tr>
				<th align="left">Telefon:</th>
				<td><%=cons.DataValues("Telefon")%></td>
			</tr>
			<%
			end if
			if (  len(trim(cons.DataValues("MobilTlf")))>0 ) then
			%>
			<tr>
				<th align="left">Mobil:</th>
				<td><%=cons.DataValues("MobilTlf")%></td>
			</tr>
			<%
			end if
			if (  len(trim(cons.DataValues("Foedselsdato")))>0 ) then
			%>
			<tr>
				<th align="left">Fødselsdato:</th>
				<td><%=cons.DataValues("Foedselsdato")%></td>
			</tr>
			<%
			end if
			%>
			
			<%
			if (  len(trim(cons.DataValues("epost")))>0 ) then
			%>
			<tr>
				<th align="left">Epost:</th>
				<td><%=cons.DataValues("epost")%></td>
			</tr>
			<%
			end if
			%>
			
			<tr>
				<th align="left">Nasjonalitet:</th>
				<td><%=countryname%></td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		<br />
		<%
		if len(trim(strKeyQualifications))> 0 then
		%>
		<h2 style="color: Green; font-weight: bold">Kandidatpresentasjon</h2>
		<%=kundepresentasjon%>
		<%
		end if
		%>
		<br />
		<br />
		<br />
		<br />
		
		<%
		if len(trim(strAndreOpp))> 0 then
		%>
		<h2 style="color: Green; font-weight: bold">Kjernekompetanse</h2>
		<%=strAndreOpp%>
		<%
		end if
		
	
		strProductAreas = request("chkProductAreas")
		set rcApproved = objCon.execute("exec [GetCVQualificationsForConsultant] " & lngVikarID & "," & "'" & strProductAreas & "'")
		
		If objCon.Errors.Count > 0 Then
			' Error message
			Call SqlError()
		End if
		if (not rcApproved.EOF ) then
		%>
			<br />
	<br />
	<br />
	<br />
		<h2 style="color: Green; font-weight: bold">Produktkompetanse</h2>
		<table ID="Table3">
			<tr>
				<td width="20%"></td>
				<td width="40%"></td>
				<td width="40%"></td>
			</tr>
			<tr>
				<th align="left" nowrap>Brukernivå</th>
				<th align="left">Produkter</th>
				<th align="left">Kommentar</th>
			</tr>
		<%
		strNivaa  = rcApproved.fields("K_Rangering").value
		if( rcApproved.RecordCount>0 and rcApproved.RecordCount=1) then
		   strPrevNivaa = strNivaa & "1"
		Else   
		   strPrevNivaa = strNivaa
		End if   
		strProduct = rcApproved.fields("Ktittel").value
		%>
			
				<td align="left" nowrap>
				<%
				if (len(trim(strNivaa))>0) then
					strDisplayNivaa= strNivaa
				else
					strDisplayNivaa = "(ikke spesifisert)"
				end if
				%>
				
				
			<%
			while (not rcApproved.EOF)
			        strNivaa = rcApproved.fields("K_Rangering").value
			       
			        if (len(trim(rcApproved.fields("K_Rangering").value))>0) then
				    strDisplayNivaa = strNivaa
				else
				    strDisplayNivaa = "(ikke spesifisert)"
				end if	
				%>    
					
				    <tr>
						<td valign ="top" align ="Left">
						<%=strDisplayNivaa%>
						</td>
						<td valign ="top" align ="Left">
						<%=rcApproved.fields("Ktittel").value%>
					
						</td>
						<td valign ="top" align ="Left">
						<%= rcApproved.fields("Kommentar").Value %>
					
						</td>
				   </tr>
				   <%
			
			   rcApproved.MoveNext
			wend
		%>
					
					
				
		</table>
		<%
		end if
		rcApproved.close
		set rcApproved = nothing

		set rcApproved = objCon.execute("exec [GetCVProfessionsForConsultant] " & lngVikarID)
	   If objCon.Errors.Count > 0 Then
		   ' Error message
	      Call SqlError()
	   End if

		if (not rcApproved.EOF ) then
		%>
			<br />
	<br />
	<br />
	<br />
		<h2 style="color: Green; font-weight: bold">Fagkompetanse</h2>
		<table>
			<tr>
				<td width="20%" align="Left"></td>
				<td width="20%" align="Left" ></td>
				<td width="20%" align="Left"></td>
				<td width="40%" align="Left"></td>
			</tr>
			<tr>
				<th align ="Left">Type</th>
				<th align ="Left">Relevant utdannelse</th>
				<th align ="Left">Relevant erfaring</th>
				<th align ="Left">Kommentar</th>
			</tr>
		<%
			while (not rcApproved.EOF)
				%>
			<tr>
				<td  valign ="top"  align="Left" ><%=rcApproved.fields("Ktittel").value%>&nbsp;</td>
				<td  valign ="top"  align="Left"><%=rcApproved.fields("k_nivaa").value%>&nbsp;</td>
				<td  valign ="top"  align="Left"><%=rcApproved.fields("K_Erfaring").value%>&nbsp;</td>
				<td  valign ="top"  align="Left"><%=rcApproved.fields("Kommentar").value%>&nbsp;</td>
			</tr>
				<%
				rcApproved.MoveNext
			wend
		%>
					</td>
				</tr>
		</table>
		<%
		end if
		rcApproved.close
		set rcApproved = nothing
		%>

		<%
	set ObjEducation	= Server.CreateObject("XtraWeb.Education")
	set ObjEducations 	= ObjCv.Educations
	if ObjEducations.Count > 0 then
		%>
			<br />
	<br />
	<br />
	<br />
		<h2 style="color: Green; font-weight: bold">Utdannelse</h2>
		<table>
			<tr>
				<td width="20%"></td>
				<td width="80%"></td>
			</tr>
		<%
		for each ObjEducation in ObjEducations
			iFromMonth 	= ObjEducation.datavalues.item("FromMonth").value
			iFromYear 	= ObjEducation.datavalues.item("FromYear").value
			iToMonth 	= ObjEducation.datavalues.item("ToMonth").value
			iToYear 	= ObjEducation.datavalues.item("ToYear").value

			if iFromMonth < 10 then
				iFromMonth = "0" & iFromMonth
			end if
			if (iToMonth < 10) and (iToMonth >= 1) then
				iToMonth = "0" & iToMonth
			end if

			if (itomonth=0) or (iToYear=0) then
				StrPeriod = iFromMonth & "/" & iFromYear & " - d.d."
			else
				StrPeriod = iFromMonth & "/" & iFromYear & " - " & iToMonth & "/" & iToYear
			end if
			%>
			<tr>
				<td nowrap align="left" valign="top"><%=StrPeriod%></td>
				<td>
					<strong><%=ObjEducation.datavalues.item("Place")%></strong><br>
					<i><%=ObjEducation.datavalues.item("Title")%></i><br>
					<%=ObjEducation.datavalues.item("Description")%>
				</td>
			</tr>
			<%
			next
			%>
	</table>
<%
end if
set ObjEducations = nothing
set ObjEducation = nothing
%>

<%
	
	if(HasRows(rsCourseList)) then	
		%>
			<br />
	<br />
	<br />
	<br />
		<h2 style="color: Green; font-weight: bold">Kurs</h2>
		<table>
			<tr>
				<td width="20%"></td>
				<td width="80%"></td>
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
				<td nowrap align="left" valign="top"><%=StrPeriod%></td>
				<td>
					<strong><%=strPlace%></strong><br>
					<i><%=strTitle%></i><br>
					<%=strDescription%>
				</td>
			</tr>
			<%
				rsCourseList.MoveNext
				Loop
				rsCourseList.Close
				%>
	</table>
<%
end if
%>

<%
	set ObjExperience = Server.CreateObject("XtraWeb.Experience")
	set ObjExperiences = ObjCv.Experiences
	if ObjExperiences.Count > 0 then
		%>
			<br />
	<br />
	<br />
	<br />
		<h2 style="color: Green; font-weight: bold">Yrkeserfaring</h2>
		<table>
			<tr>
				<td width="20%"></td>
				<td width="80%"></td>
				 
			</tr>			<%
			for each ObjExperience in ObjExperiences
				iFromMonth 	= objexperience.datavalues.item("FromMonth").value
				iFromYear 	= objexperience.datavalues.item("FromYear").value
				iToMonth 	= objexperience.datavalues.item("ToMonth").value
				iToYear 	= objexperience.datavalues.item("ToYear").value

				if iFromMonth < 10 then
					iFromMonth = "0" & iFromMonth
				end if
				if (iToMonth < 10) and (iToMonth >= 1) then
					iToMonth = "0" & iToMonth
				end if

				if (itomonth=0) or (iToYear=0) then
					StrPeriod = iFromMonth & "/" & iFromYear & " - d.d."
				else
					StrPeriod = iFromMonth & "/" & iFromYear & " - " & iToMonth & "/" & iToYear
				end if
			%>
			<tr>
				<td nowrap align="left" valign="top"><%=StrPeriod%></td>
				<td>
					<strong><%=objexperience.datavalues.item("Place")%></strong><br>
					<i><%=objexperience.datavalues.item("Title")%></i><br>
					<%=objexperience.datavalues.item("Description")%>
				</td>
			</tr>
				<%
			next
			%>
		</table>
	<%
	end if
	set ObjExperiences = nothing
	set ObjExperience = nothing

	set ref	= Server.CreateObject("XtraWeb.Reference")
	set allRef = ObjCv.References
	if allRef.Count > 0 then
	%>
		<br />
	<br />
	<br />
	<br />
	<h2 style="color: Green; font-weight: bold">Referanser</h2>
	<table>
		<tr>
			<td width="20%"></td>
			<td width="80%"></td>
		</tr>
		<%
		for each ref in allRef
		%>
		<tr>
			<th>Navn, tittel</th>
			<td><%=ref.DataValues("Name")%>,<%=ref.DataValues("Title")%></td>
		</tr>
		<tr>
			<th>Kontakt</th>
			<td><%=ref.DataValues("Firma")%></td>
		</tr>
		<tr>
			<th>Telefon</th>
			<td><%=ref.DataValues("Tel")%></td>
		</tr>
		<tr>
			<th>Kommentar</th>
			<td><%=ref.DataValues("Comment")%></td>
		</tr>
		<%
		next
		%>
	</table>
	<%
	end if
	set allRef = nothing
	set ref = nothing
	%>
		</table>
		<br />
		<br />
		<br />
		<br />
	</div>
	<div>
     <img src="http://intern.xtra.no/portals/0/site_images/header.png" alt="" height="25" width="100%"></div>
</body>
<%
objCon.close
set objCon = nothing
' sletter alle CV objekter...
set ObjCv		= nothing
cons.CV.cleanup
cons.cleanup
set cons	= nothing
%>
</html>
