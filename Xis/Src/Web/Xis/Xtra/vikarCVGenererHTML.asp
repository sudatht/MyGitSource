<%@  language="VBScript" %>
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
	dim strProductAreas
	dim strPrevNivaa
	dim rcAddress

	const C_HOMEADRESS_TYPE = 1

	if Request("VikarID") <> "" then
		lngVikarID = Request("VikarID")
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
	
	Response.Write lngVikarID
	
	'Creates a DB connection to retrieve the kundepresentasjon list
	Set kundepresentasjonConn = GetClientConnection(GetConnectionstring(XIS, ""))
	Set rskundepresentasjon	= GetFirehoseRS("SELECT convert(varchar(8000),kundepresentasjon) as kundepresentasjon from vikar where vikarid =" & lngVikarID, kundepresentasjonConn)	
	kundepresentasjon = rskundepresentasjon("kundepresentasjon")
	rskundepresentasjon.Close
	
	'Creates a DB connection to retrieve the Vikar Country
	Set countryConn = GetClientConnection(GetConnectionstring(XIS, ""))	
	set rsCountry = GetFirehoseRS("select printablename = case when c.printablename is null  then 'Ingen valgt' else c.printablename end from vikar v left outer join country c on c.countryid = v.country where v.vikarid = " & lngVikarID, countryConn)	
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

	response.Clear
	if request("dest") = "Application" then
		'Response.ContentType ="application/doc"
		Response.ContentType ="application/msword"
		Response.AddHeader "Content-Disposition", "inline;filename=CV_" & cons.DataValues("Etternavn") & "_" & cons.DataValues("Fornavn") &  ".htm"
	elseif request("dest") = "Disk" then
		Response.ContentType ="application/html"
		Response.AddHeader "Content-Disposition", "attachment;filename=CV_" & cons.DataValues("Etternavn") & "_" & cons.DataValues("Fornavn") & ".htm"
	elseif request("dest") = "Browser" then
		'Response.ContentType ="application/html"
		'Response.AddHeader "Content-Disposition", "attachment;filename=CV_" & cons.DataValues("Etternavn") & "_" & cons.DataValues("Fornavn") & ".htm"
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
    <title>Curriculum Vitae</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <style>
		BODY		{font-size: .7em; margin: 0; padding: 0 10px 10px 10px;}
		BODY, TABLE	{font-family: Arial, sans-serif;}
		H1, H3	{margin:0 0 0 0; color:#666666; background:transparent;}
		
		H1			{font-size:2.4em; font-weight:normal;}
		P           {font-size:1em; margin:0 0 1em 0;}
		TABLE 		{width:100%;}
		IMG			{border:none;}
		.newWindow	{color:#000000; background:#ffffff; padding:0 2em 2em 2em;}
		
		.newWindow .logo	{position:absolute; top:0; left:84%; width:63;}
		.showCV TD, .showCV TH	{border-bottom:1px solid #cccccc;}
		.showCV TH  {font-size: 1.1em;}
		.showCV TD  {font-size: 1.1em;}
		.showCV TABLE	{width:100%;}
		.showCV TR	{vertical-align:top;}
		.left		{text-align:left;}
		.center		{text-align:center;}
		.right		{text-align:right;}
		.warning	{color:#ff0000; background:transparent;}
		.top		{vertical-align:top;}
	</style>
</head>
<body class="newWindow">
    <table border="0" width="100%" id="Table1">
        <tr>
            <td align="right">
                <img align="right" src="http://intern.xtra.no/portals/0/site_images/xtra_logo_cv.png"
                    width="150px" height="44px" alt="Xtra logo"></td>
        </tr>
    </table>
    <h1 style="color: Green; font-weight: bold">
        CV</h1>
    <div class="showCV">
        <div class="row">
            <div>
                <img src="http://intern.xtra.no/portals/0/site_images/header.png" alt="" height="25" width="100%"></div>
        </div>
        <br />
        <br />
        <table id="Table2">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
            </tr>
            <%
			if (request("chkName")="1") then
            %>
            <tr>
                <th align="left">
                    Navn:</th>
                <td>
                    <%=cons.DataValues("Fornavn") & " " & cons.DataValues("Etternavn")%>
                </td>
            </tr>
            <%
			end if

			set rcAddress = objCon.execute("exec [dbo].[GetAddressForConsultantByType] " & lngVikarID & "," & "'" & C_HOMEADRESS_TYPE & "'")
			If objCon.Errors.Count > 0 Then
				' Error message
				Call SqlError()
			End if
			if (not rcAddress.EOF ) then
				if (request("chkAdress")="1") then
					if ( len(rcAddress.fields("Adresse"))>0) then
            %>
            <tr>
                <th align="left">
                    Adresse:</th>
                <td>
                    <%=rcAddress.fields("Adresse")%>
                </td>
            </tr>
            <%
					end if
				end if
				if (request("chkHomeAdress")="1") then
					if ( len(trim(rcAddress.fields("Postnr")))>0 and len(trim(rcAddress.fields("Poststed")))>0 ) then
            %>
            <tr>
                <th align="left">
                    Postadresse:</th>
                <td>
                    <%=rcAddress.fields("Postnr") & " " & rcAddress.fields("Poststed")%>
                </td>
            </tr>
            <%
					end if
				end if
			end if
			set rcAddress = nothing
			if (request("chkPhone")="1") then
				if (  len(trim(cons.DataValues("Telefon")))>0 ) then
            %>
            <tr>
                <th align="left">
                    Telefon:</th>
                <td>
                    <%=cons.DataValues("Telefon")%>
                </td>
            </tr>
            <%
				end if
				if (  len(trim(cons.DataValues("MobilTlf")))>0 ) then
            %>
            <tr>
                <th align="left">
                    Mobil:</th>
                <td>
                    <%=cons.DataValues("MobilTlf")%>
                </td>
            </tr>
            <%
				end if
			end if
			if (request("chkBirthDate")="1") then
				if (  len(trim(cons.DataValues("Foedselsdato")))>0 ) then
            %>
            <tr>
                <th align="left">
                    Fødselsdato:</th>
                <td>
                    <%=cons.DataValues("Foedselsdato")%>
                </td>
            </tr>
            <%
				end if
			end if
			
			if (request("chkEmail")="1") then
				if (  len(trim(cons.DataValues("epost")))>0 ) then
            %>
            <tr>
                <th align="left">
                    Epost:</th>
                <td>
                    <%=cons.DataValues("epost")%>
                </td>
            </tr>
            <%
				end if
			end if
            %>
            <%
			if (request("chkCountry")="1") then				
            %>
            <tr>
                <th align="left">
                    Nasjonalitet:</th>
                <td>
                    <%=countryname%>
                </td>
            </tr>
            <%			
			end if
            %>
        </table>
        <br />
        <br />
        <br />
        <br />
        <%
		if (request("chkKeyQualifcations")="1") then
			
        %>
        <h2 style="color: Green; font-weight: normal">
            Kandidatpresentasjon</h2>
			
		<table id="Table3" style="font-family: Arial;">
            <tr>
                <td>	
        <%=kundepresentasjon%>
		</td>
		</tr>
</table>
        <%
			
		end if
		%>
		
		<br />
        <br />
        <br />
        <br />
		
		<%
		if (request("chkOtherInformation")="1") then
			if len(trim(strAndreOpp))> 0 then
        %>
		<h2 style="color: Green; font-weight: normal">
            Kjernekompetanse</h2>
			<table id="Table3" style="font-family: Arial;">
            <tr>
                <td>
        <%=strAndreOpp%>
		</td>
		</tr>
</table>
        <%
			end if
		end if

		if (len(trim(request("chkProductAreas")))>0) then

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
        <h2 style="color: Green; font-weight: normal">
            Produktkompetanse</h2>
        <table id="Table3" style="font-family: Arial;">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
            </tr>
            <tr>
                <th align="left" nowrap>
                    Brukernivå</th>
                <th align="left">
                    Produkter</th>
            </tr>
            <%
			strNivaa  = rcApproved.fields("K_Rangering").value
			strPrevNivaa = strNivaa
			strProduct = rcApproved.fields("Ktittel").value
            %>
            <tr>
                <td align="left" nowrap>
                    <%
					if (len(trim(strNivaa))>0) then
						Response.Write strNivaa
					else
						Response.Write "(ikke spesifisert)"
					end if
                    %>
                </td>
                <td>
                    <%
				while (not rcApproved.EOF)
					if (strPrevNivaa<>rcApproved.fields("K_Rangering").value) then
                    %>
                </td>
            </tr>
            <tr>
                <td>
                    <%
						if (len(trim(rcApproved.fields("K_Rangering").value))>0) then
							Response.Write strNivaa
						else
							Response.Write "(ikke spesifisert)"
						end if
                    %>
                </td>
                <td>
                    <%
					end if
					response.write rcApproved.fields("Ktittel").value
					strPrevNivaa = rcApproved.fields("K_Rangering").value
					rcApproved.MoveNext
					if (not rcApproved.EOF) then
						strNivaa = rcApproved.fields("K_Rangering").value
						if (strPrevNivaa = strNivaa) then
							response.write ", "
						end if
					end if
				wend
                    %>
                </td>
            </tr>
        </table>
        <%
			end if
			rcApproved.close
			set rcApproved = nothing
		end if
		
        %>
        <%

		if (len(trim(request("chkEducation")))>0) then
			set ObjEducation	= Server.CreateObject("XtraWeb.Education")
			set ObjEducations 	= ObjCv.Educations
			if ObjEducations.Count > 0 then
        %>
        <br />
        <br />
        <br />
        <br />
        <h2 style="color: Green; font-weight: normal">
            Utdannelse</h2>
        <table id="Table4" style="font-family: Arial;">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
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
                <td nowrap align="left" valign="top">
                    <%=StrPeriod%>
                </td>
                <td>
                    <strong>
                        <%=ObjEducation.datavalues.item("Place")%>
                    </strong>
                    <br/>
                    <i>
                        <%=ObjEducation.datavalues.item("Title")%>
                    </i>
                    <br/>
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
		end if
        %>
        <%

		if (len(trim(request("chkCourses")))>0) then			
	
			if(HasRows(rsCourseList)) then	
        %>
           <br />
        <br />
        <br />
        <br />
        <h2 style="color: Green; font-weight: normal">
            Kurs</h2>
        <table id="Table5" style="font-family: Arial;">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
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
                <td nowrap align="left" valign="top">
                    <%=StrPeriod%>
                </td>
                <td>
                    <strong>
                        <%=strPlace%>
                    </strong>
                    <br/>
                    <i>
                        <%=strTitle%>
                    </i>
                    <br/>
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
end if
        %>
        <%

		if (len(trim(request("chkExperience")))>0) then
			set ObjExperience = Server.CreateObject("XtraWeb.Experience")
			set ObjExperiences = ObjCv.Experiences
			if ObjExperiences.Count > 0 then
        %>
           <br />
        <br />
        <br />
        <br />
        <h2 style="color: Green; font-weight: normal">
            Yrkeserfaring</h2>
        <table id="Table6" style="font-family: Arial;">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
            </tr>
            <%
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
                <td align="left" nowrap valign="top">
                    <%=StrPeriod%>
                </td>
                <td>
                    <strong>
                        <%=objexperience.datavalues.item("Place")%>
                    </strong>
                    <br/>
                    <i>
                        <%=objexperience.datavalues.item("Title")%>
                    </i>
                    <br/>
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
		end if
		%>
         
        
		<%

		if (request("chkReferences")="1") then
			set ref	= Server.CreateObject("XtraWeb.Reference")
			set allRef = ObjCv.References
			if allRef.Count > 0 then
        %>
        <br />
        <br />
        <br />
        <br />
        <h2 style="color: Green; font-weight: normal">
            Referanser</h2>
        <table id="Table7" style="font-family: Arial;">
            <tr>
                <td width="20%">
                </td>
                <td width="80%">
                </td>
            </tr>
            <%
				for each ref in allRef
            %>
            <tr>
                <th align="left" nowrap valign="top">
                    Navn, tittel</th>
                <td align="left">
                    <%=ref.DataValues("Name")%>
                    ,<%=ref.DataValues("Title")%></td>
            </tr>
            <tr>
                <th align="left" nowrap>
                    Kontakt</th>
                <td align="left">
                    <%=ref.DataValues("Firma")%>
                </td>
            </tr>
            <tr>
                <th align="left" nowrap>
                    Telefon</th>
                <td align="left">
                    <%=ref.DataValues("Tel")%>
                </td>
            </tr>
            <tr>
                <th align="left" nowrap>
                    Kommentar</th>
                <td align="left">
                    <%=ref.DataValues("Comment")%>
                    &nbsp;</td>
            </tr>
            <%
				next
            %>
        </table>
        <%
			end if
			set allRef = nothing
			set ref = nothing
		end if
        %>
        </table>
        <br />
        <br />
        <br />
        <br />
        <div class="row">
            <div>
                <img src="http://intern.xtra.no/portals/0/site_images/header.png" alt="" height="25" width="100%"></div>
        </div>
    </div>
</body>
<%
' sletter alle CV objekter...
set ObjCv		= nothing
cons.CV.cleanup
cons.cleanup
set cons	= nothing

objCon.close
set objCon = nothing
%>
</html>
