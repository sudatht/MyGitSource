<%@ Language=VBScript %>
<%option explicit%>
<%

dim lVikarID			'as long
dim strVikarID			'as string
dim StrPeriod			'as string
dim iFromMonth			'as integer
dim iFromYear			'as integer
dim iToMonth			'as integer
dim iToYear			'as integer
dim strProduktomrade 		'as string
dim strFirst			'as string
dim ObjConsultant		'as xtraweb.consultant
dim ObjCv			'as xtraweb.cv
dim Objprodgroup		'as xtraweb.productgroup
dim ObjEducation		'as xtraweb.education
dim ObjEducations		'as xtraweb.educations
dim ObjExperience		'as xtraweb.experience
dim ObjExperiences		'as xtraweb.experiences
dim rcApproved			'as adodb.recordset


if Request.QueryString("VikarID") <> "" then
	lVikarID = Request.QueryString("VikarID")
end if

strVikarID = Cstr(lVikarID)

set ObjConsultant = Server.CreateObject("XtraWeb.Consultant")
ObjConsultant.XtraConString = Application("Xtra_intern_ConnectionString")
ObjConsultant.GetConsultant(lVikarID)


set ObjCv = ObjConsultant.CV
ObjCv.XtraConString = Application("Xtra_intern_ConnectionString")
ObjCv.XtraDataShapeConString = Application("ConXtraShape")
ObjCv.Refresh


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<meta name="generator" content="Microsoft Visual Studio 6.0">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<body>
	<div class="pageContainer" id="pageContainer">
<table cellpadding='0' cellspacing='0'>
	<tr>
		<td class="right"><img src="<%=Application("XtraImgAddress")%>logo1.gif" width="300" height="74" alt="" border=""></td>
	</tr>
</table>

<h1>Curriculum vitae - CV</h1>

<table cellpadding='0' cellspacing='0'>
	<tr>
		<td colspan="2"><strong>Personalia</strong></td>
	</tr>
	<tr>
		<td>Navn: </td>
		<td><%=ObjConsultant.DataValues("Fornavn")%>&nbsp;<%=ObjConsultant.DataValues("Etternavn")%></td>
	</tr>
	<tr>
		<td>Fødselsdato:</td>
		<td><%=ObjConsultant.DataValues("Foedselsdato")%></td>
	</tr>
</table>
<p>
<%
		strProduktomrade = ""
		strFirst = ""

		set Objprodgroup = Server.CreateObject("XtraWeb.productgroup")
		set rcApproved = Objprodgroup.GetAllApproved(ObjCv.XtraConString, ObjCv.datavalues("cvid").value)
		set Objprodgroup = nothing
		if (not rcApproved.EOF ) then
		%>
		<table cellpadding='0' cellspacing='0'>
		<tr>
			<td colspan="2"><strong>Produktkompetanse*</strong></td>
		</tr>
		<%
			while (not rcApproved.EOF)
				if (rcApproved.fields("Produktomrade").value <> strProduktomrade) then
					if (strProduktomrade<>"") then
						response.write "</td></tr><tr><td height=""8""></td></tr>"
						strFirst = ""
					end if
					strProduktomrade = rcApproved.fields("Produktomrade").value
					response.write "<tr><td><i>" & strProduktomrade & "</i></td></tr><tr><td>"
				end if
				if rcApproved.fields("rangering").value > 0 then
					response.write strFirst & rcApproved.fields("ktittel").value & "&nbsp;<sup><font style='COLOR: black; FONT-FAMILY: ariel; FONT-SIZE: 7pt;'>(" & rcApproved.fields("rangering").value & ")</sup>"
				else
					response.write strFirst & rcApproved.fields("ktittel").value
				end if
				if (strFirst="") then
					strFirst = ", &nbsp;&nbsp;"
				end if
				rcApproved.MoveNext
			wend
			if (strProduktomrade<>"") then
				response.write "</td></tr>"
			end if
		%>
	</table>
	<%
	end if
	set rcApproved = nothing
	%>

<p>

<table cellpadding='0' cellspacing='0'>
	<%
	set ObjEducation		= Server.CreateObject("XtraWeb.Education")
	set ObjEducations 		= ObjCv.Educations
	if ObjEducations.Count > 0 then
	%>
	<tr>
		<td>
			<table cellpadding='0' cellspacing='0'>
				<tr>
					<td colspan="4"><strong>Utdannelse</strong></td>
				</tr>

				<%
				for each ObjEducation in ObjEducations
					with ObjEducation.datavalues
						iFromMonth 	= .item("FromMonth").value
						iFromYear 	= .item("FromYear").value
						iToMonth 	= .item("ToMonth").value
						iToYear 	= .item("ToYear").value

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
					<td><%=StrPeriod%></td>
					<td><strong><%=.item("Place")%></strong></td>
				</tr>
				<tr>
					<td></td>
					<td><i><%=.item("Title")%></i></td>
				</tr>
				<tr>
					<td></td>
					<td><%=.item("Description")%></td>
				</tr>
				<%
					end with
				next
				%>
			</table>
		</td>
	</tr>
	<%
	end if
	set ObjEducations = nothing
	set ObjEducation = nothing
	set ObjExperience = Server.CreateObject("XtraWeb.Experience")
	set ObjExperiences = ObjCv.Experiences
	if ObjExperiences.Count > 0 then
	%>
	<tr>
		<td>
			<table cellpadding='0' cellspacing='0'>
				<tr>
					<td colspan="4"><strong>Praksis</strong></td>
				</tr>

				<%
				for each ObjExperience in ObjExperiences
					with objexperience.datavalues
						iFromMonth 	= .item("FromMonth").value
						iFromYear 	= .item("FromYear").value
						iToMonth 	= .item("ToMonth").value
						iToYear 	= .item("ToYear").value

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
						<td><%=StrPeriod%></td>
						<td><strong><%=.item("Place")%></strong></td>
					</tr>
					<tr>
						<td></td>
						<td><i><%=.item("Title")%></i></td>
					</tr>
					<tr>
						<td></td>
						<td><%=.item("Description")%></td>
					</tr>
					<%
					end with
				next
				%>
			</table>
		</td>
	</tr>
	<%
	end if
	set ObjExperiences = nothing
	set ObjExperience = nothing
	%>
</table>


<h1>Nivåer ved kompetansekartlegging</h1>
<p>Generelt:</p>
<ol>
	<li>Kjennskap til produktet, ingen erfaring</li>
	<li>Kjennskap til produktet, lite erfaring</li>
	<li>Noe kompetanse, noe erfaring. Tilstrekkelig til å bruke i arbeidssammenheng</li>
	<li>God kompetanse, jevnlig erfaring (ukentlig)</li>
	<li>Generelt meget god kompetanse og erfaring</li>
	<li>Dyptgående kjennskap og erfaring.</li>
</ol>
    </div>
</body>
</html>

<%
' sletter alle CV objekter...
set ObjCv = nothing
ObjConsultant.CV.cleanup
ObjConsultant.cleanup
set ObjConsultant = nothing
%>

