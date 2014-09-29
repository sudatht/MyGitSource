<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<%
	dim cuid : cuid = request("cuid")
	dim lngFirmaID
	dim rsContact
	session("cuid") = cuid
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))


	if (len(cuid) > 0) then
			strSQL = "SELECT FirmaId FROM Firma WHERE SOCUID = " & cuid
			set rsContact = GetFirehoseRS(strSQL, conn)
			if(HasRows(rsContact) = true) then
				lngFirmaID = rsContact("FirmaId")
			end if
			rsContact.close
			set rsContact = nothing
	end if

	CloseConnection(Conn)
	set Conn = nothing
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en"  "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<meta name="Developer" content="Electric Farm ASA">
		<title>KontaktSubMeny</title>
		<meta name="Microsoft Border" content="none">
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" src="js/contentMenu.js"></script>
		<base target="bottom">
	</head>
	<body>
		<div class="contentMenu">
			<table cellpadding="0" cellspacing="2">
				<tr>
				<% 
				If HasUserRight(ACCESS_CUSTOMER, RIGHT_READ) Then
					%>
					<td id="menu1" class="menu" onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'>
						<a href="kunde-oppdrag.asp?cuid=<%=cuid%>" title="Vis oppdrag for denne kontakten">Oppdrag liste</a>
					</td>
					<td id="menu2" class="menu" onmouseover='menuOver(this.id);' onmouseout='menuOut(this.id);'>
						<a href="kunde-vikarer.asp?cuid=<%=cuid%>" title="Vikarer tilknyttet oppdrag på denne kontakten"><img src="/xtra/images/icon_consultants.gif" alt="" width="18" height="15" border="0" align="absmiddle">Vikar liste</a>
					</td>
					<td id="menu4" class="menu" onmouseover='menuOver(this.id);' onmouseout='menuOut(this.id);'>
						<a href="http://xis/xtra/WebUI/Customer/Terms.aspx?FirmaID=<%=lngFirmaID%>" title="Comment and invoice term edit page"><img src="/xtra/images/comment_blue.gif" alt="" width="18" height="15" border="0" align="absmiddle">Kommentar</a>
					</td>					
					<%
				End If 
				If HasUserRight(ACCESS_CUSTOMER, RIGHT_WRITE) Then		
					%>
					<td id="menu3" class="menu" onmouseover='menuOver(this.id);' onmouseout='menuOut(this.id);'>					
					<%=CreateSOLink(SUPEROFFICE_XIS_TASK_URL, "", "oppdragNy.asp?cuid=" & request("cuid"), "<img src='/xtra/images/icon_oppdrag-Add.gif' width='18' height='15' alt='' align='absmiddle'>&#160;" & "Nytt oppdrag", "Opprette nytt oppdrag på kontakt" )%>
					</td>
					<%
				End If 
				%>
				</tr>
			</table>
		</div>
	</body>
</html>