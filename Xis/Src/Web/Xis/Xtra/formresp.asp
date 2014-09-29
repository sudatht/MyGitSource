<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim  strFilePath
	dim  strVikarID
	dim	 oFSO
	dim  oUploader
	
	'---- Parameter validation
	if (len(trim(Request("vikarid")))=0) then
		AddErrorMessage("ingen vikar spesifisert.")
		call RenderErrorMessage()	
	else
		strVikarID = Request.QueryString("vikarid")
	end if
	
	'---- Initialize opload path
	strFilePath = Application("ConsultantFileRoot") & strVikarID & "\"

	dim util
	'Impersonate to access network resources
	set  util = Server.CreateObject("XisSystem.Util")
	call util.Logon()
		
	Set oUploader = Server.CreateObject("SoftArtisans.FileUp")
	oUploader.Path = strFilePath
	if (oUploader.IsEmpty = false) then
	'Make sure there is a file specified
		oUploader.Save
	end if
	'Clean up
	set oFSO = nothing
	set oUploader = nothing
	
	util.Logoff
	Set util = Nothing					
%>
<HTML>
	<HEAD>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<TITLE>Fil lastes opp..</TITLE>
	</HEAD>
	<body onload="javascript:window.opener.location='vikarVis.asp?vikarid=<%=strVikarID%>';javascript:self.close()"> 
	</body>
</html>
