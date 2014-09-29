<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>


<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->

<%
	dim objCon
	dim rsVer 
	dim strSql
	dim sreVersion
	dim strDescrip
	dim strIssues
	dim strFeatures
	
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Set objCon = GetConnection(GetConnectionstring(XIS, ""))
	
	strSql = "SELECT * FROM Version WHERE Date=(SELECT MAX(Date) AS max_date From Version AS v)"
	set rsVer = GetFirehoseRS(strSql, objCon)

	strVersion       = rsVer("Version").Value
        strDescrip = rsVer("Description").Value
        strIssues       = rsVer("Issues").Value
        strFeatures = rsVer("Features").Value
        
        rsVer.Close
        set rsVer = Nothing

	CloseConnection(objCon)
	set objCon = nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
</style>
<link href="/xtra/css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div style="width:450px; height:330px;">
<div class="Logo_white_outline">
<div style="padding-top:45px; padding-right:15px;">Version:  <%Response.write(strVersion)%></div>
</div>
<div class="shade">&nbsp;</div>
<div style="padding-left:15px;">
<table  style="width:400px;" border="0" cellspacing="8" cellpadding="0">
  <tr>
    <td style="width:100px; height:20px;">Issues Fixed</td>
    <td>: <%Response.write(strIssues)%></td>
  </tr>
  <tr>
    <td>New Features</td>
    <td>: <%Response.write(strFeatures)%></td>
  </tr>
  <tr>
    <td>Comments</td>
    <td>: <%Response.write(strDescrip)%></td>
  </tr>
</table>
</div>
</div>
 </body>
</html>
