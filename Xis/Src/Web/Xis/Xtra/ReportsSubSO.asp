<%@ Language=VBScript%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<%
	dim targetPage : targetPage = "rapporter\Rapporter.asp"

	dim subPageUrl : subPageUrl = CreateSubPageURL(targetPage)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<style type="text/css">
			frameset, frame	{margin:0 0 0 0; padding:0 0 0 0;}
			frameset	{border:none;}
			frame		{width:100%; border:none;}
		</style>
		<title>x|is</title>
	</head>
	<frameset rows="21,*" framespacing="0" frameborder="0">
		<frame src="ReportsSub.asp" name="middle" id="middle" scrolling="no" noresize>
		<frame src="<%=subPageUrl%>" name="bottom" id="bottom" scrolling="yes">
	</frameset>
</html>
