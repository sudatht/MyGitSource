<%@ Language=VBScript%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes/SuperOffice.Page.Navigation.inc"-->
<%
	dim cuid : cuid = request("cuid")
	dim targetPage : targetPage = "kunde-oppdrag.asp"
	dim subPageUrl : subPageUrl = CreateSubPageURL(targetPage)	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<meta http-equiv="Content-Style-Type" content="text/css">
		<meta http-equiv="Content-Script-Type" content="text/javascript">
		<style type="text/css">
			frameset, iframe	{margin:0 0 0 0; padding:0 0 0 0;}
			frameset, iframe	{border:none;}
			iframe		{width:100%; border:none;}
		</style>
		<script type="text/javascript">
			function setHeight() {
				document.getElementById("bottom").height = document.getElementById("bottom").contentWindow.document.body.scrollHeight + "px";
			}
		</script>
		<title>x|is</title>
	</head>

		<iframe src="OppdragKundeSub.asp?cuid=<%=cuid%>" name="middle" id="middle" scrolling="no" height="25px" frameBorder="0"></iframe>
		<iframe src="<%=targetPage%>?cuid=<%=cuid%>" name="bottom" id="bottom" scrolling="no" frameBorder="0" onload="setHeight();"></iframe>
		

</html>
