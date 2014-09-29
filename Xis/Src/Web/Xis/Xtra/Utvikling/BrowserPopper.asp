<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit

	dim blnIsPosted :  blnIsPosted = false
	dim strBodyOnload
	dim strRdoTarget
	dim strURL
	dim strChecked

	if (request("hdnPosted")="1")  then
		blnIsPosted = true

		strBodyOnload = "onLoad='javascript:Popwindow();' "
		strRdoTarget = request("chkCVTarget")
		strURL = "'VisNoHTML.asp?dest=" & strRdoTarget & "'"

	else
		strBodyOnload = ""
		strURL = ""
		strRdoTarget = ""
		strURL = "'VisNoHTML.asp?dest=" & strRdoTarget & "'"
	end if

%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
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
				window.open(<%=strURL%>, 'arne', 'menubar=yes,toolbar=yes,status=yes');
			}
			
		</script>		
		<title>&nbsp;Generere CV</title>
	</head>
	<body <%=strBodyOnload%>>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Pop browser window test</a></h1>
			</div>
			<div class="content">
				<form action="BrowserPopper.asp" method="POST" id="frmGenererCV" name="frmGenererCV">
					<input type="hidden" id="hdnPosted" name="hdnPosted" value="1">

					<h2>Åpne:</h2>						
					<table id="Table6">
						<tr>
							<%
							if (strRdoTarget="Application") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																								
							<td><input type="radio" class="radio" <%=strChecked%> value="Application" name="chkCVTarget" id="chkCVTarget">Word<span class="warning">&nbsp;(Åpnes som HTML, må lagres i annet format)</span></td>
						</tr>					
						<tr>
							<%
							if (strRdoTarget="Browser") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																													
							<td><input type="radio" class="radio" <%=strChecked%> checked value="Browser" name="chkCVTarget" id="Radio1">Nettleser</td>							
						</tr>					
						<tr>																								
							<%
							if (strRdoTarget="Disk") then
								strChecked = " checked "
							else
								strChecked = ""
							end if
							%>																													
							<td><input type="radio" class="radio" <%=strChecked%> value="Disk" name="chkCVTarget" id="Radio2">disk</td>
						</tr>									
						<tr>
							<td>&nbsp;</td>
						</tr>
					</table>
					<span class="menuInside" title="Poppe nytt vindu"><a href="visNoHTML.asp?target=Browser" target="_blank" ><img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Popsie daisy</a></span><br>
					<br>
					<span class="menuInside" title="Poppe nytt vindu"><a href="#" onclick="javascript:document.all.frmGenererCV.submit()"><img src="/xtra/images/icon_new2.gif" width="14" height="14" alt="" border="0" align="absmiddle">&nbsp;Popsie daisy</a></span><br>
					&nbsp;
				</form>
			</div>
		</div>
	</body>
</html>
