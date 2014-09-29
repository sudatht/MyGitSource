<%
	dim XisParameterError : XisParameterError = ""
	
	sub AddErrorMessage(message)
		XisParameterError = XisParameterError & message & "<br>"
	end sub

	function HasError()
		HasError = false
		if (len(XisParameterError) > 0) then
			HasError = true
		end if
	end function

	sub RenderErrorMessage()
		%>
		<html>
			<head>
				<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
				<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
				<title>Manglende/feil data</title>
			</head>
			<body>
				<div class="pageContainer" id="pageContainer">
					<div class="contentHead1">
						<H1>Feil eller manglende data</H1>
					</div>
					<div class="content">
						<% = XisParameterError%>
						<a href="javascript:history.go(-1)">[Tilbake]</a>
					</div>
				</div>
			</body>
		</html>
		<%
		Response.End
	end sub
%>