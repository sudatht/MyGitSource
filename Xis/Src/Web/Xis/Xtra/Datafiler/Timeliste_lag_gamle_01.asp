<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<% 
If (HasUserRight(ACCESS_ADMIN, RIGHT_ADMIN) = false) Then
	call Response.Redirect("/xtra/IngenTilgang.asp")	
End If 
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" "http://www.w3.org/tr/rec-html40/loose.dtd">
<html>
<head>
	<title></title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script type="text/javascript" language="javascript" src="../Js/javaScript.js"></script>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<div class="contentHead1">
			<h1>Rutine for å fjerne gamle timelister fra vanlig visning</h1>
		</div>
			<div class="content">
				<p>
					Setter timelistestatus til 6 på timelister der hvor både lønnsstatus og fakturastatus = 3<br>
					(Obs! Dato avrundes til nærmeste søndag.)
					<form name="FormEn" action="Timeliste_lag_gamle_db.asp" method=POST id="Form1">
						<input type=text size=8 maxlength="8" name=limitdato onblur="dateCheck(this.form, this.name)" id="limitdato">
						<input type=SUBMIT value="Overfør til gamle" id="Submit1" name="Submit1">
					</form>
				</p>
			</div>
		</div>
	</body>
</html>