<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
	<head>
		<title>Gamle timelister</title>
		<SCRIPT language="javaScript" src="../js/javaScript.js"></SCRIPT>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	</head>
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1>Søk i gamle timelister</h1>
			</div>
			<div class="content">
				<p>Utelat ansattnummer og du får en liste over vikarer.</p>
				<FORM name="dato" ACTION="Vikar_timeliste_vis_gml.asp?frakode=0" METHOD=POST>
					<table>
						<tr>
						<td>Fra dato: <input type="text" name=FRADATO2 SIZE=6 ONBLUR="dateCheck(this.form, this.name)"></td>
						<td>Til dato: <input type="text" name=TILDATO2 SIZE=6 ONBLUR="dateCheck(this.form, this.name)"></td>
						<td>Ansattnummer: <input type="text" name=ANSATTNR SIZE=4 ></td>
						<td><INPUT TYPE=SUBMIT VALUE="Søk"></td>
						</tr>
					</table>
				</form>
			</div>
		</div>
	</body>
</html>
