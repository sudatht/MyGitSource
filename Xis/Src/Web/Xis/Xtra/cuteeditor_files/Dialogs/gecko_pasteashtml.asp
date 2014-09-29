<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("PasteAsHTML") %> </title>		
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
	</head>
	<body>
		<div id="container">
			<table border="0" cellpadding="0" cellspacing="2" width="100%" height="360" ID="Table1">
					<tr>
						<td><%= GetString("UseCtrl_VtoPaste") %></td>
					</tr>
					<tr>
						<td style="height:100%" >
							<iframe id="idSource" name="idSource" src="../Template.asp" scrolling="auto" style="border:1px solid #999999; WIDTH: 440; HEIGHT: 100%;background-color:#ffffff;"></iframe>
						</td>
					</tr>
				</table>
			<div id="container-bottom">
				<input type="button" id="Button2" name="insert" class="formbutton" value="<%= GetString("Insert") %>" onclick="insertContent();" />
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" value="<%= GetString("Cancel") %>" class="formbutton" onclick="cancel();" id="Button1" name="Button1" />
			</div>
    </div>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_gecko_pasteword.js"></script>
</html>