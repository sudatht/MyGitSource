<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("PasteText") %></title>		
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
	</head>
	<body>
		<div id="container">
			<table border="0" cellpadding="0" cellspacing="2" width="100%" ID="Table1">
				<tr>
					<td><%= GetString("UseCtrl_VtoPaste") %></td>
					<td style="white-space:nowrap" >
						<input type="checkbox" name="linebreaks" id="linebreaks" checked="checked" /><%= GetString("KeepLinebreaks") %>
					</td> 
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="2" width="95%" ID="Table2">
				<tr>
					<td>
						<textarea name="idSource" id="idSource" rows="20" cols="58"></textarea>
          			</td>
				</tr>
			</table>
			<br>
			<table border="0" cellpadding="0" cellspacing="2" width="100%" ID="Table3">
				<tr>
					<td align="left"><input type="button" value="<%= GetString("CleanUpBox") %>" class="formbutton" onclick="document.getElementById('idSource').value='';" id="Button2" />
					</td>
					<td align="right" style="padding-right:100px">
						<input type="button" id="insert" name="insert" value="<%= GetString("Insert") %>" class="formbutton" onclick="insertContent();" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" value="<%= GetString("Cancel") %>" onclick="cancel();" class="formbutton" id="Button1" />
					</td>
				</tr>
			</table>	
    </div>
	</body>
	<script type="text/javascript">		
			var OxO7b9b=["width","style","idSource","100%","height"];if(!Browser_IsSafari()){document.getElementById(OxO7b9b[2])[OxO7b9b[1]][OxO7b9b[0]]=OxO7b9b[3];document.getElementById(OxO7b9b[2])[OxO7b9b[1]][OxO7b9b[4]]=OxO7b9b[3];} ;
	</script>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_gecko_pastetext.js"></script>
</html>
