<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("Find") %> </title>
		
		<meta name="content-type" content="text/html ;charset=Unicode" />
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
	</head>
	<body>
			<table cellspacing="2" cellpadding="5" width="96%">
				<tr>
					<td class="normal" valign="top">
						<table cellspacing="2" cellpadding="5" width="100%" class="normal">
							<tr>
								<td valign="top" ><%= GetString("Find what") %>:
								</td>
								<td valign="top" style="white-space:nowrap"><input type="text" size="30" name="stringSearch" id="stringSearch" style="width : 180px;" /></td>
							</tr>
							<tr>
								<td valign="top"><%= GetString("Replace with") %>:
								</td>
								<td valign="top"><input type="text" size="30" name="stringReplace" id="stringReplace" style="width : 180px;" /></td>
							</tr>
						</table>
						<input type="Checkbox" size="40" name="MatchWholeWord" id="MatchWholeWord" /><%= GetString("Match whole word") %><br />
						<input type="Checkbox" size="40" name="MatchCase" id="MatchCase" /><%= GetString("Match case") %>
					</td>
					<td valign="top" class="normal">
					    <input type="button" style="width:75px; height:22px; margin-top:3px" class="formbutton" onclick="FindTxt();" value="<%= GetString("Find Next") %>" />
					    <input type="button" style="width:75px; height:22px; margin-top:7px" class="formbutton" onclick="ReplaceTxt();" value="<%= GetString("Replace") %>" />
					    <input type="button" style="width:75px; height:22px; margin-top:7px" class="formbutton" onclick="ReplaceAllTxt();" value="<%= GetString("Replace All") %>"/>
					    <input type="button" style="width:75px; height:22px; margin-top:7px" class="formbutton" onclick="top.close();" value="<%= GetString("Cancel") %>"/>
					</td>
				</tr>
			</table>
	</body>
	<script>
	    var WordNotFound = "<%= GetString("WordNotFound")%>";
	    var WordReplaced = "<%= GetString("WordReplaced")%>";
	</script>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_Find.js"></script>
</html>
