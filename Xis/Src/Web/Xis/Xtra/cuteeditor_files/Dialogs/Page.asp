<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("DocumentPropertyPage") %></title>
		
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
	</head>
	<body>
		<div id="container">
		<table border="0" cellspacing="0" cellpadding="2" style="width:100%;" id="Table1">
			<tr>
				<td valign="top">
						<table border="0" cellpadding="3" cellspacing="0" style="border-collapse:collapse;" class="normal"
							id="Table2">
							<tr>
								<td><%= GetString("Title") %>:</td>
								<td colspan="4">
									<input type="text" id="inp_title" style="width:320px" name="inp_title" />
								</td>
							</tr>
							<tr>
								<td>DOCTYPE:</td>
								<td colspan="4">
									<input type="text" id="inp_doctype" style="width:320px" name="inp_doctype" />
								</td>
							</tr>
							<tr>
								<td><%= GetString("Description") %>:</td>
								<td colspan="4">
									<textarea id="inp_description" rows="3" cols="20" style="width:320px" name="inp_description"></textarea>
								</td>
							</tr>
							<tr>
								<td><%= GetString("Keywords") %>:</td>
								<td colspan="4">
									<textarea id="inp_keywords" rows="3" cols="20" style="width:320px" name="inp_keywords"></textarea>
								</td>
							</tr>
							<tr>
								<td><%= GetString("PageLanguage") %>:</td>
								<td colspan="4">
									<input type="text" id="PageLanguage" name="PageLanguage" size="15" style="WIDTH:320px" />
								</td>
							</tr>
							<tr>
								<td><%= GetString("HTMLEncoding") %>:</td>
								<td colspan="4">
									<input type="text" id="HTMLEncoding" name="HTMLEncoding" size="15" style="WIDTH:320px" />
								</td>
							</tr>
							<tr>
								<td><%= GetString("Backgroundcolor") %>:</td>
								<td>
									<input type="text" id="bgcolor" name="bgcolor" size="7" style="WIDTH:57px" /> 
									<img alt="" id="bgcolor_Preview" src="../Images/colorpicker.gif" style="vertical-align:top;" />
								</td>
								<td style="width:5"></td>
								<td><%= GetString("ForeColor") %>:</td>
								<td>
									<input autocomplete="off" type="text" id="fontcolor" name="fontcolor" size="7" />
								    <img alt="" src="../Images/colorpicker.gif" id="fontcolor_Preview" style="vertical-align:top;" />
								</td>
							</tr>
							<tr>
								<td><%= GetString("Backgroundimage") %>:</td>
								<td colspan="4">
									<input type="text" id="Backgroundimage" style="width:250px" name="Backgroundimage" />
									<input type="button"  class="formbutton" id="btnbrowse" value="<%= GetString("Browse") %>"/>
								</td>
							</tr>
							<tr>
								<td><%= GetString("TopMargin") %>:</td>
								<td>
									<input type="text" id="TopMargin" name="TopMargin" size="7" style="WIDTH:57px" /> 
									Pixels
								</td>
								<td style="width:5"></td>
								<td><%= GetString("BottomMargin") %>:</td>
								<td>
									<input type="text" id="BottomMargin" name="BottomMargin" size="7" style="WIDTH:57px" />
									Pixels
								</td>
							</tr>
							<tr>
								<td><%= GetString("LeftMargin") %>:</td>
								<td>
									<input type="text" id="LeftMargin" name="LeftMargin" size="7" style="WIDTH:57px" />
									Pixels
								</td>
								<td style="width:5"></td>
								<td><%= GetString("RightMargin") %>:</td>
								<td>
									<input type="text" id="RightMargin" name="RightMargin" size="7" style="WIDTH:57px" />
									Pixels
								</td>
							</tr>
							<tr>
								<td><%= GetString("MarginWidth") %>:</td>
								<td>
									<input type="text" id="MarginWidth" name="RightMargin" size="7" style="WIDTH:57px" />
									Pixels
								</td>
								<td style="width:5"></td>
								<td><%= GetString("MarginHeight") %>:</td>
								<td>
									<input type="text" id="MarginHeight" name="MarginHeight" size="7" style="WIDTH:57px" />
									Pixels
								</td>
							</tr>
						</table>
				</td>
			</tr>
		</table>
			<br>
			<div id="container-bottom">
<input type="button" id="btnok" style='width:80px' class="formbutton" value=" <%= GetString("OK") %> " />
					&nbsp; &nbsp;&nbsp;&nbsp;
<input type="button" id="btncc" style='width:80px' class="formbutton" value=" <%= GetString("Cancel") %> " />
			</div>					
		</div>
	</body>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_Page.js"></script>
</html>
