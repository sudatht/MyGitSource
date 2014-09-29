<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("Hyperlink_Information") %> 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		</title>
		
		<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.1)" />
		<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.1)" />
		<link href="../Themes/<%=Theme%>/dialog.css" type="text/css" rel="stylesheet" />
		<!--[if IE]>
			<link href="../Style/IE.css" type="text/css" rel="stylesheet" />
		<![endif]-->
		<script type="text/javascript" src="../Scripts/Dialog/DialogHead.js"></script>
		<style type="text/css">
		    .btn {border: 1px solid buttonface;padding: 1px;width:14px;height: 12px;vertical-align: middle;}
		</style>
	</head>
	<body  style="margin:0px;border-width:0px;padding:4px;">
		<table border="0" cellspacing="2" cellpadding="5" width="100%" style="text-align:center">
			<tr>
				<td>
					<div>
					<fieldset>
						<table class="normal">
							<tr>
								<td style="width:60px"><%= GetString("Url") %>:</td>
								<td><input type="text" id="inp_src" style="width:200px" /></td>
								<td>
								    <input id="btnbrowse" class="formbutton" type="button" value="<%= GetString("Browse") %>" />
								</td>
							</tr>
							<tr>
								<td style="width:60px"><%= GetString("Title") %>:</td>
								<td colspan="2"><input type="text" id="inp_title" style="width:200px" /></td>
							</tr>
							<tr>
								<td style="width:60px"><%= GetString("Type") %>:</td>
								<td colspan="2">
									<select id="sel_protocol" onchange="sel_protocol_change()">
										<option value="http://">http://</option>
										<option value="https://">https://</option>
										<option value="ftp://">ftp://</option>
										<option value="news://">news://</option>
										<option value="mailto:">mailto:</option>
										<!-- last one : if move this to front , change the script too -->
										<option value="others"><%= GetString("Other") %></option>
									</select>
								</td>
							</tr>
							<tr>
								<td style="width:60px"><%= GetString("Target") %></td>
								<td colspan="2">
									<select id="inp_target" name="inp_target">
										<option value=""><%= GetString("NotSet") %></option>
										<option value="_blank"><%= GetString("Newwindow") %></option>
										<option value="_self"><%= GetString("Samewindow") %></option>
										<option value="_top"><%= GetString("Topmostwindow") %></option>
										<option value="_parent"><%= GetString("Parentwindow") %></option>
									</select>
								</td>
							</tr>		
						</table>
					</fieldset>
					</div>
					<div style="margin-top:8px;width:60%; text-align:center">
<input class="formbutton" type="button" value="<%= GetString("Insert") %>" style="width:80px" onclick="insert_link()" />&nbsp;&nbsp;
<input class="formbutton" type="button" value="<%= GetString("Cancel") %>" style="width:80px" onclick="do_Close()" />
					</div>
				</td>
			</tr>
		</table>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/Dialog/Dialog_MapLink.js"></script>
</html>
