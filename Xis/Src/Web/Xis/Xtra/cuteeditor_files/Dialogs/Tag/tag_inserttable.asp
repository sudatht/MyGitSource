<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<br />
<table width="400" cellpadding="1" cellspacing="0" onmouseover="CuteEditor_ColorPicker_ButtonOver(this);"
	onclick="selectTemplates()" id="richDropDown" style="border:1px solid #cccccc;height:30;">
	<tr>
		<td style="padding-left:20px; background-color:white">
			<img src="../Images/h-f-3Columns-Body.gif" alt="<%= GetString("Table layout") %>" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			<%= GetString("Table layout") %>
		</td>
		<td style='width:8px;padding:0px 1px 1px 1px;color:black;text-align:center;font-family:Webdings;font-size:8px;'>6</td>
	</tr>
</table>
<div id="list_Templates" style="display:none;">
	<div style="position:absolute; top:0; left:0; overflow:scroll; overflow-x:hidden;width:400; height:220; border-bottom:0px solid black;">
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#99ccff, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px;cursor:hand;">
			<span onclick="parent.selectTemplate(1)"><img src="../Images/One-Column-Table.gif" alt="<%= GetString("1ColumnTable") %>" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
				<%= GetString("1ColumnTable") %></span>
		</div>
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=yellowgreen, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px; cursor:hand; border-top:0px solid black">
			<span onclick="parent.selectTemplate(2)"><img src="../Images/Two-Column-Table.gif" alt="<%= GetString("2ColumnTable") %>" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
				<%= GetString("2ColumnTable") %></span>
		</div>
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=gold, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px; cursor:hand; border-top:0px solid black">
			<span onclick="parent.selectTemplate(3)"><img src="../Images/Three-Column-Table.gif" alt="<%= GetString("3ColumnTable") %>"
					/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= GetString("3ColumnTable") %></span>
		</div>
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=violet, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px; cursor:hand; border-top:0px solid black">
			<span onclick="parent.selectTemplate(4)"><img src="../Images/h-R-t-Body.gif" alt="<%= GetString("Header-Right-TopLeft-Body") %>"
					/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= GetString("Header-Right-TopLeft-Body") %></span>
		</div>
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#99ccff, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px; cursor:hand; border-top:0px solid black">
			<span onclick="parent.selectTemplate(5)"><img src="../Images/h-l-tr-Body.gif" alt="<%= GetString("Header-Left-TopRight-Body") %>"
					/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= GetString("Header-Left-TopRight-Body") %></span>
		</div>
		<div onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#99ccff, EndColorStr=#FFFFFF)';"
			onmouseout="this.style.filter='';" style="font:normal 11px Tahoma; height:25px; background:#ffffff; border:1px solid black; padding:3px; padding-left:20px;  cursor:hand; border-top:0px solid black;">
			<span onclick="parent.selectTemplate(6)"><img src="../Images/h-f-3Columns-Body.gif" alt="<%= GetString("Header-Footer-3-Columns") %>"
					/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= GetString("Header-Footer-3-Columns") %></span>
		</div>
	</div>
</div>
<style type="text/css">
.addsub
{
	width:21px;
	font-family:fixedsys;
}
</style>
<br >
<div style="padding: 4px 5px 0 4px; width:99%">
	<table border="0" cellspacing="0" cellpadding="2"  class="normal" width="99%">
		<tr>
			<td align="center">
				<%= GetString("EditCell") %>
			</td>
			<td align="center">
				<%= GetString("Columns") %>: <button id="subcolumns" class="addsub">-</button> <button id="addcolumns" class="addsub">
					+</button> ColSpan : <button id="subcolspan" class="addsub">-</button> <button id="addcolspan" class="addsub">
					+</button>
			</td>
			<td></td>
		</tr>
		<tr>
			<td valign="top">
				<table class="normal">
					<tr>
						<td colspan="2" align="center">
							<input type="button" id="btn_row_dialog" value="<%= GetString("EditRow") %>" /></td>
					</tr>
					<tr>
						<td colspan="2" align="center">
							<input type="button" id="btn_cell_dialog" value="<%= GetString("EditCell") %>" /></td>
					</tr>
					<tr>
						<td><%= GetString("Width") %></td>
						<td><input type="text" id="inp_cell_width" style="width:60px" /></td>
					</tr>
					<tr>
						<td><%= GetString("Height") %>:</td>
						<td><input type="text" id="inp_cell_height" style="width:60px" /></td>
					</tr>
					<!-- //TODO: add more cell useful properties here -->
					<tr>
						<td colspan="2" align="center">
							<input type="button" id="btn_cell_editcell" value="<%= GetString("EditHtml") %>" /></td>
					</tr>
				</table>
				<br />
			</td>
			<td>
				<div style="border:1px solid gray;padding:1px;OVERFLOW: auto; HEIGHT: 215px; HEIGHT: 215px; ">
					<table id="tabledesign" border="1" cellspacing="" style='border-color:#FFA500;background-color:white;width:100%;height:210px;border-collapse:collapse' class="normal">
					</table>
				</div>
			</td>
			<td align="center">
				R<br />
				o<br />
				w<br />
				s<br />
				<button id="subrows" class="addsub">-</button><br />
				<button id="addrows" class="addsub">+</button>
				<br />
				S<br />
				p<br />
				a<br />
				n<br />
				<button id="subrowspan" class="addsub">-</button><br />
				<button id="addrowspan" class="addsub">+</button>
			</td>
		</tr>
	</table>
</div>

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_InsertTable.js"></script>