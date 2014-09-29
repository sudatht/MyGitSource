<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset style="padding: 3px;"><legend><%= GetString("Backgroundcolor") %></legend>
	<input type="text" id="inp_color" name="inp_color" size="7" style="WIDTH:57px" />
	<img alt="" src="../images/colorpicker.gif" id="inp_color_Preview" style='vertical-align:top;' />
	
</fieldset>
<fieldset style="padding: 3px;"><legend><%= GetString("Backgroundimage") %></legend>
	<div>
		<%= GetString("Url") %>: <input id="tb_image" type="text" style="width:220px" />
		<input type="button" id="btnbrowse" value=" ... "/>
	</div>
	<div style="padding-left: 32px;">
		<table border="0" cellpadding="2" cellspacing="0" class="normal">
			<tr>
				<td style="width:80"><%= GetString("Tiling") %>: </td>
				<td><select id="sel_bgrepeat" style="width:140px">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="repeat"><%= GetString("Tilingboth") %></option>
						<option value="repeat-x"><%= GetString("Tilingorizontal") %></option>
						<option value="repeat-y"><%= GetString("Tilingvertical") %></option>
						<option value="no-repeat"><%= GetString("NoTiling") %></option>
					</select>
				</td>
			</tr>
			<tr>
				<td><%= GetString("Scrolling") %>: </td>
				<td><select id="sel_bgattach" style="width:140px">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="scroll"><%= GetString("Scrollingbackground") %></option>
						<option value="fixed"><%= GetString("ScrollingFixed") %></option>
					</select>
				</td>
			</tr>
		</table>
	</div>
	<fieldset><legend><%= GetString("Position") %></legend>
		<table border="0" cellpadding="2" cellspacing="0" class="normal">
			<tr>
				<td><%= GetString("Horizontal") %></td>
				<td><select style="width:64" id="sel_hor">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="left"><%= GetString("Left") %></option>
						<option value="center"><%= GetString("Center") %></option>
						<option value="right"><%= GetString("Right") %></option>
					</select>
					<%= GetString("OR") %> <input type="text" id="tb_hor" style="width:42" />
					<select id="sel_hor_unit">
						<option value="px">px</option>
						<option value="pt">pt</option>
						<option value="pc">pc</option>
						<option value="em">em</option>
						<option value="cm">cm</option>
						<option value="mm">mm</option>
						<option value="in">in</option>
					</select>
				</td>
			</tr>
			<tr>
				<td><%= GetString("Vertical") %></td>
				<td><select style="width:64" id="sel_ver">
						<option value=""><%= GetString("NotSet") %></option>
						<option value="top"><%= GetString("top") %></option>
						<option value="center"><%= GetString("Center") %></option>
						<option value="bottom"><%= GetString("Bottom") %></option>
					</select>
					<%= GetString("OR") %> <input type="text" id="tb_ver" style="width:42" />
					<select id="sel_ver_unit">
						<option value="px">px</option>
						<option value="pt">pt</option>
						<option value="pc">pc</option>
						<option value="em">em</option>
						<option value="cm">cm</option>
						<option value="mm">mm</option>
						<option value="in">in</option>
					</select>
				</td>
			</tr>
		</table>
	</fieldset>
</fieldset>

<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style_Background.js"></script>