<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" --><style type="text/css">
#div_selector_event
{
	width:45;height:45;padding:5px;border:1px solid white;
}
#div_selector
{
	width:40;height:40;border:4px solid white;
}
</style>
<fieldset>
	<legend>
		<%= GetString("Borders") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:48">
				<div id="div_selector_event">
					<div id="div_selector">
					</div>
				</div>
			</td>
			<td>
				<select id="sel_part">
					<option value=""><%= GetString("All") %></option>
					<option value="Top"><%= GetString("Top") %></option>
					<option value="Left"><%= GetString("Left") %></option>
					<option value="Right"><%= GetString("Right") %></option>
					<option value="Bottom"><%= GetString("Bottom") %></option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>
<fieldset>
	<legend>
		<%= GetString("Border") %></legend>
	<table border="0" cellspacing="0" cellpadding="2" class="normal">
		<tr>
			<td style="width:48"><%= GetString("Margin") %></td>
			<td>
				<input type="text" id="tb_margin" style="width:80" />
				<select id="sel_margin_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
		<tr>
			<td><%= GetString("Padding") %></td>
			<td><input type="text" id="tb_padding" style="width:80" />
				<select id="sel_padding_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select></td>
		</tr>
		<tr>
			<td><%= GetString("Border") %></td>
			<td><input type="text" id="tb_border" style="width:80" />
				<select id="sel_border_unit">
					<option value="px">px</option>
					<option value="pt">pt</option>
					<option value="pc">pc</option>
					<option value="em">em</option>
					<option value="cm">cm</option>
					<option value="mm">mm</option>
					<option value="in">in</option>
				</select>
				<%= GetString("OR") %>
				<select id="sel_border">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="medium"><%= GetString("Medium") %></option>
					<option value="thin"><%= GetString("Thin") %></option>
					<option value="thick"><%= GetString("Thick") %></option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("Style") %></td>
			<td><select id="sel_style">
					<option value=""><%= GetString("NotSet") %></option>
					<option value="none"><%= GetString("None") %></option>
					<option value="solid">solid</option>
					<option value="inset">inset</option>
					<option value="outset">outset</option>
					<option value="ridge">ridge</option>
					<option value="dotted">dotted</option>
					<option value="dashed">dashed</option>
					<option value="groove">groove</option>
					<option value="double">double</option>
				</select></td>
		</tr>
		<tr>
			<td><%= GetString("Color") %></td>
			<td>
				<input autocomplete="off" size="7" type="text" id="inp_color" style="width:57px"/>
				<img alt="" id="inp_color_Preview" src="../Images/colorpicker.gif" style="vertical-align:top" />
			</td>
		</tr>
	</table>
</fieldset>

<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />
<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Style_Border.js"></script>