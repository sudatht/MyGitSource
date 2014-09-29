<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset><legend><%= GetString("InsertForm") %></legend>
	<table class="normal">
		<tr>
			<td style="width:100"><%= GetString("Action") %>:</td>
			<td><input type="text" id="inp_action" style="width:200px" /></td>
		</tr>
		<tr>
			<td><%= GetString("Method") %>:</td>
			<td>
				<select id="sel_Method" style="width:200px">
					<option value="post">post</option>
					<option value="get">get</option>
				</select>
			</td>
		</tr>
		<tr>
			<td><%= GetString("Name") %>:</td>
			<td><input type="text" id="inp_name" style="width:200px" /></td>
		</tr>
		<tr>
			<td><%= GetString("ID") %>:</td>
			<td><input type="text" id="inp_id" style="width:200px" /></td>
		</tr>
		<tr>
			<td><%= GetString("EncodingType") %>:</td>
			<td><input type="text" id="inp_encode" style="width:200px" /></td>
		</tr>
		<tr>
			<td><%= GetString("Target") %>:</td>
			<td>				
				<select id="sel_target" name="sel_target">
					<option value=""></option>
					<option value="_blank">_blank</option>
					<option value="_self">_self</option>
					<option value="_top">_top</option>
					<option value="_parent">_parent</option>
				</select>
			</td>
		</tr>
	</table>
</fieldset>

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Form.js"></script>