<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" -->
<fieldset>
	<legend>
	<%= GetString("Cursor") %>
	</legend>
	<select id="sel_cursor">
		<option value=""><%= GetString("NotSet") %></option>
		<option value="Default"><%= GetString("Default") %></option>
		<option value="Move"><%= GetString("Move") %></option>
		<option value="Text">Text</option>
		<option value="Wait">Wait</option>
		<option value="Help">Help</option>
		<!-- x-resize -->
	</select>
</fieldset>

<fieldset>
	<legend>
	<%= GetString("Filter") %>
	</legend>
	<input type="text" id="inp_filter" style="width:240px" /> <!--button filter builder-->
</fieldset>

<div id="outer" style="height:100px; margin-bottom:10px; padding:5px;"><div id="div_demo"><%= GetString("DemoText") %></div></div>
<br />

<script type="text/javascript">

var OxO92e0=["sel_cursor","inp_filter","outer","div_demo","cssText","style","value","cursor","filter"];var sel_cursor=Window_GetElement(window,OxO92e0[0],true);var inp_filter=Window_GetElement(window,OxO92e0[1],true);var outer=Window_GetElement(window,OxO92e0[2],true);var div_demo=Window_GetElement(window,OxO92e0[3],true);function UpdateState(){div_demo[OxO92e0[5]][OxO92e0[4]]=element[OxO92e0[5]][OxO92e0[4]];} ;function SyncToView(){sel_cursor[OxO92e0[6]]=element[OxO92e0[5]][OxO92e0[7]];if(Browser_IsWinIE()){inp_filter[OxO92e0[6]]=element[OxO92e0[5]][OxO92e0[8]];} ;} ;function SyncTo(element){element[OxO92e0[5]][OxO92e0[7]]=sel_cursor[OxO92e0[6]];if(Browser_IsWinIE()){element[OxO92e0[5]][OxO92e0[8]]=inp_filter[OxO92e0[6]];} ;} ;

</script>