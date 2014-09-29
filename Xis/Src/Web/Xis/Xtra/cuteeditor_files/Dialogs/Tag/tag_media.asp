<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<!-- #include file = "include_Security.asp" --><fieldset><legend><%= GetString("Src") %></legend>
	<input type="text" id="inp_src" style="width:320px" /><button id="btnbrowse"><%= GetString("Browse") %></button>
</fieldset>
<fieldset style="height:180px;width:270px;overflow:hidden;"><legend><%= GetString("Demo") %></legend>
	<img id="img_demo" src="" alt="" />
</fieldset>

<script type="text/javascript" src="../Scripts/Dialog/Dialog_Tag_Media.js"></script>
