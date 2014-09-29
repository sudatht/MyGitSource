<!-- #include file = "Include_GetString.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title><%= GetString("SyntaxHighlighter") %></title>
		
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
			<table>
				<tr>
					<td width="80"><%= GetString("CodeLanguage") %>:</td>
					<td><select id="sel_lang"></select></td>
				</tr>
				<tr>
					<td colspan="2"><textarea id="ta_code" name="ta_code_name" style="width:400px;height:300px"></textarea></td>
				</tr>
			</table>
			
			<div id="container-bottom">
				<input type="button" value="<%= GetString("OK") %>" class="formbutton" onclick="DoHighlight()">
				<input type="button" value="<%= GetString("Cancel") %>" class="formbutton" onclick="Close()">
			</div>					
		</div>
	</body>
	<script type="text/javascript" src="../Scripts/Dialog/DialogFoot.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/_shCore.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushCpp.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushCSharp.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushCss.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushDelphi.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushJava.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushJScript.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushPhp.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushPython.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushRuby.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushSql.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushVb.js"></script>
	<script type="text/javascript" src="../Scripts/SyntaxHighlighter/shBrushXml.js"></script>
<script>
	

var OxOa809=["=","; path=/;"," expires=",";","cookie","length","","sel_lang","ta_code","Brushes","sh","Aliases","options","CESHBRUSH","value","language",":nocontrols","opera","all","innerHTML","display","style","previousSibling","\x3Cdiv class=\x22dp-highlighter\x22\x3E","\x3C/div\x3E","parentNode"];function SetCookie(name,Oxb9,Oxba){var Oxbb=name+OxOa809[0]+escape(Oxb9)+OxOa809[1];if(Oxba){var Oxbc= new Date();Oxbc.setSeconds(Oxbc.getSeconds()+Oxba);Oxbb+=OxOa809[2]+Oxbc.toUTCString()+OxOa809[3];} ;document[OxOa809[4]]=Oxbb;} ;function GetCookie(name){var Oxbe=document[OxOa809[4]].split(OxOa809[3]);for(var i=0;i<Oxbe[OxOa809[5]];i++){var Oxbf=Oxbe[i].split(OxOa809[0]);if(name==Oxbf[0].replace(/\s/g,OxOa809[6])){return unescape(Oxbf[1]);} ;} ;} ;var editor=Window_GetDialogArguments(window);var sel_lang=document.getElementById(OxOa809[7]);var ta_code=document.getElementById(OxOa809[8]);for(var brush in dp[OxOa809[10]][OxOa809[9]]){var aliases=dp[OxOa809[10]][OxOa809[9]][brush][OxOa809[11]];if(aliases==null){continue ;} ;sel_lang[OxOa809[12]].add( new Option(aliases,brush));var b=GetCookie(OxOa809[13]);if(b){sel_lang[OxOa809[14]]=b;} ;} ;function DoHighlight(){SetCookie(OxOa809[13],sel_lang.value,3600*24*30);var b=dp[OxOa809[10]][OxOa809[9]][sel_lang[OxOa809[14]]];ta_code[OxOa809[15]]=b[OxOa809[11]][0]+OxOa809[16];if(window[OxOa809[17]]||!document[OxOa809[18]]){ta_code[OxOa809[19]]=ta_code[OxOa809[14]];} ;dp[OxOa809[10]].HighlightAll(ta_code.name);ta_code[OxOa809[21]][OxOa809[20]]=OxOa809[6];var Oxc5=ta_code[OxOa809[22]];editor.PasteHTML(OxOa809[23]+Oxc5[OxOa809[19]]+OxOa809[24]);Oxc5[OxOa809[25]].removeChild(Oxc5);Close();} ;



	</script>
</html>