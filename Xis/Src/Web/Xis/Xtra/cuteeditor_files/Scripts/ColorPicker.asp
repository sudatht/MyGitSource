<%  
   Response.ContentType = "text/x-component"
   dim Culture
   Culture = Trim(Request.QueryString("UC"))  
   dim FilePath
   FilePath = Trim(Request.QueryString("F"))  
   Public Function GetString(instring)
	    dim t
    	
	    t = GetStringByCulture(instring,Culture)
    	
	    If t = ""  then
		    t= GetStringByCulture(instring,"_default")
	    End If
    	
	    If t = ""  then
		    t= "{"&instring&"}"	
	    End If
	    GetString= t	
    End Function
    
    
   FilePath = Left(request.ServerVariables("PATH_INFO"),InStr(request.ServerVariables("PATH_INFO"),FilePath)+len(FilePath)-1)

    Public Function GetStringByCulture(instring,input_culture)
	    dim scriptname,xmlfilename,doc,temp
	    dim node,selectednode,optionnodelist,errobj
	    dim selectednodes

	    xmlfilename= Server.MapPath(FilePath&"/languages/"&input_culture&".xml")

	    ' Create an object to hold the XML
	    set doc = server.CreateObject("Microsoft.XMLDOM")

	    ' For ASP, wait until the XML is all ready before continuing
	    doc.async = False

	    ' Load the XML file or return an error message and stop the script
	    if not Doc.Load(xmlfilename) then
		    Response.Write "Failed to load the language text from the XML file.<br>"
		    Response.End
	    end if

	    ' Make sure that the interpreter knows that we are using XPath as our selection language
	    doc.setProperty "SelectionLanguage", "XPath"

	    set selectednode= doc.selectSingleNode("/resources/resource[@name='"&instring&"']")
	    if IsObject(selectednode) and not selectednode is nothing  then
		    GetStringByCulture=CuteEditor_Decode(selectednode.text)
	    else
		    GetStringByCulture=""		
	    end if
    End Function    
    
     PUBLIC FUNCTION CuteEditor_Decode(input_str)        
	    input_str=Replace(input_str,"#1","<")
	    input_str=Replace(input_str,"#2",">")
	    input_str=Replace(input_str,"#3","&")
	    input_str=Replace(input_str,"#4","*")
	    input_str=Replace(input_str,"#5","o")
	    input_str=Replace(input_str,"#6","O")
	    input_str=Replace(input_str,"#7","s")
	    input_str=Replace(input_str,"#8","S")
	    input_str=Replace(input_str,"#9","e")
	    input_str=Replace(input_str,"#a","E")
	    input_str=Replace(input_str,"#0","#")
	    CuteEditor_Decode = input_str
     END FUNCTION
%>
<PUBLIC:COMPONENT>
	<PUBLIC:EVENT ID="event_oncolorchange" name="oncolorchange" />
	<PUBLIC:EVENT ID="event_oncolorpopup" name="oncolorpopup" />
	<PUBLIC:PROPERTY name="selectedColor" GET="_get_selectedColor" PUT="_set_selectedColor"/>
	<PUBLIC:METHOD name="popupColor" INTERNALNAME="_mtd_popupColor" />
	<PUBLIC:ATTACH EVENT="onclick" ONEVENT="_mtd_onclick()" />
</PUBLIC:COMPONENT>

<script type="text/javascript">
var OxOd0b0=["#000000","#993300","#333300","#003300","#003366","#000080","#333399","#333333","#800000","#FF6600","#808000","#008000","#008080","#0000FF","#666699","#808080","#FF0000","#FF9900","#99CC00","#339966","#33CCCC","#3366FF","#800080","#999999","#FF00FF","#FFCC00","#FFFF00","#00FF00","#00FFFF","#00CCFF","#993366","#C0C0C0","#FF99CC","#FFCC99","#FFFF99","#CCFFCC","#CCFFFF","#99CCFF","#CC99FF","#FFFFFF","dialogWidth:500px;dialogHeight:330px;help:0;status:0;resizable:1","disableVisual","","\x3CDIV style=\x27width=140;cursor:default;position:absolute;z-index:88888888;background-color:white;border:0px;overflow:visible;\x27\x3E","length","\x3Ctable cellpadding=0 cellspacing=5 style=\x27width:100%;font-family: Verdana; font-size: 6px; BORDER: #666666 1px solid;\x27 bgcolor=#f9f8f7\x3E\x3Ctr\x3E\x3Ctd\x3E","\x3Ctable cellpadding=0 cellspacing=2 style=\x27font-size: 3px;\x27\x3E","\x3Ctr\x3E","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;margin:1px\x22 onmouseup=\x22document.all.","uniqueID","._cphtc_sel(this.ColorValue)\x22  ColorValue=\x22\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22\x3E","\x3Ctable cellspacing=0 cellpadding=0 border=0 width=90% style=\x22font-size:3px\x22\x3E","\x3C/table\x3E","\x3C/td\x3E","\x3C/tr\x3E","\x3Ctr\x3E\x3Ctd\x3E\x26nbsp;\x3C/td\x3E\x3C/tr\x3E","\x3Ctd title="," align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 ColorValue=\x22","\x22 onmouseup=\x22document.all.","._cphtc_sel(this.ColorValue);\x22\x3E","\x3Cdiv style=\x22background-color:","; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;\x22\x3E\x26nbsp;\x3C/div\x3E","\x3C/td\x3E\x3C/tr\x3E","\x3Ctd colspan=10 align=center style=\x22padding:1px;border:solid 1px #f9f8f7;\x22 onmouseover=\x22CuteEditor_ColorPicker_ButtonOver(this);\x22 onmouseup=\x22document.all.","._cphtc_dlg();\x22\x3E","innerHTML","body","document","onclick","SELECT","all","visibility","currentStyle","hidden","runtimeStyle","style","_visibility","top","left","display","block","offsetHeight","px","unselectable","on","none"];var colorsArray= new Array(OxOd0b0[0],OxOd0b0[1],OxOd0b0[2],OxOd0b0[3],OxOd0b0[4],OxOd0b0[5],OxOd0b0[6],OxOd0b0[7],OxOd0b0[8],OxOd0b0[9],OxOd0b0[10],OxOd0b0[11],OxOd0b0[12],OxOd0b0[13],OxOd0b0[14],OxOd0b0[15],OxOd0b0[16],OxOd0b0[17],OxOd0b0[18],OxOd0b0[19],OxOd0b0[20],OxOd0b0[21],OxOd0b0[22],OxOd0b0[23],OxOd0b0[24],OxOd0b0[25],OxOd0b0[26],OxOd0b0[27],OxOd0b0[28],OxOd0b0[29],OxOd0b0[30],OxOd0b0[31],OxOd0b0[32],OxOd0b0[33],OxOd0b0[34],OxOd0b0[35],OxOd0b0[36],OxOd0b0[37],OxOd0b0[38],OxOd0b0[39]);var ShowMoreColors=true;var dlgurl='<%=FilePath%>/Dialogs/ColorPicker.Asp?<%=Request.ServerVariables("QUERY_STRING") %>&setting=<%= Request.Cookies("CESecurity") %>';function element._cphtc_sel(Ox37a){_color=Ox37a;event_oncolorchange.fire();} ;function element._cphtc_dlg(){CloseDiv();event_oncolorpopup.fire();var Ox37b=OxOd0b0[40];if(element[OxOd0b0[41]]){var Ox37c=showModalDialog(dlgurl,{color:Ox37d},Ox37b);if(Ox37c!=null&&Ox37c!=false){_color=Ox37c;event_oncolorchange.fire();} ;} else {var Ox37d=_color;var Ox37c=showModalDialog(dlgurl,{color:Ox37d,onchange:Ox37e},Ox37b);if(Ox37c!=null&&Ox37c!=false){_color=Ox37c;} else {_color=Ox37d;} ;event_oncolorchange.fire();function Ox37e(Oxd9){_color=Oxd9;event_oncolorchange.fire();} ;} ;} ;var _color=OxOd0b0[42];function _get_selectedColor(){return _color;} ;function _set_selectedColor(Oxd9){_color=Oxd9;} ;var div;var selects;var isopen=false;function _mtd_onclick(){_mtd_popupColor();} ;function _mtd_popupColor(){if(div==null){div=document.createElement(OxOd0b0[43]);var Ox387=OxOd0b0[42];var Ox388=colorsArray[OxOd0b0[44]];var Ox389=8;Ox387+=OxOd0b0[45];Ox387+=OxOd0b0[46];Ox387+=OxOd0b0[47];Ox387+=OxOd0b0[48]+element[OxOd0b0[49]]+OxOd0b0[50];Ox387+=OxOd0b0[51];Ox387+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px; font-size: 3px;">&nbsp;</div></td><td align=center style="font:normal 11px verdana;">&nbsp;<%= GetString("Automatic") %></td></tr>';Ox387+=OxOd0b0[52];Ox387+=OxOd0b0[53];Ox387+=OxOd0b0[54];Ox387+=OxOd0b0[55];for(var i=0;i<Ox388;i++){if((i%Ox389)==0){Ox387+=OxOd0b0[47];} ;Ox387+=OxOd0b0[56]+colorsArray[i]+OxOd0b0[57]+colorsArray[i]+OxOd0b0[58]+element[OxOd0b0[49]]+OxOd0b0[59];Ox387+=OxOd0b0[60]+colorsArray[i]+OxOd0b0[61];Ox387+=OxOd0b0[53];if(((i+1)>=Ox388)||(((i+1)%Ox389)==0)){Ox387+=OxOd0b0[54];} ;} ;Ox387+=OxOd0b0[55];Ox387+=OxOd0b0[52];Ox387+=OxOd0b0[62];if(ShowMoreColors){Ox387+=OxOd0b0[47];Ox387+=OxOd0b0[63]+element[OxOd0b0[49]]+OxOd0b0[64];Ox387+=OxOd0b0[51];Ox387+='<tr><td width=18><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px;font-size: 3px;"></div></td><td align=center style="font-size:11px"><%= GetString("MoreColors") %></td></tr>';Ox387+=OxOd0b0[52];Ox387+=OxOd0b0[53];Ox387+=OxOd0b0[54];} ;Ox387+=OxOd0b0[52];div[OxOd0b0[65]]=Ox387;element[OxOd0b0[67]][OxOd0b0[66]].appendChild(div);div[OxOd0b0[68]]=CloseDiv;} ;if(isopen){CloseDiv();} ;isopen=true;selects=[];var Ox82=element[OxOd0b0[67]][OxOd0b0[70]].tags(OxOd0b0[69]);for(var i=0;i<Ox82[OxOd0b0[44]];i++){var Ox38a=Ox82[i];if(Ox38a[OxOd0b0[72]][OxOd0b0[71]]!=OxOd0b0[73]){selects[selects[OxOd0b0[44]]]=Ox38a;var Ox38b=Ox38a[OxOd0b0[74]]||Ox38a[OxOd0b0[75]];Ox38b[OxOd0b0[76]]=Ox38b[OxOd0b0[71]];Ox38b[OxOd0b0[71]]=OxOd0b0[73];} ;} ;div[OxOd0b0[75]][OxOd0b0[77]]=0;div[OxOd0b0[75]][OxOd0b0[78]]=0;div[OxOd0b0[75]][OxOd0b0[79]]=OxOd0b0[80];var Ox11e=CalcPosition(div,element);Ox11e[OxOd0b0[77]]+=element[OxOd0b0[81]];AdjustMirror(div,element,Ox11e);div[OxOd0b0[75]][OxOd0b0[78]]=Ox11e[OxOd0b0[78]]+OxOd0b0[82];div[OxOd0b0[75]][OxOd0b0[77]]=Ox11e[OxOd0b0[77]]+OxOd0b0[82];var Ox82=div[OxOd0b0[70]];for(var i=0;i<Ox82[OxOd0b0[44]];i++){Ox82[i][OxOd0b0[83]]=OxOd0b0[84];} ;div.focus();var Ox38c= new CaptureManager(element,handlelosecapture);Ox38c.AddElement(div);} ;function handlelosecapture(){CloseDiv();} ;function CloseDiv(){CaptureManager.Unregister(element);isopen=false;div[OxOd0b0[75]][OxOd0b0[79]]=OxOd0b0[85];for(var i=0;i<selects[OxOd0b0[44]];i++){var Ox38a=selects[i];Ox38a[OxOd0b0[74]][OxOd0b0[71]]=Ox38a[OxOd0b0[74]][OxOd0b0[76]];} ;} ;

</script>

<script type="text/javascript">


var OxO74ea=["body","document","compatMode","CSS1Compat","documentElement","scrollLeft","scrollTop","clientLeft","clientTop","parentElement","position","currentStyle","absolute","relative","left","top","clientWidth","clientHeight","offsetWidth","offsetHeight","element","capturemanager","\x3CDIV style=\x27width:0px;height:0px;left:0px;top:0px;position:absolute\x27\x3E","afterBegin","onlosecapture","onmousedown","onmousemove","onmouseover","onmouseout","length"];function GetScrollPosition(Ox3e){var b=window[OxO74ea[1]][OxO74ea[0]];var p=b;if(window[OxO74ea[1]][OxO74ea[2]]==OxO74ea[3]){p=window[OxO74ea[1]][OxO74ea[4]];} ;if(Ox3e==b){return {left:0,top:0};} ;with(Ox3e.getBoundingClientRect()){return {left:p[OxO74ea[5]]+left,top:p[OxO74ea[6]]+top};} ;} ;function GetClientPosition(Ox3e){var b=window[OxO74ea[1]][OxO74ea[0]];var p=b;if(window[OxO74ea[1]][OxO74ea[2]]==OxO74ea[3]){p=window[OxO74ea[1]][OxO74ea[4]];} ;if(Ox3e==b){return {left:-p[OxO74ea[5]],top:-p[OxO74ea[6]]};} ;with(Ox3e.getBoundingClientRect()){return {left:left-p[OxO74ea[7]],top:top-p[OxO74ea[8]]};} ;} ;function GetStandParent(Ox3e){for(var Ox393=Ox3e[OxO74ea[9]];Ox393!=null;Ox393=Ox393[OxO74ea[9]]){var Ox1ad=Ox393[OxO74ea[11]][OxO74ea[10]];if(Ox1ad==OxO74ea[12]||Ox1ad==OxO74ea[13]){return Ox393;} ;} ;return window[OxO74ea[1]][OxO74ea[0]];} ;function CalcPosition(Ox395,Ox3e){var Ox396=GetScrollPosition(Ox3e);var Ox397=GetScrollPosition(GetStandParent(Ox395));var Ox398=GetStandParent(Ox395);return {left:Ox396[OxO74ea[14]]-Ox397[OxO74ea[14]]-Ox398[OxO74ea[7]],top:Ox396[OxO74ea[15]]-Ox397[OxO74ea[15]]-Ox398[OxO74ea[8]]};} ;function AdjustMirror(Ox395,Ox3e,Ox11e){var Ox39a=window[OxO74ea[1]][OxO74ea[0]][OxO74ea[16]];var Ox39b=window[OxO74ea[1]][OxO74ea[0]][OxO74ea[17]];if(window[OxO74ea[1]][OxO74ea[2]]==OxO74ea[3]){Ox39a=window[OxO74ea[1]][OxO74ea[4]][OxO74ea[16]];Ox39b=window[OxO74ea[1]][OxO74ea[4]][OxO74ea[17]];} ;var Ox39c=Ox395[OxO74ea[18]];var Ox39d=Ox395[OxO74ea[19]];var Ox39e=GetClientPosition(GetStandParent(Ox395));var Ox39f={left:Ox39e[OxO74ea[14]]+Ox11e[OxO74ea[14]]+Ox39c/2,top:Ox39e[OxO74ea[15]]+Ox11e[OxO74ea[15]]+Ox39d/2};var Ox3a0={left:Ox39e[OxO74ea[14]]+Ox11e[OxO74ea[14]],top:Ox39e[OxO74ea[15]]+Ox11e[OxO74ea[15]]};if(Ox3e!=null){var Ox3a1=GetClientPosition(Ox3e);Ox3a0={left:Ox3a1[OxO74ea[14]]+Ox3e[OxO74ea[18]]/2,top:Ox3a1[OxO74ea[15]]+Ox3e[OxO74ea[19]]/2};} ;var Ox3a2=true;if(Ox39f[OxO74ea[14]]-Ox39c/2<0){if((Ox3a0[OxO74ea[14]]*2-Ox39f[OxO74ea[14]])+Ox39c/2<=Ox39a){Ox39f[OxO74ea[14]]=Ox3a0[OxO74ea[14]]*2-Ox39f[OxO74ea[14]];} else {if(Ox3a2){Ox39f[OxO74ea[14]]=Ox39c/2+4;} ;} ;} else {if(Ox39f[OxO74ea[14]]+Ox39c/2>Ox39a){if((Ox3a0[OxO74ea[14]]*2-Ox39f[OxO74ea[14]])-Ox39c/2>=0){Ox39f[OxO74ea[14]]=Ox3a0[OxO74ea[14]]*2-Ox39f[OxO74ea[14]];} else {if(Ox3a2){Ox39f[OxO74ea[14]]=Ox39a-Ox39c/2-4;} ;} ;} ;} ;if(Ox39f[OxO74ea[15]]-Ox39d/2<0){if((Ox3a0[OxO74ea[15]]*2-Ox39f[OxO74ea[15]])+Ox39d/2<=Ox39b){Ox39f[OxO74ea[15]]=Ox3a0[OxO74ea[15]]*2-Ox39f[OxO74ea[15]];} else {if(Ox3a2){Ox39f[OxO74ea[15]]=Ox39d/2+4;} ;} ;} else {if(Ox39f[OxO74ea[15]]+Ox39d/2>Ox39b){if((Ox3a0[OxO74ea[15]]*2-Ox39f[OxO74ea[15]])-Ox39d/2>=0){Ox39f[OxO74ea[15]]=Ox3a0[OxO74ea[15]]*2-Ox39f[OxO74ea[15]];} else {if(Ox3a2){Ox39f[OxO74ea[15]]=Ox39b-Ox39d/2-4;} ;} ;} ;} ;Ox11e[OxO74ea[14]]=Ox39f[OxO74ea[14]]-Ox39e[OxO74ea[14]]-Ox39c/2;Ox11e[OxO74ea[15]]=Ox39f[OxO74ea[15]]-Ox39e[OxO74ea[15]]-Ox39d/2;} ;function CaptureManager(element,handlelosecapture){if(CaptureManager[OxO74ea[20]]&&CaptureManager[OxO74ea[20]][OxO74ea[21]]){CaptureManager[OxO74ea[20]][OxO74ea[21]].Unregister();} ;var Ox3a4=true;var Ox3a5=[];var Ox3a6=true;var Ox3a7=false;element[OxO74ea[21]]=Ox3ac;CaptureManager[OxO74ea[20]]=element;Ox3ac.AddElement(element);var Ox3a8=element[OxO74ea[1]].createElement(OxO74ea[22]);element[OxO74ea[1]][OxO74ea[0]].insertAdjacentElement(OxO74ea[23],Ox3a8);Ox3a8.attachEvent(OxO74ea[24],Ox3ad);Ox3a9(Ox3a8);Ox3a8.setCapture(true);Ox3a7=true;return Ox3ac;function Ox3a9(Ox3aa){Ox3aa.attachEvent(OxO74ea[25],Ox3ae);Ox3aa.attachEvent(OxO74ea[26],Ox3b0);Ox3aa.attachEvent(OxO74ea[27],Ox3b2);Ox3aa.attachEvent(OxO74ea[28],Ox3b3);} ;function Ox3ab(Ox3aa){Ox3aa.detachEvent(OxO74ea[25],Ox3ae);Ox3aa.detachEvent(OxO74ea[26],Ox3b0);Ox3aa.detachEvent(OxO74ea[27],Ox3b2);Ox3aa.detachEvent(OxO74ea[28],Ox3b3);} ;function Ox3ac(){} ;function Ox3ac.Unregister(){Ox3a8.detachEvent(OxO74ea[24],Ox3ad);Ox3ab(Ox3a8);Ox3a8.removeNode(true);for(var i=0;i<Ox3a5[OxO74ea[29]];i++){var Ox3aa=Ox3a5[i];Ox3ab(Ox3aa);} ;Ox3a4=false;element[OxO74ea[21]]=null;CaptureManager[OxO74ea[20]]=null;if(Ox3a7){Ox3a7=false;Ox3a8.releaseCapture();Ox3ac.FireLoseCapture();} ;} ;function Ox3ac.AddElement(Ox3aa){Ox3a9(Ox3aa);Ox3a5[Ox3a5[OxO74ea[29]]]=Ox3aa;} ;function Ox3ac.DelElement(Ox3aa){var len=Ox3a5[OxO74ea[29]];for(var i=0;i<len;i++){if(Ox3a5[i]==Ox3aa){Ox3ab(Ox3aa);for(var j=i;j<len-1;j++){Ox3a5[j]=Ox3a5[j+1];} ;Ox3a5[OxO74ea[29]]=Ox3a5[OxO74ea[29]]-1;return ;} ;} ;} ;function Ox3ac.FireLoseCapture(){handlelosecapture();} ;function Ox3ad(){if(Ox3a7){Ox3a7=false;try{Ox3ac.FireLoseCapture();} finally{Ox3ac.Unregister();} ;} ;} ;function Ox3ae(){var Ox3af=element[OxO74ea[1]].elementFromPoint(event.clientX,event.clientY);for(var i=0;i<Ox3a5[OxO74ea[29]];i++){var Ox3aa=Ox3a5[i];if(Ox3aa.contains(Ox3af)){return ;} ;} ;Ox3ac.Unregister();} ;function Ox3b0(){var Ox3af=element[OxO74ea[1]].elementFromPoint(event.clientX,event.clientY);Ox3b1(Ox3af);} ;function Ox3b1(Ox3af){for(var i=0;i<Ox3a5[OxO74ea[29]];i++){var Ox3aa=Ox3a5[i];if(Ox3aa.contains(Ox3af)){if(Ox3a7){Ox3a7=false;Ox3a8.releaseCapture();} ;return ;} ;} ;if(!Ox3a7){Ox3a7=true;Ox3a8.setCapture(true);} ;} ;function Ox3b2(){var Ox3af=element[OxO74ea[1]].elementFromPoint(event.clientX,event.clientY);Ox3a6=false;for(var i=0;i<Ox3a5[OxO74ea[29]];i++){var Ox3aa=Ox3a5[i];if(Ox3aa.contains(event.fromElement)){return ;} ;if(Ox3aa.contains(Ox3af)){if(Ox3a7){Ox3a7=false;Ox3a8.releaseCapture();} ;} ;} ;} ;function Ox3b3(){var Ox3af=element[OxO74ea[1]].elementFromPoint(event.clientX,event.clientY);Ox3a6=false;for(var i=0;i<Ox3a5[OxO74ea[29]];i++){var Ox3aa=Ox3a5[i];if(Ox3aa.contains(event.toElement)){return ;} ;} ;if(!Ox3a7){Ox3a7=true;Ox3a8.setCapture(true);} ;} ;} ;function CaptureManager.Register(element,handlelosecapture){return  new CaptureManager(element,handlelosecapture);} ;function CaptureManager.Unregister(element){if(element&&element[OxO74ea[21]]){element[OxO74ea[21]].Unregister();} ;} ;
</script>
