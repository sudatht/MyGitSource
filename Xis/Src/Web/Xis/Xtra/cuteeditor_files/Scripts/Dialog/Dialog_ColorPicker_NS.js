var OxO8818=["onload","onclick","btnCancel","btnOK","onkeyup","txtHSB_Hue","onkeypress","txtHSB_Saturation","txtHSB_Brightness","txtRGB_Red","txtRGB_Green","txtRGB_Blue","txtHex","btnWebSafeColor","rdoHSB_Hue","rdoHSB_Saturation","rdoHSB_Brightness","pnlGradient_Top","onmousemove","onmousedown","onmouseup","pnlVertical_Top","pnlWebSafeColor","pnlWebSafeColorBorder","pnlOldColor","lblHSB_Hue","lblHSB_Saturation","lblHSB_Brightness","length","\x5C{","\x5C}","BadNumber","A number between {0} and {1} is required. Closest value inserted.","Title","Color Picker","SelectAColor","Select a color:","OKButton","OK","CancelButton","Cancel","AboutButton","About","Recent","WebSafeWarning","Warning: not a web safe color","WebSafeClick","Click to select web safe color","HsbHue","H:","HsbHueTooltip","Hue","HsbHueUnit","%","HsbSaturation","S:","HsbSaturationTooltip","Saturation","HsbSaturationUnit","HsbBrightness","B:","HsbBrightnessTooltip","Brightness","HsbBrightnessUnit","RgbRed","R:","RgbRedTooltip","Red","RgbGreen","G:","RgbGreenTooltip","Green","RgbBlue","RgbBlueTooltip","Blue","Hex","#","RecentTooltip","Recent:","\x0D\x0ALewies Color Pickerversion 1.1\x0D\x0A\x0D\x0AThis form was created by Lewis Moten in May of 2004.\x0D\x0AIt simulates the color picker in a popular graphics application.\x0D\x0AIt gives users a visual way to choose colors from a large and dynamic palette.\x0D\x0A\x0D\x0AVisit the authors web page?\x0D\x0Awww.lewismoten.com\x0D\x0A","lblSelectColorMessage","lblRecent","lblRGB_Red","lblRGB_Green","lblRGB_Blue","lblHex","lblUnitHSB_Hue","lblUnitHSB_Saturation","lblUnitHSB_Brightness","pnlHSB_Hue","pnlHSB_Saturation","pnlHSB_Brightness","pnlRGB_Red","pnlRGB_Green","pnlRGB_Blue","frmColorPicker","Color","","FFFFFF","value","checked","ColorMode","ColorType","RecentColors","pnlRecent","border","style","0px","http://www.lewismoten.com","_blank","backgroundColor","target","rgb","(",")",",","display","none","title","innerHTML","backgroundPosition","px ","px","pnlGradientHsbHue_Hue","pnlGradientHsbHue_Black","pnlGradientHsbHue_White","pnlVerticalHsbHue_Background","pnlVerticalHsbSaturation_Hue","pnlVerticalHsbSaturation_White","pnlVerticalHsbBrightness_Hue","pnlVerticalHsbBrightness_Black","pnlVerticalRgb_Start","pnlVerticalRgb_End","pnlGradientRgb_Base","pnlGradientRgb_Invert","pnlGradientRgb_Overlay1","pnlGradientRgb_Overlay2","src","imgGradient","../Images/cpns_ColorSpace1.png","../Images/cpns_ColorSpace2.png","../Images/cpns_Vertical1.png","#000000","../Images/cpns_Vertical2.png","#ffffff","01234567879","which","abcdef","01234567879ABCDEF","opener","pnlGradientPosition","pnlNewColor","0123456789ABCDEFabcdef","000000","0","id","top","pnlVerticalPosition","backgroundImage","url(../Images/cpns_GradientPositionDark.gif)","url(../Images/cpns_GradientPositionLight.gif)","cancelBubble","pageX","pageY","className","GradientNormal","GradientFullScreen","_isverdown","=","; path=/;"," expires=",";","cookie","search","location","\x26","00336699CCFF","0x","do_select","frm","__cphex"];var POSITIONADJUSTX=22;var POSITIONADJUSTY=52;var POSITIONADJUSTZ=48;var ColorMode=1;var GradientPositionDark= new Boolean(false);var frm= new Object();var msg= new Object();var _xmlDocs= new Array();var _xmlIndex=-1;var _xml=null;LoadLanguage();window[OxO8818[0]]=window_load;function initialize(){frm[OxO8818[2]][OxO8818[1]]=btnCancel_Click;frm[OxO8818[3]][OxO8818[1]]=btnOK_Click;frm[OxO8818[5]][OxO8818[4]]=Hsb_Changed;frm[OxO8818[5]][OxO8818[6]]=validateNumber;frm[OxO8818[7]][OxO8818[4]]=Hsb_Changed;frm[OxO8818[7]][OxO8818[6]]=validateNumber;frm[OxO8818[8]][OxO8818[4]]=Hsb_Changed;frm[OxO8818[8]][OxO8818[6]]=validateNumber;frm[OxO8818[9]][OxO8818[4]]=Rgb_Changed;frm[OxO8818[9]][OxO8818[6]]=validateNumber;frm[OxO8818[10]][OxO8818[4]]=Rgb_Changed;frm[OxO8818[10]][OxO8818[6]]=validateNumber;frm[OxO8818[11]][OxO8818[4]]=Rgb_Changed;frm[OxO8818[11]][OxO8818[6]]=validateNumber;frm[OxO8818[12]][OxO8818[4]]=Hex_Changed;frm[OxO8818[12]][OxO8818[6]]=validateHex;frm[OxO8818[13]][OxO8818[1]]=btnWebSafeColor_Click;frm[OxO8818[14]][OxO8818[1]]=rdoHsb_Hue_Click;frm[OxO8818[15]][OxO8818[1]]=rdoHsb_Saturation_Click;frm[OxO8818[16]][OxO8818[1]]=rdoHsb_Brightness_Click;document.getElementById(OxO8818[17])[OxO8818[1]]=pnlGradient_Top_Click;document.getElementById(OxO8818[17])[OxO8818[18]]=pnlGradient_Top_MouseMove;document.getElementById(OxO8818[17])[OxO8818[19]]=pnlGradient_Top_MouseDown;document.getElementById(OxO8818[17])[OxO8818[20]]=pnlGradient_Top_MouseUp;document.getElementById(OxO8818[21])[OxO8818[1]]=pnlVertical_Top_Click;document.getElementById(OxO8818[21])[OxO8818[18]]=pnlVertical_Top_MouseMove;document.getElementById(OxO8818[21])[OxO8818[19]]=pnlVertical_Top_MouseDown;document.getElementById(OxO8818[21])[OxO8818[20]]=pnlVertical_Top_MouseUp;document.getElementById(OxO8818[22])[OxO8818[1]]=btnWebSafeColor_Click;document.getElementById(OxO8818[23])[OxO8818[1]]=btnWebSafeColor_Click;document.getElementById(OxO8818[24])[OxO8818[1]]=pnlOldClick_Click;document.getElementById(OxO8818[25])[OxO8818[1]]=rdoHsb_Hue_Click;document.getElementById(OxO8818[26])[OxO8818[1]]=rdoHsb_Saturation_Click;document.getElementById(OxO8818[27])[OxO8818[1]]=rdoHsb_Brightness_Click;frm[OxO8818[5]].focus();window.focus();} ;function formatString(Ox4ed){Ox4ed= new String(Ox4ed);for(var i=1;i<arguments[OxO8818[28]];i++){Ox4ed=Ox4ed.replace( new RegExp(OxO8818[29]+(i-1)+OxO8818[30]),arguments[i]);} ;return Ox4ed;} ;function AddValue(Ox4ef,Oxb9){Oxb9= new String(Oxb9).toLowerCase();for(var i=0;i<Ox4ef[OxO8818[28]];i++){if(Ox4ef[i]==Oxb9){return ;} ;} ;Ox4ef[Ox4ef[OxO8818[28]]]=Oxb9;} ;function SniffLanguage(Ox60){} ;function LoadNextLanguage(){} ;function LoadLanguage(){msg[OxO8818[31]]=OxO8818[32];msg[OxO8818[33]]=OxO8818[34];msg[OxO8818[35]]=OxO8818[36];msg[OxO8818[37]]=OxO8818[38];msg[OxO8818[39]]=OxO8818[40];msg[OxO8818[41]]=OxO8818[42];msg[OxO8818[43]]=OxO8818[43];msg[OxO8818[44]]=OxO8818[45];msg[OxO8818[46]]=OxO8818[47];msg[OxO8818[48]]=OxO8818[49];msg[OxO8818[50]]=OxO8818[51];msg[OxO8818[52]]=OxO8818[53];msg[OxO8818[54]]=OxO8818[55];msg[OxO8818[56]]=OxO8818[57];msg[OxO8818[58]]=OxO8818[53];msg[OxO8818[59]]=OxO8818[60];msg[OxO8818[61]]=OxO8818[62];msg[OxO8818[63]]=OxO8818[53];msg[OxO8818[64]]=OxO8818[65];msg[OxO8818[66]]=OxO8818[67];msg[OxO8818[68]]=OxO8818[69];msg[OxO8818[70]]=OxO8818[71];msg[OxO8818[72]]=OxO8818[60];msg[OxO8818[73]]=OxO8818[74];msg[OxO8818[75]]=OxO8818[76];msg[OxO8818[77]]=OxO8818[78];msg[OxO8818[42]]=OxO8818[79];} ;function AssignLanguage(){} ;function localize(){SetHTML(document.getElementById(OxO8818[80]),msg.SelectAColor,document.getElementById(OxO8818[81]),msg.Recent,document.getElementById(OxO8818[25]),msg.HsbHue,document.getElementById(OxO8818[26]),msg.HsbSaturation,document.getElementById(OxO8818[27]),msg.HsbBrightness,document.getElementById(OxO8818[82]),msg.RgbRed,document.getElementById(OxO8818[83]),msg.RgbGreen,document.getElementById(OxO8818[84]),msg.RgbBlue,document.getElementById(OxO8818[85]),msg.Hex,document.getElementById(OxO8818[86]),msg.HsbHueUnit,document.getElementById(OxO8818[87]),msg.HsbSaturationUnit,document.getElementById(OxO8818[88]),msg.HsbBrightnessUnit);SetValue(frm.btnCancel,msg.CancelButton,frm.btnOK,msg.OKButton,frm.btnAbout,msg.AboutButton);SetTitle(frm.btnWebSafeColor,msg.WebSafeWarning,document.getElementById(OxO8818[22]),msg.WebSafeClick,document.getElementById(OxO8818[89]),msg.HsbHueTooltip,document.getElementById(OxO8818[90]),msg.HsbSaturationTooltip,document.getElementById(OxO8818[91]),msg.HsbBrightnessTooltip,document.getElementById(OxO8818[92]),msg.RgbRedTooltip,document.getElementById(OxO8818[93]),msg.RgbGreenTooltip,document.getElementById(OxO8818[94]),msg.RgbBlueTooltip);} ;function window_load(Ox3e){frm=document.getElementById(OxO8818[95]);localize();initialize();var hex=GetQuery(OxO8818[96]).toUpperCase();if(hex==OxO8818[97]){hex=OxO8818[98];} ;if(hex[OxO8818[28]]==7){hex=hex.substr(1,6);} ;frm[OxO8818[12]][OxO8818[99]]=hex;Hex_Changed(Ox3e);hex=Form_Get_Hex();SetBg(document.getElementById(OxO8818[24]),hex);frm[OxO8818[102]][ new Number(GetCookie(OxO8818[101])||0)][OxO8818[100]]=true;ColorMode_Changed(Ox3e);var Ox4e2=GetCookie(OxO8818[103])||OxO8818[97];var Ox4f4=msg[OxO8818[77]];for(var i=1;i<33;i++){if(Ox4e2[OxO8818[28]]/6>=i){hex=Ox4e2.substr((i-1)*6,6);var Ox4f5=HexToRgb(hex);var title=formatString(msg.RecentTooltip,hex,Ox4f5[0],Ox4f5[1],Ox4f5[2]);SetBg(document.getElementById(OxO8818[104]+i),hex);SetTitle(document.getElementById(OxO8818[104]+i),title);document.getElementById(OxO8818[104]+i)[OxO8818[1]]=pnlRecent_Click;} else {document.getElementById(OxO8818[104]+i)[OxO8818[106]][OxO8818[105]]=OxO8818[107];} ;} ;} ;function btnAbout_Click(){if(confirm(msg.About)){window.open(OxO8818[108],OxO8818[109]);} ;} ;function pnlRecent_Click(Ox3e){var Ox37a=Ox3e[OxO8818[111]][OxO8818[106]][OxO8818[110]];if(Ox37a.indexOf(OxO8818[112])!=-1){var Ox4f5= new Array();Ox37a=Ox37a.substr(Ox37a.indexOf(OxO8818[113])+1);Ox37a=Ox37a.substr(0,Ox37a.indexOf(OxO8818[114]));Ox4f5[0]= new Number(Ox37a.substr(0,Ox37a.indexOf(OxO8818[115])));Ox37a=Ox37a.substr(Ox37a.indexOf(OxO8818[115])+1);Ox4f5[1]= new Number(Ox37a.substr(0,Ox37a.indexOf(OxO8818[115])));Ox4f5[2]= new Number(Ox37a.substr(Ox37a.indexOf(OxO8818[115])+1));Ox37a=RgbToHex(Ox4f5);} else {Ox37a=Ox37a.substr(1,6).toUpperCase();} ;frm[OxO8818[12]][OxO8818[99]]=Ox37a;Hex_Changed(Ox3e);} ;function pnlOldClick_Click(Ox3e){frm[OxO8818[12]][OxO8818[99]]=document.getElementById(OxO8818[24])[OxO8818[106]][OxO8818[110]].substr(1,6).toUpperCase();Hex_Changed(Ox3e);} ;function rdoHsb_Hue_Click(Ox3e){frm[OxO8818[14]][OxO8818[100]]=true;ColorMode_Changed(Ox3e);} ;function rdoHsb_Saturation_Click(Ox3e){frm[OxO8818[15]][OxO8818[100]]=true;ColorMode_Changed(Ox3e);} ;function rdoHsb_Brightness_Click(Ox3e){frm[OxO8818[16]][OxO8818[100]]=true;ColorMode_Changed(Ox3e);} ;function Hide(){for(var i=0;i<arguments[OxO8818[28]];i++){if(arguments[i]){arguments[i][OxO8818[106]][OxO8818[116]]=OxO8818[117];} ;} ;} ;function Show(){for(var i=0;i<arguments[OxO8818[28]];i++){if(arguments[i]){arguments[i][OxO8818[106]][OxO8818[116]]=OxO8818[97];} ;} ;} ;function SetValue(){for(var i=0;i<arguments[OxO8818[28]];i+=2){arguments[i][OxO8818[99]]=arguments[i+1];} ;} ;function SetTitle(){for(var i=0;i<arguments[OxO8818[28]];i+=2){arguments[i][OxO8818[118]]=arguments[i+1];} ;} ;function SetHTML(){for(var i=0;i<arguments[OxO8818[28]];i+=2){arguments[i][OxO8818[119]]=arguments[i+1];} ;} ;function SetBg(){for(var i=0;i<arguments[OxO8818[28]];i+=2){if(arguments[i]){arguments[i][OxO8818[106]][OxO8818[110]]=OxO8818[76]+arguments[i+1];} ;} ;} ;function SetBgPosition(){for(var i=0;i<arguments[OxO8818[28]];i+=3){arguments[i][OxO8818[106]][OxO8818[120]]=arguments[i+1]+OxO8818[121]+arguments[i+2]+OxO8818[122];} ;} ;function ColorMode_Changed(Ox3e){for(var i=0;i<3;i++){if(frm[OxO8818[102]][i][OxO8818[100]]){ColorMode=i;} ;} ;SetCookie(OxO8818[101],ColorMode,60*60*24*365);Hide(document.getElementById(OxO8818[123]),document.getElementById(OxO8818[124]),document.getElementById(OxO8818[125]),document.getElementById(OxO8818[126]),document.getElementById(OxO8818[127]),document.getElementById(OxO8818[128]),document.getElementById(OxO8818[129]),document.getElementById(OxO8818[130]),document.getElementById(OxO8818[131]),document.getElementById(OxO8818[132]),document.getElementById(OxO8818[133]),document.getElementById(OxO8818[134]),document.getElementById(OxO8818[135]),document.getElementById(OxO8818[136]));switch(ColorMode){case 0:document.getElementById(OxO8818[138])[OxO8818[137]]=OxO8818[139];Show(document.getElementById(OxO8818[123]),document.getElementById(OxO8818[124]),document.getElementById(OxO8818[125]),document.getElementById(OxO8818[126]));Hsb_Changed(Ox3e);break ;;case 1:document.getElementById(OxO8818[138])[OxO8818[137]]=OxO8818[140];document.getElementById(OxO8818[127])[OxO8818[137]]=OxO8818[141];Show(document.getElementById(OxO8818[123]),document.getElementById(OxO8818[127]));document.getElementById(OxO8818[123])[OxO8818[106]][OxO8818[110]]=OxO8818[142];Hsb_Changed(Ox3e);break ;;case 2:document.getElementById(OxO8818[138])[OxO8818[137]]=OxO8818[140];document.getElementById(OxO8818[127])[OxO8818[137]]=OxO8818[143];Show(document.getElementById(OxO8818[123]),document.getElementById(OxO8818[127]));document.getElementById(OxO8818[123])[OxO8818[106]][OxO8818[110]]=OxO8818[144];Hsb_Changed(Ox3e);break ;;default:break ;;} ;} ;function btnWebSafeColor_Click(Ox3e){var Ox4f5=HexToRgb(frm[OxO8818[12]].value);Ox4f5=RgbToWebSafeRgb(Ox4f5);frm[OxO8818[12]][OxO8818[99]]=RgbToHex(Ox4f5);Hex_Changed(Ox3e);} ;function checkWebSafe(){var Ox4f5=Form_Get_Rgb();if(RgbIsWebSafe(Ox4f5)){Hide(frm.btnWebSafeColor,document.getElementById(OxO8818[22]),document.getElementById(OxO8818[23]));} else {Ox4f5=RgbToWebSafeRgb(Ox4f5);SetBg(document.getElementById(OxO8818[22]),RgbToHex(Ox4f5));Show(frm.btnWebSafeColor,document.getElementById(OxO8818[22]),document.getElementById(OxO8818[23]));} ;} ;function validateNumber(Ox3e){var Ox14a=String.fromCharCode(Ox3e.which);if(IgnoreKey(Ox3e)){return ;} ;if(OxO8818[145].indexOf(Ox14a)!=-1){return ;} ;Ox3e[OxO8818[146]]=0;} ;function validateHex(Ox3e){if(IgnoreKey(Ox3e)){return ;} ;var Ox14a=String.fromCharCode(Ox3e.which);if(OxO8818[147].indexOf(Ox14a)!=-1){return ;} ;if(OxO8818[148].indexOf(Ox14a)!=-1){return ;} ;} ;function IgnoreKey(Ox3e){var Ox14a=String.fromCharCode(Ox3e.which);var Ox157= new Array(0,8,9,13,27);if(Ox14a==null){return true;} ;for(var i=0;i<5;i++){if(Ox3e[OxO8818[146]]==Ox157[i]){return true;} ;} ;return false;} ;function btnCancel_Click(){if(window[OxO8818[149]]){window[OxO8818[149]].focus();} ;top.close();} ;function btnOK_Click(){var hex= new String(frm[OxO8818[12]].value);if(window[OxO8818[149]]){try{window[OxO8818[149]].ColorPicker_Picked(hex);} catch(e){} ;window[OxO8818[149]].focus();} ;recent=GetCookie(OxO8818[103])||OxO8818[97];for(var i=0;i<recent[OxO8818[28]];i+=6){if(recent.substr(i,6)==hex){recent=recent.substr(0,i)+recent.substr(i+6);i-=6;} ;} ;if(recent[OxO8818[28]]>31*6){recent=recent.substr(0,31*6);} ;recent=frm[OxO8818[12]][OxO8818[99]]+recent;SetCookie(OxO8818[103],recent,60*60*24*365);top.close();} ;function SetGradientPosition(Ox3e,x,y){x=x-POSITIONADJUSTX+5;y=y-POSITIONADJUSTY+5;x-=7;y-=27;x=x<0?0:x>255?255:x;y=y<0?0:y>255?255:y;SetBgPosition(document.getElementById(OxO8818[150]),x-5,y-5);switch(ColorMode){case 0:var Ox50f= new Array(0,0,0);Ox50f[1]=x/255;Ox50f[2]=1-(y/255);frm[OxO8818[7]][OxO8818[99]]=Math.round(Ox50f[1]*100);frm[OxO8818[8]][OxO8818[99]]=Math.round(Ox50f[2]*100);Hsb_Changed(Ox3e);break ;;case 1:var Ox50f= new Array(0,0,0);Ox50f[0]=x/255;Ox50f[2]=1-(y/255);frm[OxO8818[5]][OxO8818[99]]=Ox50f[0]==1?0:Math.round(Ox50f[0]*360);frm[OxO8818[8]][OxO8818[99]]=Math.round(Ox50f[2]*100);Hsb_Changed(Ox3e);break ;;case 2:var Ox50f= new Array(0,0,0);Ox50f[0]=x/255;Ox50f[1]=1-(y/255);frm[OxO8818[5]][OxO8818[99]]=Ox50f[0]==1?0:Math.round(Ox50f[0]*360);frm[OxO8818[7]][OxO8818[99]]=Math.round(Ox50f[1]*100);Hsb_Changed(Ox3e);break ;;} ;} ;function Hex_Changed(Ox3e){var hex=Form_Get_Hex();var Ox4f5=HexToRgb(hex);var Ox50f=RgbToHsb(Ox4f5);Form_Set_Rgb(Ox4f5);Form_Set_Hsb(Ox50f);SetBg(document.getElementById(OxO8818[151]),hex);SetupCursors(Ox3e);SetupGradients();checkWebSafe();} ;function Rgb_Changed(Ox3e){var Ox4f5=Form_Get_Rgb();var Ox50f=RgbToHsb(Ox4f5);var hex=RgbToHex(Ox4f5);Form_Set_Hsb(Ox50f);Form_Set_Hex(hex);SetBg(document.getElementById(OxO8818[151]),hex);SetupCursors(Ox3e);SetupGradients();checkWebSafe();} ;function Hsb_Changed(Ox3e){var Ox50f=Form_Get_Hsb();var Ox4f5=HsbToRgb(Ox50f);var hex=RgbToHex(Ox4f5);Form_Set_Rgb(Ox4f5);Form_Set_Hex(hex);SetBg(document.getElementById(OxO8818[151]),hex);SetupCursors(Ox3e);SetupGradients();checkWebSafe();} ;function Form_Set_Hex(hex){frm[OxO8818[12]][OxO8818[99]]=hex;} ;function Form_Get_Hex(){var hex= new String(frm[OxO8818[12]].value);for(var i=0;i<hex[OxO8818[28]];i++){if(OxO8818[152].indexOf(hex.substr(i,1))==-1){hex=OxO8818[153];frm[OxO8818[12]][OxO8818[99]]=hex;alert(formatString(msg.BadNumber,OxO8818[153],OxO8818[98]));break ;} ;} ;while(hex[OxO8818[28]]<6){hex=OxO8818[154]+hex;} ;return hex;} ;function Form_Get_Hsb(){var Ox50f= new Array(0,0,0);Ox50f[0]= new Number(frm[OxO8818[5]].value)/360;Ox50f[1]= new Number(frm[OxO8818[7]].value)/100;Ox50f[2]= new Number(frm[OxO8818[8]].value)/100;if(Ox50f[0]>1||isNaN(Ox50f[0])){Ox50f[0]=1;frm[OxO8818[5]][OxO8818[99]]=360;alert(formatString(msg.BadNumber,0,360));} ;if(Ox50f[1]>1||isNaN(Ox50f[1])){Ox50f[1]=1;frm[OxO8818[7]][OxO8818[99]]=100;alert(formatString(msg.BadNumber,0,100));} ;if(Ox50f[2]>1||isNaN(Ox50f[2])){Ox50f[2]=1;frm[OxO8818[8]][OxO8818[99]]=100;alert(formatString(msg.BadNumber,0,100));} ;return Ox50f;} ;function Form_Set_Hsb(Ox50f){SetValue(frm.txtHSB_Hue,Math.round(Ox50f[0]*360),frm.txtHSB_Saturation,Math.round(Ox50f[1]*100),frm.txtHSB_Brightness,Math.round(Ox50f[2]*100));} ;function Form_Get_Rgb(){var Ox4f5= new Array(0,0,0);Ox4f5[0]= new Number(frm[OxO8818[9]].value);Ox4f5[1]= new Number(frm[OxO8818[10]].value);Ox4f5[2]= new Number(frm[OxO8818[11]].value);if(Ox4f5[0]>255||isNaN(Ox4f5[0])||Ox4f5[0]!=Math.round(Ox4f5[0])){Ox4f5[0]=255;frm[OxO8818[9]][OxO8818[99]]=255;alert(formatString(msg.BadNumber,0,255));} ;if(Ox4f5[1]>255||isNaN(Ox4f5[1])||Ox4f5[1]!=Math.round(Ox4f5[1])){Ox4f5[1]=255;frm[OxO8818[10]][OxO8818[99]]=255;alert(formatString(msg.BadNumber,0,255));} ;if(Ox4f5[2]>255||isNaN(Ox4f5[2])||Ox4f5[2]!=Math.round(Ox4f5[2])){Ox4f5[2]=255;frm[OxO8818[11]][OxO8818[99]]=255;alert(formatString(msg.BadNumber,0,255));} ;return Ox4f5;} ;function Form_Set_Rgb(Ox4f5){frm[OxO8818[9]][OxO8818[99]]=Ox4f5[0];frm[OxO8818[10]][OxO8818[99]]=Ox4f5[1];frm[OxO8818[11]][OxO8818[99]]=Ox4f5[2];} ;function SetupCursors(Ox3e){var Ox50f=Form_Get_Hsb();var Ox4f5=Form_Get_Rgb();if(RgbToYuv(Ox4f5)[0]>=0.5){SetGradientPositionDark();} else {SetGradientPositionLight();} ;if(Ox3e[OxO8818[111]]!=null){if(Ox3e[OxO8818[111]][OxO8818[155]]==OxO8818[17]){return ;} ;if(Ox3e[OxO8818[111]][OxO8818[155]]==OxO8818[21]){return ;} ;} ;var x;var y;var z;if(ColorMode>=0&&ColorMode<=2){for(var i=0;i<3;i++){Ox50f[i]*=255;} ;} ;switch(ColorMode){case 0:x=Ox50f[1];y=Ox50f[2];z=Ox50f[0]==0?1:Ox50f[0];break ;;case 1:x=Ox50f[0]==0?1:Ox50f[0];y=Ox50f[2];z=Ox50f[1];break ;;case 2:x=Ox50f[0]==0?1:Ox50f[0];y=Ox50f[1];z=Ox50f[2];break ;;} ;y=255-y;z=255-z;SetBgPosition(document.getElementById(OxO8818[150]),x-5,y-5);document.getElementById(OxO8818[157])[OxO8818[106]][OxO8818[156]]=(z+27)+OxO8818[122];} ;function SetupGradients(){var Ox50f=Form_Get_Hsb();var Ox4f5=Form_Get_Rgb();switch(ColorMode){case 0:SetBg(document.getElementById(OxO8818[123]),RgbToHex(HueToRgb(Ox50f[0])));break ;;case 1:SetBg(document.getElementById(OxO8818[127]),RgbToHex(HsbToRgb( new Array(Ox50f[0],1,Ox50f[2]))));break ;;case 2:SetBg(document.getElementById(OxO8818[127]),RgbToHex(HsbToRgb( new Array(Ox50f[0],Ox50f[1],1))));break ;;default:;} ;} ;function SetGradientPositionDark(){if(GradientPositionDark){return ;} ;GradientPositionDark=true;document.getElementById(OxO8818[150])[OxO8818[106]][OxO8818[158]]=OxO8818[159];} ;function SetGradientPositionLight(){if(!GradientPositionDark){return ;} ;GradientPositionDark=false;document.getElementById(OxO8818[150])[OxO8818[106]][OxO8818[158]]=OxO8818[160];} ;function pnlGradient_Top_Click(Ox3e){Ox3e[OxO8818[161]]=true;SetGradientPosition(Ox3e,Ox3e[OxO8818[162]]-5,Ox3e[OxO8818[163]]-5);document.getElementById(OxO8818[17])[OxO8818[164]]=OxO8818[165];_down=false;} ;var _down=false;function pnlGradient_Top_MouseMove(Ox3e){Ox3e[OxO8818[161]]=true;if(!_down){return ;} ;SetGradientPosition(Ox3e,Ox3e[OxO8818[162]]-5,Ox3e[OxO8818[163]]-5);} ;function pnlGradient_Top_MouseDown(Ox3e){Ox3e[OxO8818[161]]=true;_down=true;SetGradientPosition(Ox3e,Ox3e[OxO8818[162]]-5,Ox3e[OxO8818[163]]-5);document.getElementById(OxO8818[17])[OxO8818[164]]=OxO8818[166];} ;function pnlGradient_Top_MouseUp(Ox3e){_down=false;Ox3e[OxO8818[161]]=true;SetGradientPosition(Ox3e,Ox3e[OxO8818[162]]-5,Ox3e[OxO8818[163]]-5);document.getElementById(OxO8818[17])[OxO8818[164]]=OxO8818[165];} ;function Document_MouseUp(){e[OxO8818[161]]=true;document.getElementById(OxO8818[17])[OxO8818[164]]=OxO8818[165];} ;function SetVerticalPosition(Ox3e,z){var z=z-POSITIONADJUSTZ;if(z<27){z=27;} ;if(z>282){z=282;} ;document.getElementById(OxO8818[157])[OxO8818[106]][OxO8818[156]]=z+OxO8818[122];z=1-((z-27)/255);switch(ColorMode){case 0:if(z==1){z=0;} ;frm[OxO8818[5]][OxO8818[99]]=Math.round(z*360);Hsb_Changed(Ox3e);break ;;case 1:frm[OxO8818[7]][OxO8818[99]]=Math.round(z*100);Hsb_Changed(Ox3e);break ;;case 2:frm[OxO8818[8]][OxO8818[99]]=Math.round(z*100);Hsb_Changed(Ox3e);break ;;} ;} ;function pnlVertical_Top_Click(Ox3e){SetVerticalPosition(Ox3e,Ox3e[OxO8818[163]]-5);Ox3e[OxO8818[161]]=true;} ;function pnlVertical_Top_MouseMove(Ox3e){if(!window[OxO8818[167]]){return ;} ;if(Ox3e[OxO8818[146]]!=1){return ;} ;SetVerticalPosition(Ox3e,Ox3e[OxO8818[163]]-5);Ox3e[OxO8818[161]]=true;} ;function pnlVertical_Top_MouseDown(Ox3e){window[OxO8818[167]]=true;SetVerticalPosition(Ox3e,Ox3e[OxO8818[163]]-5);Ox3e[OxO8818[161]]=true;} ;function pnlVertical_Top_MouseUp(Ox3e){window[OxO8818[167]]=false;SetVerticalPosition(Ox3e,Ox3e[OxO8818[163]]-5);Ox3e[OxO8818[161]]=true;} ;function SetCookie(name,Oxb9,Oxba){var Oxbb=name+OxO8818[168]+escape(Oxb9)+OxO8818[169];if(Oxba){var Oxbc= new Date();Oxbc.setSeconds(Oxbc.getSeconds()+Oxba);Oxbb+=OxO8818[170]+Oxbc.toUTCString()+OxO8818[171];} ;document[OxO8818[172]]=Oxbb;} ;function GetCookie(name){var Oxbe=document[OxO8818[172]].split(OxO8818[171]);for(var i=0;i<Oxbe[OxO8818[28]];i++){var Oxbf=Oxbe[i].split(OxO8818[168]);if(name==Oxbf[0].replace(/\s/g,OxO8818[97])){return unescape(Oxbf[1]);} ;} ;} ;function GetCookieDictionary(){var Ox3b8={};var Oxbe=document[OxO8818[172]].split(OxO8818[171]);for(var i=0;i<Oxbe[OxO8818[28]];i++){var Oxbf=Oxbe[i].split(OxO8818[168]);Ox3b8[Oxbf[0].replace(/\s/g,OxO8818[97])]=unescape(Oxbf[1]);} ;return Ox3b8;} ;function GetQuery(name){var i=0;while(window[OxO8818[174]][OxO8818[173]].indexOf(name+OxO8818[168],i)!=-1){var Oxb9=window[OxO8818[174]][OxO8818[173]].substr(window[OxO8818[174]][OxO8818[173]].indexOf(name+OxO8818[168],i));Oxb9=Oxb9.substr(name[OxO8818[28]]+1);if(Oxb9.indexOf(OxO8818[175])!=-1){if(Oxb9.indexOf(OxO8818[175])==0){Oxb9=OxO8818[97];} else {Oxb9=Oxb9.substr(0,Oxb9.indexOf(OxO8818[175]));} ;} ;return unescape(Oxb9);} ;return OxO8818[97];} ;function RgbIsWebSafe(Ox4f5){var hex=RgbToHex(Ox4f5);for(var i=0;i<3;i++){if(OxO8818[176].indexOf(hex.substr(i*2,2))==-1){return false;} ;} ;return true;} ;function RgbToWebSafeRgb(Ox4f5){var Ox529= new Array(Ox4f5[0],Ox4f5[1],Ox4f5[2]);if(RgbIsWebSafe(Ox4f5)){return Ox529;} ;var Ox52a= new Array(0x00,0x33,0x66,0x99,0xCC,0xFF);for(var i=0;i<3;i++){for(var j=1;j<6;j++){if(Ox529[i]>Ox52a[j-1]&&Ox529[i]<Ox52a[j]){if(Ox529[i]-Ox52a[j-1]>Ox52a[j]-Ox529[i]){Ox529[i]=Ox52a[j];} else {Ox529[i]=Ox52a[j-1];} ;break ;} ;} ;} ;return Ox529;} ;function RgbToYuv(Ox4f5){var Ox52c= new Array();Ox52c[0]=(Ox4f5[0]*0.299+Ox4f5[1]*0.587+Ox4f5[2]*0.114)/255;Ox52c[1]=(Ox4f5[0]*-0.169+Ox4f5[1]*-0.332+Ox4f5[2]*0.500+128)/255;Ox52c[2]=(Ox4f5[0]*0.500+Ox4f5[1]*-0.419+Ox4f5[2]*-0.0813+128)/255;return Ox52c;} ;function RgbToHsb(Ox4f5){var Ox52e= new Array(Ox4f5[0],Ox4f5[1],Ox4f5[2]);var Ox52f= new Number(1);var Ox530= new Number(0);var Ox531= new Number(1);var Ox50f= new Array(0,0,0);var Ox532= new Array();for(var i=0;i<3;i++){Ox52e[i]=Ox4f5[i]/255;if(Ox52e[i]<Ox52f){Ox52f=Ox52e[i];} ;if(Ox52e[i]>Ox530){Ox530=Ox52e[i];} ;} ;Ox531=Ox530-Ox52f;Ox50f[2]=Ox530;if(Ox531==0){return Ox50f;} ;Ox50f[1]=Ox531/Ox530;for(var i=0;i<3;i++){Ox532[i]=(((Ox530-Ox52e[i])/6)+(Ox531/2))/Ox531;} ;if(Ox52e[0]==Ox530){Ox50f[0]=Ox532[2]-Ox532[1];} else {if(Ox52e[1]==Ox530){Ox50f[0]=(1/3)+Ox532[0]-Ox532[2];} else {if(Ox52e[2]==Ox530){Ox50f[0]=(2/3)+Ox532[1]-Ox532[0];} ;} ;} ;if(Ox50f[0]<0){Ox50f[0]+=1;} else {if(Ox50f[0]>1){Ox50f[0]-=1;} ;} ;return Ox50f;} ;function HsbToRgb(Ox50f){var Ox4f5=HueToRgb(Ox50f[0]);var Ox398=Ox50f[2]*255;for(var i=0;i<3;i++){Ox4f5[i]=Ox4f5[i]*Ox50f[2];Ox4f5[i]=((Ox4f5[i]-Ox398)*Ox50f[1])+Ox398;Ox4f5[i]=Math.round(Ox4f5[i]);} ;return Ox4f5;} ;function RgbToHex(Ox4f5){var hex= new String();for(var i=0;i<3;i++){Ox4f5[2-i]=Math.round(Ox4f5[2-i]);hex=Ox4f5[2-i].toString(16)+hex;if(hex[OxO8818[28]]%2==1){hex=OxO8818[154]+hex;} ;} ;return hex.toUpperCase();} ;function HexToRgb(hex){var Ox4f5= new Array();for(var i=0;i<3;i++){Ox4f5[i]= new Number(OxO8818[177]+hex.substr(i*2,2));} ;return Ox4f5;} ;function HueToRgb(Ox537){var Ox538=Ox537*360;var Ox4f5= new Array(0,0,0);var Ox539=(Ox538%60)/60;if(Ox538<60){Ox4f5[0]=255;Ox4f5[1]=Ox539*255;} else {if(Ox538<120){Ox4f5[1]=255;Ox4f5[0]=(1-Ox539)*255;} else {if(Ox538<180){Ox4f5[1]=255;Ox4f5[2]=Ox539*255;} else {if(Ox538<240){Ox4f5[2]=255;Ox4f5[1]=(1-Ox539)*255;} else {if(Ox538<300){Ox4f5[2]=255;Ox4f5[0]=Ox539*255;} else {if(Ox538<360){Ox4f5[0]=255;Ox4f5[2]=(1-Ox539)*255;} ;} ;} ;} ;} ;} ;return Ox4f5;} ;function CheckHexSelect(){if(window[OxO8818[178]]&&window[OxO8818[179]]&&frm[OxO8818[12]]){var Ox37a=OxO8818[76]+frm[OxO8818[12]][OxO8818[99]];if(Ox37a[OxO8818[28]]==7){if(window[OxO8818[180]]!=Ox37a){window[OxO8818[180]]=Ox37a;window.do_select(Ox37a);} ;} ;} ;} ;setInterval(CheckHexSelect,10);