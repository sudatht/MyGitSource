var OxOd494=["value","","onload","upload_image","contentWindow","browse_Frame","height","style","250px","btn_CreateDir","btn_zoom_in","btn_zoom_out","btn_Actualsize","TargetUrl","framepreview","innerHTML","HTML","document","body","DIV","innerText","?","\x26#",";","zoom","wrapupPrompt","iepromptfield","display","none","div","id","IEPromptBox","promptBlackout","border","1px solid #b0bec7","backgroundColor","#f0f0f0","position","absolute","width","330px","zIndex","100","\x3Cdiv style=\x22width: 100%; padding-top:3px;background-color: #DCE7EB; font-family: verdana; font-size: 10pt; font-weight: bold; height: 22px; text-align:center; background:url(../Images/formbg2.gif) repeat-x left top;\x22\x3E","\x3C/div\x3E","\x3Cdiv style=\x22padding: 10px\x22\x3E","\x3CBR\x3E\x3CBR\x3E","\x3Cform action=\x22\x22 onsubmit=\x22return wrapupPrompt()\x22\x3E","\x3Cinput id=\x22iepromptfield\x22 name=\x22iepromptdata\x22 type=text size=46 value=\x22","\x22\x3E","\x3Cbr\x3E\x3Cbr\x3E\x3Ccenter\x3E","\x3Cinput type=\x22submit\x22 value=\x22\x26nbsp;\x26nbsp;\x26nbsp;","\x26nbsp;\x26nbsp;\x26nbsp;\x22\x3E","\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;\x26nbsp;","\x3Cinput type=\x22button\x22 onclick=\x22wrapupPrompt(true)\x22 value=\x22\x26nbsp;","\x26nbsp;\x22\x3E","\x3C/form\x3E\x3C/div\x3E","top","100px","left","offsetWidth","px","block","onmouseover","CuteEditor_ColorPicker_ButtonOver(this)"];setMouseOver();function setMouseOver(){} ;function ResetFields(){TargetUrl[OxOd494[0]]=OxOd494[1];} ;function reset_hiddens(){} ;Event_Attach(window,OxOd494[2],reset_hiddens);var upload_image=Window_GetElement(window,OxOd494[3],true);upload_image=upload_image[OxOd494[4]];var browse_Frame=Window_GetElement(window,OxOd494[5],true);if(!Browser_IsWinIE()){browse_Frame[OxOd494[7]][OxOd494[6]]=OxOd494[8];} ;browse_Frame=browse_Frame[OxOd494[4]];var btn_CreateDir=Window_GetElement(window,OxOd494[9],true);var btn_zoom_in=Window_GetElement(window,OxOd494[10],true);var btn_zoom_out=Window_GetElement(window,OxOd494[11],true);var btn_Actualsize=Window_GetElement(window,OxOd494[12],true);var TargetUrl=Window_GetElement(window,OxOd494[13],true);var framepreview=document.getElementById(OxOd494[14])[OxOd494[4]];var editor=Window_GetDialogArguments(window);var htmlcode=OxOd494[1];function do_preview(){try{htmlcode=framepreview[OxOd494[17]].getElementsByTagName(OxOd494[16])[0][OxOd494[15]];} catch(er){htmlcode=framepreview[OxOd494[17]][OxOd494[18]][OxOd494[15]];var div=document.createElement(OxOd494[19]);div[OxOd494[15]]=htmlcode;htmlcode=div[OxOd494[20]];} ;} ;function do_insert(){var Ox140=TargetUrl[OxOd494[0]];if(Ox140.indexOf(OxOd494[21])!=-1){htmlcode=framepreview[OxOd494[17]][OxOd494[18]][OxOd494[15]];} ;htmlcode=htmlcode.replace(/[\u00A0-\u00FF|\u00FF-\uFFFF]/g,function (Oxac,b,Ox29){return OxOd494[22]+Oxac.charCodeAt(0)+OxOd494[23];} );editor.PasteHTML(htmlcode);Window_CloseDialog(window);} ;function do_Close(){Window_CloseDialog(window);} ;function Zoom_In(){if(framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]!=0){framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]*=1.1;} else {framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]=1.1;} ;} ;function Zoom_Out(){if(framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]!=0){framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]*=0.8;} else {framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]=0.8;} ;} ;function Actualsize(){framepreview[OxOd494[17]][OxOd494[18]][OxOd494[7]][OxOd494[24]]=1;do_preview(htmlcode);} ;if(Browser_IsIE7()){var _dialogPromptID=null;function IEprompt(Ox17,Ox463,Ox464){that=this;this[OxOd494[25]]=function (Ox465){val=document.getElementById(OxOd494[26])[OxOd494[0]];_dialogPromptID[OxOd494[7]][OxOd494[27]]=OxOd494[28];document.getElementById(OxOd494[26])[OxOd494[0]]=OxOd494[1];if(Ox465){val=OxOd494[1];} ;Ox17(val);return false;} ;if(Ox464==undefined){Ox464=OxOd494[1];} ;if(_dialogPromptID==null){var Ox466=document.getElementsByTagName(OxOd494[18])[0];tnode=document.createElement(OxOd494[29]);tnode[OxOd494[30]]=OxOd494[31];Ox466.appendChild(tnode);_dialogPromptID=document.getElementById(OxOd494[31]);tnode=document.createElement(OxOd494[29]);tnode[OxOd494[30]]=OxOd494[32];Ox466.appendChild(tnode);_dialogPromptID[OxOd494[7]][OxOd494[33]]=OxOd494[34];_dialogPromptID[OxOd494[7]][OxOd494[35]]=OxOd494[36];_dialogPromptID[OxOd494[7]][OxOd494[37]]=OxOd494[38];_dialogPromptID[OxOd494[7]][OxOd494[39]]=OxOd494[40];_dialogPromptID[OxOd494[7]][OxOd494[41]]=OxOd494[42];} ;var Ox467=OxOd494[43]+InputRequired+OxOd494[44];Ox467+=OxOd494[45]+Ox463+OxOd494[46];Ox467+=OxOd494[47];Ox467+=OxOd494[48]+Ox464+OxOd494[49];Ox467+=OxOd494[50];Ox467+=OxOd494[51]+OK+OxOd494[52];Ox467+=OxOd494[53];Ox467+=OxOd494[54]+Cancel+OxOd494[55];Ox467+=OxOd494[56];_dialogPromptID[OxOd494[15]]=Ox467;_dialogPromptID[OxOd494[7]][OxOd494[57]]=OxOd494[58];_dialogPromptID[OxOd494[7]][OxOd494[59]]=parseInt((document[OxOd494[18]][OxOd494[60]]-315)/2)+OxOd494[61];_dialogPromptID[OxOd494[7]][OxOd494[27]]=OxOd494[62];var Ox468=document.getElementById(OxOd494[26]);try{var Ox469=Ox468.createTextRange();Ox469.collapse(false);Ox469.select();} catch(x){Ox468.focus();} ;} ;} ;if(!Browser_IsWinIE()){btn_zoom_in[OxOd494[7]][OxOd494[27]]=btn_zoom_out[OxOd494[7]][OxOd494[27]]=btn_Actualsize[OxOd494[7]][OxOd494[27]]=OxOd494[28];} ;if(btn_CreateDir){btn_CreateDir[OxOd494[63]]= new Function(OxOd494[64]);} ;if(btn_zoom_in){btn_zoom_in[OxOd494[63]]= new Function(OxOd494[64]);} ;if(btn_zoom_out){btn_zoom_out[OxOd494[63]]= new Function(OxOd494[64]);} ;if(btn_Actualsize){btn_Actualsize[OxOd494[63]]= new Function(OxOd494[64]);} ;