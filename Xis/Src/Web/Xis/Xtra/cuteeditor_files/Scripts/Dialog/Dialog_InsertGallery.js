var OxO882b=["upbtn","browse_Img_gallery","TargetUrl","onclick","location","src","onmouseover","upload_image","upload.asp?","\x26FP=","\x26Type=Image","value","lightyellow","0px","-3px","all","getElementById","\x3Cdiv id=\x22tooltipdiv\x22 style=\x22visibility:hidden;background-color:","\x22 \x3E\x3C/div\x3E","tooltipdiv","left","offsetLeft","offsetTop","offsetParent","style","top","visibility","compatMode","BackCompat","documentElement","body","rightedge","opera","scrollLeft","clientWidth","pageXOffset","innerWidth","contentmeasure","offsetWidth","x","scrollTop","clientHeight","pageYOffset","innerHeight","offsetHeight","y","innerHTML","visible","hidden","px","bottomedge","undefined","hidetip()","element","editor","editdoc","^[a-z]*:[/][/][^/]*","","width","height","IMG","border","alt","product","Gecko","src_cetemp","Edit"];var upbtn=Window_GetElement(window,OxO882b[0],true);var browse_Img_gallery=Window_GetElement(window,OxO882b[1],true);var TargetUrl=Window_GetElement(window,OxO882b[2],true);upbtn[OxO882b[3]]=function upbtn_onclick(){if(Browser_IsOpera()){browse_Img_gallery[OxO882b[4]]=currentfolder;} else {browse_Img_gallery[OxO882b[5]]=currentfolder;} ;} ;upbtn[OxO882b[6]]=function upbtn_onclick(){CuteEditor_ColorPicker_ButtonOver(this);} ;function SetUpload_imagePath(Oxd){if(document.getElementById(OxO882b[7])){document.getElementById(OxO882b[7])[OxO882b[5]]=OxO882b[8]+setting+OxO882b[9]+Oxd+OxO882b[10];} ;} ;function row_click(Oxd){TargetUrl[OxO882b[11]]=Oxd;} ;function cancel(){Window_CloseDialog(window);} ;var tipbgcolor=OxO882b[12];var disappeardelay=250;var vertical_offset=OxO882b[13];var horizontal_offset=OxO882b[14];var delayhidetimerid;var ie4=document[OxO882b[15]];var ns6=document[OxO882b[16]]&&!document[OxO882b[15]];if(ie4||ns6){document.write(OxO882b[17]+tipbgcolor+OxO882b[18]);var dropmenuobj=Window_GetElement(window,OxO882b[19],true);} ;function getposOffset(Ox38,Ox39){var Ox3a=(Ox39==OxO882b[20])?Ox38[OxO882b[21]]:Ox38[OxO882b[22]];var Ox3b=Ox38[OxO882b[23]];while(Ox3b!=null){Ox3a+=(Ox39==OxO882b[20])?Ox3b[OxO882b[21]]:Ox3b[OxO882b[22]];Ox3b=Ox3b[OxO882b[23]];} ;return Ox3a;} ;function showhide(obj,Ox3f,Ox40){if(ie4||ns6){dropmenuobj[OxO882b[24]][OxO882b[20]]=dropmenuobj[OxO882b[24]][OxO882b[25]]=-500;} ;obj[OxO882b[26]]=Ox3f;} ;function iecompattest(){return (document[OxO882b[27]]&&document[OxO882b[27]]!=OxO882b[28])?document[OxO882b[29]]:document[OxO882b[30]];} ;function clearbrowseredge(obj,Ox43){var Ox44=(Ox43==OxO882b[31])?parseInt(horizontal_offset)*-1:parseInt(vertical_offset)*-1;if(Ox43==OxO882b[31]){var Ox45=ie4&&!window[OxO882b[32]]?iecompattest()[OxO882b[33]]+iecompattest()[OxO882b[34]]-15:window[OxO882b[35]]+window[OxO882b[36]]-15;dropmenuobj[OxO882b[37]]=dropmenuobj[OxO882b[38]];if(Ox45-dropmenuobj[OxO882b[39]]<dropmenuobj[OxO882b[37]]){Ox44=dropmenuobj[OxO882b[37]]-obj[OxO882b[38]];} ;} else {var Ox45=ie4&&!window[OxO882b[32]]?iecompattest()[OxO882b[40]]+iecompattest()[OxO882b[41]]-15:window[OxO882b[42]]+window[OxO882b[43]]-18;dropmenuobj[OxO882b[37]]=dropmenuobj[OxO882b[44]];if(Ox45-dropmenuobj[OxO882b[45]]<dropmenuobj[OxO882b[37]]){Ox44=dropmenuobj[OxO882b[37]]+obj[OxO882b[44]];} ;} ;return Ox44;} ;function showTooltip(Ox47,obj){Event_CancelEvent();clearhidetip();dropmenuobj[OxO882b[46]]=Ox47;if(ie4||ns6){showhide(dropmenuobj.style,OxO882b[47],OxO882b[48]);dropmenuobj[OxO882b[39]]=getposOffset(obj,OxO882b[20]);dropmenuobj[OxO882b[45]]=getposOffset(obj,OxO882b[25]);dropmenuobj[OxO882b[24]][OxO882b[20]]=dropmenuobj[OxO882b[39]]-clearbrowseredge(obj,OxO882b[31])+OxO882b[49];dropmenuobj[OxO882b[24]][OxO882b[25]]=dropmenuobj[OxO882b[45]]-clearbrowseredge(obj,OxO882b[50])+obj[OxO882b[44]]*1.1+2+OxO882b[49];} ;} ;function hidetip(){if( typeof dropmenuobj!=OxO882b[51]){if(ie4||ns6){dropmenuobj[OxO882b[24]][OxO882b[26]]=OxO882b[48];} ;} ;} ;function delayhidetip(){if(ie4||ns6){delayhidetimerid=setTimeout(OxO882b[52],disappeardelay);} ;} ;function clearhidetip(){clearTimeout(delayhidetimerid);} ;function cancel(){Window_CloseDialog(window);} ;var obj=Window_GetDialogArguments(window);var element=obj[OxO882b[53]];var editor=obj[OxO882b[54]];var editdoc=obj[OxO882b[55]];function insert(src){if(src){var Ox4c2=src.replace( new RegExp(OxO882b[56],OxO882b[57]),OxO882b[57]);function Actualsize(){var Ox601= new Image();Ox601[OxO882b[5]]=Ox4c2;if(Ox601[OxO882b[58]]>0&&Ox601[OxO882b[59]]>0){element[OxO882b[58]]=Ox601[OxO882b[58]];element[OxO882b[59]]=Ox601[OxO882b[59]];} else {setTimeout(Actualsize,400);} ;} ;if(element){element[OxO882b[5]]=Ox4c2;} else {element=editdoc.createElement(OxO882b[60]);element[OxO882b[5]]=Ox4c2;element[OxO882b[61]]=0;element[OxO882b[62]]=OxO882b[57];Actualsize();} ;if(navigator[OxO882b[63]]==OxO882b[64]){try{element.setAttribute(OxO882b[65],Ox4c2);} catch(e){} ;} else {if(editor.GetActiveTab()==OxO882b[66]){element.setAttribute(OxO882b[65],Ox4c2);} ;} ;editor.InsertElement(element);Window_CloseDialog(window);} ;} ;