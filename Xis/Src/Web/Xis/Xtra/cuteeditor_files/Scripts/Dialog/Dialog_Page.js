var OxO9cf8=["Table1","Table2","inp_title","inp_doctype","inp_description","inp_keywords","PageLanguage","HTMLEncoding","bgcolor","bgcolor_Preview","fontcolor","fontcolor_Preview","Backgroundimage","btnbrowse","TopMargin","BottomMargin","LeftMargin","RightMargin","MarginWidth","MarginHeight","btnok","btncc","editor","window","document","body","head","title","value","innerHTML","DOCTYPE","meta","length","name","keywords","content","description","httpEquiv","content-type","content-language","background","color","style","backgroundColor","bgColor","topMargin","bottomMargin","leftMargin","rightMargin","marginWidth","marginHeight","onclick","","ValidNumber","Please enter a valid color number.","text","childNodes","parentNode","http-equiv","Content-Type","Content-Language","\x3CMETA ","=\x22","\x22 CONTENT=\x22","\x22\x3E","META"];var Table1=Window_GetElement(window,OxO9cf8[0],true);var Table2=Window_GetElement(window,OxO9cf8[1],true);var inp_title=Window_GetElement(window,OxO9cf8[2],true);var inp_doctype=Window_GetElement(window,OxO9cf8[3],true);var inp_description=Window_GetElement(window,OxO9cf8[4],true);var inp_keywords=Window_GetElement(window,OxO9cf8[5],true);var PageLanguage=Window_GetElement(window,OxO9cf8[6],true);var HTMLEncoding=Window_GetElement(window,OxO9cf8[7],true);var bgcolor=Window_GetElement(window,OxO9cf8[8],true);var bgcolor_Preview=Window_GetElement(window,OxO9cf8[9],true);var fontcolor=Window_GetElement(window,OxO9cf8[10],true);var fontcolor_Preview=Window_GetElement(window,OxO9cf8[11],true);var Backgroundimage=Window_GetElement(window,OxO9cf8[12],true);var btnbrowse=Window_GetElement(window,OxO9cf8[13],true);var TopMargin=Window_GetElement(window,OxO9cf8[14],true);var BottomMargin=Window_GetElement(window,OxO9cf8[15],true);var LeftMargin=Window_GetElement(window,OxO9cf8[16],true);var RightMargin=Window_GetElement(window,OxO9cf8[17],true);var MarginWidth=Window_GetElement(window,OxO9cf8[18],true);var MarginHeight=Window_GetElement(window,OxO9cf8[19],true);var btnok=Window_GetElement(window,OxO9cf8[20],true);var btncc=Window_GetElement(window,OxO9cf8[21],true);var obj=Window_GetDialogArguments(window);var editor=obj[OxO9cf8[22]];var editwin=obj[OxO9cf8[23]];var editdoc=obj[OxO9cf8[24]];var body=editdoc[OxO9cf8[25]];var head=obj[OxO9cf8[26]];var title=head.getElementsByTagName(OxO9cf8[27])[0];if(title){inp_title[OxO9cf8[28]]=title[OxO9cf8[29]];} ;inp_doctype[OxO9cf8[28]]=obj[OxO9cf8[30]];var metas=head.getElementsByTagName(OxO9cf8[31]);for(var m=0;m<metas[OxO9cf8[32]];m++){if(metas[m][OxO9cf8[33]].toLowerCase()==OxO9cf8[34]){inp_keywords[OxO9cf8[28]]=metas[m][OxO9cf8[35]];} ;if(metas[m][OxO9cf8[33]].toLowerCase()==OxO9cf8[36]){inp_description[OxO9cf8[28]]=metas[m][OxO9cf8[35]];} ;if(metas[m][OxO9cf8[37]].toLowerCase()==OxO9cf8[38]){HTMLEncoding[OxO9cf8[28]]=metas[m][OxO9cf8[35]];} ;if(metas[m][OxO9cf8[37]].toLowerCase()==OxO9cf8[39]){PageLanguage[OxO9cf8[28]]=metas[m][OxO9cf8[35]];} ;} ;if(editdoc[OxO9cf8[25]][OxO9cf8[40]]){Backgroundimage[OxO9cf8[28]]=editdoc[OxO9cf8[25]][OxO9cf8[40]];} ;if(editdoc[OxO9cf8[25]][OxO9cf8[42]][OxO9cf8[41]]){fontcolor[OxO9cf8[28]]=editdoc[OxO9cf8[25]][OxO9cf8[42]][OxO9cf8[41]];fontcolor[OxO9cf8[42]][OxO9cf8[43]]=fontcolor[OxO9cf8[28]];fontcolor_Preview[OxO9cf8[42]][OxO9cf8[43]]=fontcolor[OxO9cf8[28]];} ;var body_bgcolor=editdoc[OxO9cf8[25]][OxO9cf8[42]][OxO9cf8[43]]||editdoc[OxO9cf8[25]][OxO9cf8[44]];if(body_bgcolor){bgcolor[OxO9cf8[28]]=body_bgcolor;bgcolor[OxO9cf8[42]][OxO9cf8[43]]=body_bgcolor;bgcolor_Preview[OxO9cf8[42]][OxO9cf8[43]]=body_bgcolor;} ;if(Browser_IsWinIE()){if(body[OxO9cf8[45]]){TopMargin[OxO9cf8[28]]=body[OxO9cf8[45]];} ;if(body[OxO9cf8[46]]){BottomMargin[OxO9cf8[28]]=body[OxO9cf8[46]];} ;if(body[OxO9cf8[47]]){LeftMargin[OxO9cf8[28]]=body[OxO9cf8[47]];} ;if(body[OxO9cf8[48]]){RightMargin[OxO9cf8[28]]=body[OxO9cf8[48]];} ;if(body[OxO9cf8[49]]){MarginWidth[OxO9cf8[28]]=body[OxO9cf8[49]];} ;if(body[OxO9cf8[50]]){MarginHeight[OxO9cf8[28]]=body[OxO9cf8[50]];} ;} else {if(body.getAttribute(OxO9cf8[45])){TopMargin[OxO9cf8[28]]=body.getAttribute(OxO9cf8[45]);} ;if(body.getAttribute(OxO9cf8[46])){BottomMargin[OxO9cf8[28]]=body.getAttribute(OxO9cf8[46]);} ;if(body.getAttribute(OxO9cf8[47])){LeftMargin[OxO9cf8[28]]=body.getAttribute(OxO9cf8[47]);} ;if(body.getAttribute(OxO9cf8[48])){RightMargin[OxO9cf8[28]]=body.getAttribute(OxO9cf8[48]);} ;if(body.getAttribute(OxO9cf8[18])){MarginWidth[OxO9cf8[28]]=body.getAttribute(OxO9cf8[49]);} ;if(body.getAttribute(OxO9cf8[50])){MarginHeight[OxO9cf8[28]]=body.getAttribute(OxO9cf8[50]);} ;} ;btnok[OxO9cf8[51]]=function btnok_onclick(){try{if(Browser_IsWinIE()){body[OxO9cf8[45]]=TopMargin[OxO9cf8[28]];body[OxO9cf8[46]]=BottomMargin[OxO9cf8[28]];body[OxO9cf8[47]]=LeftMargin[OxO9cf8[28]];body[OxO9cf8[48]]=RightMargin[OxO9cf8[28]];if(MarginWidth[OxO9cf8[28]]){body[OxO9cf8[49]]=MarginWidth[OxO9cf8[28]];} ;if(MarginHeight[OxO9cf8[28]]){body[OxO9cf8[50]]=MarginHeight[OxO9cf8[28]];} ;} else {body.setAttribute(OxO9cf8[45],TopMargin.value);body.setAttribute(OxO9cf8[46],BottomMargin.value);body.setAttribute(OxO9cf8[47],LeftMargin.value);body.setAttribute(OxO9cf8[48],RightMargin.value);body.setAttribute(OxO9cf8[49],MarginWidth.value);body.setAttribute(OxO9cf8[50],MarginHeight.value);if(body.getAttribute(OxO9cf8[45])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[45]);} ;if(body.getAttribute(OxO9cf8[46])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[46]);} ;if(body.getAttribute(OxO9cf8[47])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[47]);} ;if(body.getAttribute(OxO9cf8[48])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[48]);} ;if(body.getAttribute(OxO9cf8[49])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[49]);} ;if(body.getAttribute(OxO9cf8[50])==OxO9cf8[52]){body.removeAttribute(OxO9cf8[50]);} ;} ;} catch(er){alert(CE_GetStr(OxO9cf8[53]));return ;} ;try{editdoc[OxO9cf8[25]][OxO9cf8[42]][OxO9cf8[43]]=bgcolor[OxO9cf8[28]];editdoc[OxO9cf8[25]][OxO9cf8[42]][OxO9cf8[41]]=fontcolor[OxO9cf8[28]];if(Backgroundimage[OxO9cf8[28]]){editdoc[OxO9cf8[25]][OxO9cf8[40]]=Backgroundimage[OxO9cf8[28]];} else {body.removeAttribute(OxO9cf8[40]);} ;} catch(er){alert(OxO9cf8[54]);return ;} ;if(!title){title=document.createElement(OxO9cf8[27]);head.appendChild(title);} ;if(Browser_IsWinIE()){title[OxO9cf8[55]]=inp_title[OxO9cf8[28]];} else {var Ox64f=document.createTextNode(inp_title.value);try{title.removeChild(title[OxO9cf8[56]].item(0));} catch(x){} ;title.appendChild(Ox64f);} ;for(var m=0;m<metas[OxO9cf8[32]];m++){var Oxb2=metas[m];if(Oxb2){if(Oxb2[OxO9cf8[33]].toLowerCase()==OxO9cf8[34]||Oxb2[OxO9cf8[33]].toLowerCase()==OxO9cf8[36]||Oxb2[OxO9cf8[33]].toLowerCase()==OxO9cf8[38]||Oxb2[OxO9cf8[33]].toLowerCase()==OxO9cf8[39]){Oxb2[OxO9cf8[57]].removeChild(Oxb2);Oxb2=null;} ;} ;} ;try{if(inp_keywords[OxO9cf8[28]]){Ox650(OxO9cf8[33],OxO9cf8[34],inp_keywords.value);} ;if(inp_description[OxO9cf8[28]]){Ox650(OxO9cf8[33],OxO9cf8[36],inp_description.value);} ;if(HTMLEncoding[OxO9cf8[28]]){Ox650(OxO9cf8[58],OxO9cf8[59],HTMLEncoding.value);} ;if(PageLanguage[OxO9cf8[28]]){Ox650(OxO9cf8[58],OxO9cf8[60],PageLanguage.value);} ;} catch(x){} ;function Ox650(Oxdf,name,Ox4d2){var Ox651;if(Browser_IsWinIE()){Ox651=editdoc.createElement(OxO9cf8[61]+Oxdf+OxO9cf8[62]+name+OxO9cf8[63]+Ox4d2+OxO9cf8[64]);} else {var metas=head.getElementsByTagName(OxO9cf8[31]);for(var m=0;m<metas[OxO9cf8[32]];m++){if(metas[m][OxO9cf8[33]].toLowerCase()==name.toLowerCase()){metas[m][OxO9cf8[57]].removeChild(metas[m]);} ;} ;var Ox651=editdoc.createElement(OxO9cf8[65]);Ox651.setAttribute(Oxdf,name);Ox651.setAttribute(OxO9cf8[35],Ox4d2);} ;head.appendChild(Ox651);} ;Window_SetDialogReturnValue(window,inp_doctype[OxO9cf8[28]]+OxO9cf8[52]);Window_CloseDialog(window);} ;btnbrowse[OxO9cf8[51]]=function btnbrowse_onclick(){function Ox595(Ox37c){if(Ox37c){Backgroundimage[OxO9cf8[28]]=Ox37c;} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectImageDialog(Ox595,Backgroundimage.value,Backgroundimage);} else {editor.ShowSelectImageDialog(Ox595,Backgroundimage.value);} ;} ;btncc[OxO9cf8[51]]=function btncc_onclick(){Window_CloseDialog(window);} ;fontcolor[OxO9cf8[51]]=fontcolor_Preview[OxO9cf8[51]]=function fontcolor_onclick(){SelectColor(fontcolor,fontcolor_Preview);} ;bgcolor[OxO9cf8[51]]=bgcolor_Preview[OxO9cf8[51]]=function bgcolor_onclick(){SelectColor(bgcolor,bgcolor_Preview);} ;