var OxO4df1=["inp_width","inp_height","sel_align","sel_valign","inp_bgColor","inp_borderColor","inp_borderColorLight","inp_borderColorDark","inp_class","inp_id","inp_tooltip","value","bgColor","backgroundColor","style","","id","borderColor","borderColorLight","borderColorDark","className","width","height","align","vAlign","title","ValidNumber","ValidID","class","valign","onclick"];var inp_width=Window_GetElement(window,OxO4df1[0],true);var inp_height=Window_GetElement(window,OxO4df1[1],true);var sel_align=Window_GetElement(window,OxO4df1[2],true);var sel_valign=Window_GetElement(window,OxO4df1[3],true);var inp_bgColor=Window_GetElement(window,OxO4df1[4],true);var inp_borderColor=Window_GetElement(window,OxO4df1[5],true);var inp_borderColorLight=Window_GetElement(window,OxO4df1[6],true);var inp_borderColorDark=Window_GetElement(window,OxO4df1[7],true);var inp_class=Window_GetElement(window,OxO4df1[8],true);var inp_id=Window_GetElement(window,OxO4df1[9],true);var inp_tooltip=Window_GetElement(window,OxO4df1[10],true);SyncToView=function SyncToView_Tr(){inp_bgColor[OxO4df1[11]]=element.getAttribute(OxO4df1[12])||element[OxO4df1[14]][OxO4df1[13]]||OxO4df1[15];inp_id[OxO4df1[11]]=element.getAttribute(OxO4df1[16])||OxO4df1[15];inp_bgColor[OxO4df1[14]][OxO4df1[13]]=inp_bgColor[OxO4df1[11]]||OxO4df1[15];inp_borderColor[OxO4df1[11]]=element.getAttribute(OxO4df1[17])||OxO4df1[15];inp_borderColor[OxO4df1[14]][OxO4df1[13]]=inp_borderColor[OxO4df1[11]]||OxO4df1[15];inp_borderColorLight[OxO4df1[11]]=element.getAttribute(OxO4df1[18])||OxO4df1[15];inp_borderColorLight[OxO4df1[14]][OxO4df1[13]]=inp_borderColorLight[OxO4df1[11]]||OxO4df1[15];inp_borderColorDark[OxO4df1[11]]=element.getAttribute(OxO4df1[19])||OxO4df1[15];inp_borderColorDark[OxO4df1[14]][OxO4df1[13]]=inp_borderColorDark[OxO4df1[11]]||OxO4df1[15];inp_class[OxO4df1[11]]=element[OxO4df1[20]];inp_width[OxO4df1[11]]=element.getAttribute(OxO4df1[21])||element[OxO4df1[14]][OxO4df1[21]]||OxO4df1[15];inp_height[OxO4df1[11]]=element.getAttribute(OxO4df1[22])||element[OxO4df1[14]][OxO4df1[22]]||OxO4df1[15];sel_align[OxO4df1[11]]=element.getAttribute(OxO4df1[23])||OxO4df1[15];sel_valign[OxO4df1[11]]=element.getAttribute(OxO4df1[24])||OxO4df1[15];inp_tooltip[OxO4df1[11]]=element.getAttribute(OxO4df1[25])||OxO4df1[15];} ;SyncTo=function SyncTo_Tr(element){if(inp_bgColor[OxO4df1[11]]){if(element[OxO4df1[14]][OxO4df1[13]]){element[OxO4df1[14]][OxO4df1[13]]=inp_bgColor[OxO4df1[11]];} else {element[OxO4df1[12]]=inp_bgColor[OxO4df1[11]];} ;} else {element.removeAttribute(OxO4df1[12]);} ;element[OxO4df1[17]]=inp_borderColor[OxO4df1[11]];element[OxO4df1[18]]=inp_borderColorLight[OxO4df1[11]];element[OxO4df1[19]]=inp_borderColorDark[OxO4df1[11]];element[OxO4df1[20]]=inp_class[OxO4df1[11]];if(element[OxO4df1[14]][OxO4df1[21]]||element[OxO4df1[14]][OxO4df1[22]]){try{element[OxO4df1[14]][OxO4df1[21]]=inp_width[OxO4df1[11]];element[OxO4df1[14]][OxO4df1[22]]=inp_height[OxO4df1[11]];} catch(er){alert(CE_GetStr(OxO4df1[26]));} ;} else {try{element[OxO4df1[21]]=inp_width[OxO4df1[11]];element[OxO4df1[22]]=inp_height[OxO4df1[11]];} catch(er){alert(CE_GetStr(OxO4df1[26]));} ;} ;var Ox5b0=/[^a-z\d]/i;if(Ox5b0.test(inp_id.value)){alert(CE_GetStr(OxO4df1[27]));return ;} ;element[OxO4df1[23]]=sel_align[OxO4df1[11]];element[OxO4df1[16]]=inp_id[OxO4df1[11]];element[OxO4df1[24]]=sel_valign[OxO4df1[11]];element[OxO4df1[25]]=inp_tooltip[OxO4df1[11]];if(element[OxO4df1[16]]==OxO4df1[15]){element.removeAttribute(OxO4df1[16]);} ;if(element[OxO4df1[12]]==OxO4df1[15]){element.removeAttribute(OxO4df1[12]);} ;if(element[OxO4df1[17]]==OxO4df1[15]){element.removeAttribute(OxO4df1[17]);} ;if(element[OxO4df1[18]]==OxO4df1[15]){element.removeAttribute(OxO4df1[18]);} ;if(element[OxO4df1[7]]==OxO4df1[15]){element.removeAttribute(OxO4df1[7]);} ;if(element[OxO4df1[20]]==OxO4df1[15]){element.removeAttribute(OxO4df1[20]);} ;if(element[OxO4df1[20]]==OxO4df1[15]){element.removeAttribute(OxO4df1[28]);} ;if(element[OxO4df1[23]]==OxO4df1[15]){element.removeAttribute(OxO4df1[23]);} ;if(element[OxO4df1[24]]==OxO4df1[15]){element.removeAttribute(OxO4df1[29]);} ;if(element[OxO4df1[25]]==OxO4df1[15]){element.removeAttribute(OxO4df1[25]);} ;if(element[OxO4df1[21]]==OxO4df1[15]){element.removeAttribute(OxO4df1[21]);} ;if(element[OxO4df1[22]]==OxO4df1[15]){element.removeAttribute(OxO4df1[22]);} ;} ;inp_borderColor[OxO4df1[30]]=function inp_borderColor_onclick(){SelectColor(inp_borderColor);} ;inp_bgColor[OxO4df1[30]]=function inp_bgColor_onclick(){SelectColor(inp_bgColor);} ;inp_borderColorLight[OxO4df1[30]]=function inp_borderColorLight_onclick(){SelectColor(inp_borderColorLight);} ;inp_borderColorDark[OxO4df1[30]]=function inp_borderColorDark_onclick(){SelectColor(inp_borderColorDark);} ;