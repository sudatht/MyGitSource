var OxO223d=["","removeNode","parentNode","firstChild","nodeName","TABLE","length","Can\x27t Get The Position ?","Map","RowCount","ColCount","rows","cells","Unknown Error , pos ",":"," already have cell","rowSpan","colSpan","Unknown Error , Unable to find bestpos","inp_cellspacing","inp_cellpadding","inp_id","inp_border","inp_bgcolor","inp_bordercolor","sel_rules","inp_collapse","inp_summary","btn_editcaption","btn_delcaption","btn_insthead","btn_instfoot","inp_class","inp_width","sel_width_unit","inp_height","sel_height_unit","sel_align","sel_textalign","sel_float","inp_tooltip","onclick","tHead","tFoot","caption","innerHTML","innerText","Unable to delete the caption. Please remove it in HTML source.","display","style","none","disabled","value","cellSpacing","cellPadding","id","border","borderColor","backgroundColor","bgColor","checked","borderCollapse","collapse","rules","summary","className","width","options","selectedIndex","height","align","styleFloat","cssFloat","textAlign","title","bordercolor","0","%","class","CaptionTable"];function ParseFloatToString(Ox78){var Ox57=parseFloat(Ox78);if(isNaN(Ox57)){return OxO223d[0];} ;return Ox57+OxO223d[0];} ;function Element_RemoveNode(element,Ox706){if(element[OxO223d[1]]){element.removeNode(Ox706);return ;} ;var p=element[OxO223d[2]];if(!p){return ;} ;if(Ox706){p.removeChild(element);return ;} ;while(true){var Ox29=element[OxO223d[3]];if(!Ox29){break ;} ;p.insertBefore(Ox29,element);} ;p.removeChild(element);} ;function Table_GetTable(Ox3e){for(;Ox3e!=null;Ox3e=Ox3e[OxO223d[2]]){if(Ox3e[OxO223d[4]]==OxO223d[5]){return Ox3e;} ;} ;return null;} ;function Table_GetCellPositionFromMap(Ox700,Ox9b){for(var y=0;y<Ox700[OxO223d[6]];y++){var Ox703=Ox700[y];for(var x=0;x<Ox703[OxO223d[6]];x++){if(Ox703[x]==Ox9b){return {rowIndex:y,cellIndex:x};} ;} ;} ;throw ( new Error(-1,OxO223d[7]));} ;function Table_GetCellMap(Ox99){return Table_CalculateTableInfo(Ox99)[OxO223d[8]];} ;function Table_GetRowCount(Ox3e){return Table_CalculateTableInfo(Ox3e)[OxO223d[9]];} ;function Table_GetColCount(Ox3e){return Table_CalculateTableInfo(Ox3e)[OxO223d[10]];} ;function Table_CalculateTableInfo(Ox3e){var Ox99=Table_GetTable(Ox3e);var Ox713=Ox99[OxO223d[11]];for(var Ox59=Ox713[OxO223d[6]]-1;Ox59>=0;Ox59--){var Ox1b=Ox713.item(Ox59);if(Ox1b[OxO223d[12]][OxO223d[6]]==0){Element_RemoveNode(Ox1b,true);continue ;} ;} ;var Ox714=Ox713[OxO223d[6]];var Ox715=0;var Ox716= new Array(Ox713.length);for(var Ox717=0;Ox717<Ox714;Ox717++){Ox716[Ox717]=[];} ;function Ox718(Ox59,Ox29,Ox9b){while(Ox59>=Ox714){Ox716[Ox714]=[];Ox714++;} ;var Ox719=Ox716[Ox59];if(Ox29>=Ox715){Ox715=Ox29+1;} ;if(Ox719[Ox29]!=null){throw ( new Error(-1,OxO223d[13]+Ox59+OxO223d[14]+Ox29+OxO223d[15]));} ;Ox719[Ox29]=Ox9b;} ;function Ox71a(Ox59,Ox29){var Ox719=Ox716[Ox59];if(Ox719){return Ox719[Ox29];} ;} ;for(var Ox717=0;Ox717<Ox713[OxO223d[6]];Ox717++){var Ox1b=Ox713.item(Ox717);var Ox71b=Ox1b[OxO223d[12]];for(var Oxa2=0;Oxa2<Ox71b[OxO223d[6]];Oxa2++){var Ox9b=Ox71b.item(Oxa2);var Ox71c=Ox9b[OxO223d[16]];var Ox71d=Ox9b[OxO223d[17]];var Ox719=Ox716[Ox717];var Ox71e=-1;for(var Ox71f=0;Ox71f<Ox715+Ox71d+1;Ox71f++){if(Ox71c==1&&Ox71d==1){if(Ox719[Ox71f]==null){Ox71e=Ox71f;break ;} ;} else {var Ox720=true;for(var Ox721=0;Ox721<Ox71c;Ox721++){for(var Ox722=0;Ox722<Ox71d;Ox722++){if(Ox71a(Ox717+Ox721,Ox71f+Ox722)!=null){Ox720=false;break ;} ;} ;} ;if(Ox720){Ox71e=Ox71f;break ;} ;} ;} ;if(Ox71e==-1){throw ( new Error(-1,OxO223d[18]));} ;if(Ox71c==1&&Ox71d==1){Ox718(Ox717,Ox71e,Ox9b);} else {for(var Ox721=0;Ox721<Ox71c;Ox721++){for(var Ox722=0;Ox722<Ox71d;Ox722++){Ox718(Ox717+Ox721,Ox71e+Ox722,Ox9b);} ;} ;} ;} ;} ;var Ox37c={};Ox37c[OxO223d[9]]=Ox714;Ox37c[OxO223d[10]]=Ox715;Ox37c[OxO223d[8]]=Ox716;for(var Ox59=0;Ox59<Ox714;Ox59++){var Ox719=Ox716[Ox59];for(var Ox29=0;Ox29<Ox715;Ox29++){} ;} ;return Ox37c;} ;var inp_cellspacing=Window_GetElement(window,OxO223d[19],true);var inp_cellpadding=Window_GetElement(window,OxO223d[20],true);var inp_id=Window_GetElement(window,OxO223d[21],true);var inp_border=Window_GetElement(window,OxO223d[22],true);var inp_bgcolor=Window_GetElement(window,OxO223d[23],true);var inp_bordercolor=Window_GetElement(window,OxO223d[24],true);var sel_rules=Window_GetElement(window,OxO223d[25],true);var inp_collapse=Window_GetElement(window,OxO223d[26],true);var inp_summary=Window_GetElement(window,OxO223d[27],true);var btn_editcaption=Window_GetElement(window,OxO223d[28],true);var btn_delcaption=Window_GetElement(window,OxO223d[29],true);var btn_insthead=Window_GetElement(window,OxO223d[30],true);var btn_instfoot=Window_GetElement(window,OxO223d[31],true);var inp_class=Window_GetElement(window,OxO223d[32],true);var inp_width=Window_GetElement(window,OxO223d[33],true);var sel_width_unit=Window_GetElement(window,OxO223d[34],true);var inp_height=Window_GetElement(window,OxO223d[35],true);var sel_height_unit=Window_GetElement(window,OxO223d[36],true);var sel_align=Window_GetElement(window,OxO223d[37],true);var sel_textalign=Window_GetElement(window,OxO223d[38],true);var sel_float=Window_GetElement(window,OxO223d[39],true);var inp_tooltip=Window_GetElement(window,OxO223d[40],true);function insertOneRow(Ox812,Ox604){var Ox1b=Ox812.insertRow(-1);for(var i=0;i<Ox604;i++){Ox1b.insertCell();} ;} ;btn_insthead[OxO223d[41]]=function btn_insthead_onclick(){if(element[OxO223d[42]]){element.deleteTHead();} else {var Ox814=Table_GetColCount(element);var Ox815=element.createTHead();insertOneRow(Ox815,Ox814);} ;} ;btn_instfoot[OxO223d[41]]=function btn_instfoot_onclick(){if(element[OxO223d[43]]){element.deleteTFoot();} else {var Ox814=Table_GetColCount(element);var Ox817=element.createTFoot();insertOneRow(Ox817,Ox814);} ;} ;btn_editcaption[OxO223d[41]]=function btn_editcaption_onclick(){var Ox819=element[OxO223d[44]];if(Ox819!=null){var Ox4bc=editor.EditInWindow(Ox819.innerHTML,window);if(Ox4bc!=null&&Ox4bc!==false){Ox819[OxO223d[45]]=Ox4bc;} ;} else {var Ox819=element.createCaption();if(Browser_IsGecko()){Ox819[OxO223d[45]]=Caption;} else {Ox819[OxO223d[46]]=Caption;} ;} ;} ;btn_delcaption[OxO223d[41]]=function btn_delcaption_onclick(){if(element[OxO223d[44]]!=null){alert(OxO223d[47]);} ;} ;UpdateState=function UpdateState_Table(){if(Browser_IsGecko()){btn_insthead[OxO223d[45]]=element[OxO223d[42]]?Delete:Insert;btn_instfoot[OxO223d[45]]=element[OxO223d[43]]?Delete:Insert;} else {btn_insthead[OxO223d[46]]=element[OxO223d[42]]?Delete:Insert;btn_instfoot[OxO223d[46]]=element[OxO223d[43]]?Delete:Insert;} ;if(element[OxO223d[44]]!=null){if(Browser_IsGecko()){btn_editcaption[OxO223d[45]]=Edit;} else {btn_editcaption[OxO223d[46]]=Edit;} ;btn_editcaption[OxO223d[49]][OxO223d[48]]=OxO223d[50];btn_delcaption[OxO223d[51]]=false;} else {if(Browser_IsGecko()){btn_editcaption[OxO223d[45]]=Insert;} else {btn_editcaption[OxO223d[46]]=Insert;} ;btn_delcaption[OxO223d[51]]=true;} ;} ;var t_inp_width=OxO223d[0];var t_inp_height=OxO223d[0];SyncToView=function SyncToView_Table(){inp_cellspacing[OxO223d[52]]=element.getAttribute(OxO223d[53])||OxO223d[0];inp_cellpadding[OxO223d[52]]=element.getAttribute(OxO223d[54])||OxO223d[0];inp_id[OxO223d[52]]=element.getAttribute(OxO223d[55])||OxO223d[0];inp_border[OxO223d[52]]=element.getAttribute(OxO223d[56])||OxO223d[0];inp_bordercolor[OxO223d[52]]=element.getAttribute(OxO223d[57])||OxO223d[0];inp_bordercolor[OxO223d[49]][OxO223d[58]]=inp_bordercolor[OxO223d[52]]||OxO223d[0];inp_bgcolor[OxO223d[52]]=element.getAttribute(OxO223d[59])||element[OxO223d[49]][OxO223d[58]]||OxO223d[0];inp_bgcolor[OxO223d[49]][OxO223d[58]]=inp_bgcolor[OxO223d[52]]||OxO223d[0];inp_collapse[OxO223d[60]]=element[OxO223d[49]][OxO223d[61]]==OxO223d[62];sel_rules[OxO223d[52]]=element.getAttribute(OxO223d[63])||OxO223d[0];inp_summary[OxO223d[52]]=element.getAttribute(OxO223d[64])||OxO223d[0];inp_class[OxO223d[52]]=element[OxO223d[65]];if(element.getAttribute(OxO223d[66])){t_inp_width=element.getAttribute(OxO223d[66]);} else {if(element[OxO223d[49]][OxO223d[66]]){t_inp_width=element[OxO223d[49]][OxO223d[66]];} ;} ;if(t_inp_width){inp_width[OxO223d[52]]=ParseFloatToString(t_inp_width);for(var i=0;i<sel_width_unit[OxO223d[67]][OxO223d[6]];i++){var Oxd9=sel_width_unit[OxO223d[67]][i][OxO223d[52]];if(Oxd9&&t_inp_width.indexOf(Oxd9)!=-1){sel_width_unit[OxO223d[68]]=i;break ;} ;} ;} ;if(element.getAttribute(OxO223d[69])){t_inp_height=element.getAttribute(OxO223d[69]);} else {if(element[OxO223d[49]][OxO223d[69]]){t_inp_height=element[OxO223d[49]][OxO223d[69]];} ;} ;if(t_inp_height){inp_height[OxO223d[52]]=ParseFloatToString(t_inp_height);for(var i=0;i<sel_height_unit[OxO223d[67]][OxO223d[6]];i++){var Oxd9=sel_height_unit[OxO223d[67]][i][OxO223d[52]];if(Oxd9&&t_inp_height.indexOf(Oxd9)!=-1){sel_height_unit[OxO223d[68]]=i;break ;} ;} ;} ;sel_align[OxO223d[52]]=element[OxO223d[70]];if(Browser_IsWinIE()){sel_float[OxO223d[52]]=element[OxO223d[49]][OxO223d[71]];} else {sel_float[OxO223d[52]]=element[OxO223d[49]][OxO223d[72]];} ;sel_textalign[OxO223d[52]]=element[OxO223d[49]][OxO223d[73]];inp_tooltip[OxO223d[52]]=element[OxO223d[74]];} ;SyncTo=function SyncTo_Table(element){if(Browser_IsWinIE()){element[OxO223d[57]]=inp_bordercolor[OxO223d[52]];} else {element.setAttribute(OxO223d[75],inp_bordercolor.value);} ;if(inp_bgcolor[OxO223d[52]]){if(element[OxO223d[49]][OxO223d[58]]){element[OxO223d[49]][OxO223d[58]]=inp_bgcolor[OxO223d[52]];} else {element[OxO223d[59]]=inp_bgcolor[OxO223d[52]];} ;} else {element.removeAttribute(OxO223d[59]);} ;element[OxO223d[49]][OxO223d[61]]=inp_collapse[OxO223d[60]]?OxO223d[62]:OxO223d[0];element[OxO223d[63]]=sel_rules[OxO223d[52]]||OxO223d[0];element[OxO223d[64]]=inp_summary[OxO223d[52]];element[OxO223d[65]]=inp_class[OxO223d[52]];element[OxO223d[53]]=inp_cellspacing[OxO223d[52]];element[OxO223d[54]]=inp_cellpadding[OxO223d[52]];var Ox5b0=/[^a-z\d]/i;if(Ox5b0.test(inp_id.value)){alert(ValidID);return ;} ;element[OxO223d[55]]=inp_id[OxO223d[52]];if(inp_border[OxO223d[52]]==OxO223d[0]){element[OxO223d[56]]=OxO223d[76];} else {element[OxO223d[56]]=inp_border[OxO223d[52]];} ;if(inp_width[OxO223d[52]]==OxO223d[0]){element.removeAttribute(OxO223d[66]);element[OxO223d[49]][OxO223d[66]]=OxO223d[0];} else {try{t_inp_width=inp_width[OxO223d[52]];if(sel_width_unit[OxO223d[52]]==OxO223d[77]){t_inp_width=inp_width[OxO223d[52]]+sel_width_unit[OxO223d[52]];} ;if(element[OxO223d[49]][OxO223d[66]]&&element[OxO223d[66]]){element[OxO223d[49]][OxO223d[66]]=t_inp_width;element[OxO223d[66]]=t_inp_width;} else {if(element[OxO223d[49]][OxO223d[66]]){element[OxO223d[49]][OxO223d[66]]=t_inp_width;} else {element[OxO223d[66]]=t_inp_width;} ;} ;} catch(x){} ;} ;if(inp_height[OxO223d[52]]==OxO223d[0]){element.removeAttribute(OxO223d[69]);element[OxO223d[49]][OxO223d[69]]=OxO223d[0];} else {try{t_inp_height=inp_height[OxO223d[52]];if(sel_height_unit[OxO223d[52]]==OxO223d[77]){t_inp_height=inp_height[OxO223d[52]]+sel_height_unit[OxO223d[52]];} ;t_inp_height=inp_height[OxO223d[52]]+sel_height_unit[OxO223d[52]];if(element[OxO223d[49]][OxO223d[69]]&&element[OxO223d[69]]){element[OxO223d[49]][OxO223d[69]]=t_inp_height;element[OxO223d[69]]=t_inp_height;} else {if(element[OxO223d[49]][OxO223d[69]]){element[OxO223d[49]][OxO223d[69]]=t_inp_height;} else {element[OxO223d[69]]=t_inp_height;} ;} ;} catch(x){} ;} ;element[OxO223d[70]]=sel_align[OxO223d[52]];if(Browser_IsWinIE()){element[OxO223d[49]][OxO223d[71]]=sel_float[OxO223d[52]];} else {element[OxO223d[49]][OxO223d[72]]=sel_float[OxO223d[52]];} ;element[OxO223d[49]][OxO223d[73]]=sel_textalign[OxO223d[52]];element[OxO223d[74]]=inp_tooltip[OxO223d[52]];if(element[OxO223d[74]]==OxO223d[0]){element.removeAttribute(OxO223d[74]);} ;if(element[OxO223d[64]]==OxO223d[0]){element.removeAttribute(OxO223d[64]);} ;if(element[OxO223d[65]]==OxO223d[0]){element.removeAttribute(OxO223d[65]);} ;if(element[OxO223d[65]]==OxO223d[0]){element.removeAttribute(OxO223d[78]);} ;if(element[OxO223d[55]]==OxO223d[0]){element.removeAttribute(OxO223d[55]);} ;if(element[OxO223d[70]]==OxO223d[0]){element.removeAttribute(OxO223d[70]);} ;if(element[OxO223d[63]]==OxO223d[0]){element.removeAttribute(OxO223d[63]);} ;} ;inp_bgcolor[OxO223d[41]]=function inp_bgcolor_onclick(){SelectColor(inp_bgcolor);} ;inp_bordercolor[OxO223d[41]]=function inp_bordercolor_onclick(){SelectColor(inp_bordercolor);} ;if(!Browser_IsWinIE()){Window_GetElement(window,OxO223d[79],true)[OxO223d[49]][OxO223d[48]]=OxO223d[50];} ;