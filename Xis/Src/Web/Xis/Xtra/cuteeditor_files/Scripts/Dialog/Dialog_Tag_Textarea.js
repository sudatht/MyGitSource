var OxOf339=["inp_name","inp_cols","inp_rows","inp_value","sel_Wrap","inp_id","inp_access","inp_index","inp_Disabled","inp_Readonly","Name","value","name","id","cols","","rows","checked","disabled","readOnly","wrap","tabIndex","accessKey","textContent"];var inp_name=Window_GetElement(window,OxOf339[0],true);var inp_cols=Window_GetElement(window,OxOf339[1],true);var inp_rows=Window_GetElement(window,OxOf339[2],true);var inp_value=Window_GetElement(window,OxOf339[3],true);var sel_Wrap=Window_GetElement(window,OxOf339[4],true);var inp_id=Window_GetElement(window,OxOf339[5],true);var inp_access=Window_GetElement(window,OxOf339[6],true);var inp_index=Window_GetElement(window,OxOf339[7],true);var inp_Disabled=Window_GetElement(window,OxOf339[8],true);var inp_Readonly=Window_GetElement(window,OxOf339[9],true);UpdateState=function UpdateState_Textarea(){} ;SyncToView=function SyncToView_Textarea(){if(element[OxOf339[10]]){inp_name[OxOf339[11]]=element[OxOf339[10]];} ;if(element[OxOf339[12]]){inp_name[OxOf339[11]]=element[OxOf339[12]];} ;inp_id[OxOf339[11]]=element[OxOf339[13]];inp_value[OxOf339[11]]=element[OxOf339[11]];if(element[OxOf339[14]]){if(element[OxOf339[14]]==20){inp_cols[OxOf339[11]]=OxOf339[15];} else {inp_cols[OxOf339[11]]=element[OxOf339[14]];} ;} ;if(element[OxOf339[16]]){if(element[OxOf339[16]]==2){inp_rows[OxOf339[11]]=OxOf339[15];} else {inp_rows[OxOf339[11]]=element[OxOf339[16]];} ;} ;inp_Disabled[OxOf339[17]]=element[OxOf339[18]];inp_Readonly[OxOf339[17]]=element[OxOf339[19]];sel_Wrap[OxOf339[11]]=element[OxOf339[20]];if(element[OxOf339[21]]==0){inp_index[OxOf339[11]]=OxOf339[15];} else {inp_index[OxOf339[11]]=element[OxOf339[21]];} ;if(element[OxOf339[22]]){inp_access[OxOf339[11]]=element[OxOf339[22]];} ;} ;SyncTo=function SyncTo_Textarea(element){element[OxOf339[12]]=inp_name[OxOf339[11]];if(element[OxOf339[10]]){element[OxOf339[10]]=inp_name[OxOf339[11]];} else {if(element[OxOf339[12]]){element.removeAttribute(OxOf339[12],0);element[OxOf339[10]]=inp_name[OxOf339[11]];} else {element[OxOf339[10]]=inp_name[OxOf339[11]];} ;} ;element[OxOf339[13]]=inp_id[OxOf339[11]];element[OxOf339[11]]=inp_value[OxOf339[11]];if(!Browser_IsWinIE()){try{element[OxOf339[23]]=inp_value[OxOf339[11]];} catch(x){} ;} ;element[OxOf339[21]]=inp_index[OxOf339[11]];element[OxOf339[18]]=inp_Disabled[OxOf339[17]];element[OxOf339[19]]=inp_Readonly[OxOf339[17]];element[OxOf339[22]]=inp_access[OxOf339[11]];if(inp_cols[OxOf339[11]]==OxOf339[15]){element[OxOf339[14]]=20;} else {element[OxOf339[14]]=inp_cols[OxOf339[11]];} ;if(inp_rows[OxOf339[11]]==OxOf339[15]){element[OxOf339[16]]=2;} else {element[OxOf339[16]]=inp_rows[OxOf339[11]];} ;try{element[OxOf339[20]]=sel_Wrap[OxOf339[11]];} catch(e){element.removeAttribute(OxOf339[20]);} ;element[OxOf339[21]]=inp_index[OxOf339[11]];if(element[OxOf339[21]]==OxOf339[15]){element.removeAttribute(OxOf339[21]);} ;if(element[OxOf339[22]]==OxOf339[15]){element.removeAttribute(OxOf339[22]);} ;} ;