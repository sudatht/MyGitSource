var OxOd0a9=["inp_action","sel_Method","inp_name","inp_id","inp_encode","sel_target","Name","value","name","id","action","method","encoding","application/x-www-form-urlencoded","","target"];var inp_action=Window_GetElement(window,OxOd0a9[0],true);var sel_Method=Window_GetElement(window,OxOd0a9[1],true);var inp_name=Window_GetElement(window,OxOd0a9[2],true);var inp_id=Window_GetElement(window,OxOd0a9[3],true);var inp_encode=Window_GetElement(window,OxOd0a9[4],true);var sel_target=Window_GetElement(window,OxOd0a9[5],true);UpdateState=function UpdateState_Form(){} ;SyncToView=function SyncToView_Form(){if(element[OxOd0a9[6]]){inp_name[OxOd0a9[7]]=element[OxOd0a9[6]];} ;if(element[OxOd0a9[8]]){inp_name[OxOd0a9[7]]=element[OxOd0a9[8]];} ;inp_id[OxOd0a9[7]]=element[OxOd0a9[9]];if(element[OxOd0a9[10]]){inp_action[OxOd0a9[7]]=element[OxOd0a9[10]];} ;if(element[OxOd0a9[11]]){sel_Method[OxOd0a9[7]]=element[OxOd0a9[11]];} ;if(element[OxOd0a9[12]]==OxOd0a9[13]){inp_encode[OxOd0a9[7]]=OxOd0a9[14];} else {inp_encode[OxOd0a9[7]]=element[OxOd0a9[12]];} ;if(element[OxOd0a9[15]]){sel_target[OxOd0a9[7]]=element[OxOd0a9[15]];} ;} ;SyncTo=function SyncTo_Form(element){element[OxOd0a9[8]]=inp_name[OxOd0a9[7]];if(element[OxOd0a9[6]]){element[OxOd0a9[6]]=inp_name[OxOd0a9[7]];} else {if(element[OxOd0a9[8]]){element.removeAttribute(OxOd0a9[8],0);element[OxOd0a9[6]]=inp_name[OxOd0a9[7]];} else {element[OxOd0a9[6]]=inp_name[OxOd0a9[7]];} ;} ;element[OxOd0a9[9]]=inp_id[OxOd0a9[7]];element[OxOd0a9[10]]=inp_action[OxOd0a9[7]];element[OxOd0a9[11]]=sel_Method[OxOd0a9[7]];try{element[OxOd0a9[12]]=inp_encode[OxOd0a9[7]];} catch(e){} ;element[OxOd0a9[15]]=sel_target[OxOd0a9[7]];if(element[OxOd0a9[15]]==OxOd0a9[14]){element.removeAttribute(OxOd0a9[15]);} ;if(element[OxOd0a9[6]]==OxOd0a9[14]){element.removeAttribute(OxOd0a9[6]);} ;if(element[OxOd0a9[10]]==OxOd0a9[14]){element.removeAttribute(OxOd0a9[10]);} ;} ;