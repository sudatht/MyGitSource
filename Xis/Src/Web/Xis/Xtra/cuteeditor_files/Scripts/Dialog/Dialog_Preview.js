var OxO694a=["formSearch","idSource","inc_width","inc_height","W640","W800","W1024","onload","availWidth","screen","window","availHeight","contentWindow","outerHTML","documentElement","text/html","replace","onresize","value","dialogWidth","dialogHeight","innerWidth","innerHeight","px","dialogTop","dialogLeft","screenY","screenX","checked","contentDocument","document"];var formSearch=Window_GetElement(window,OxO694a[0],true);var idSource=Window_GetElement(window,OxO694a[1],true);var inc_width=Window_GetElement(window,OxO694a[2],true);var inc_height=Window_GetElement(window,OxO694a[3],true);var W640=Window_GetElement(window,OxO694a[4],true);var W800=Window_GetElement(window,OxO694a[5],true);var W1024=Window_GetElement(window,OxO694a[6],true);var ParentW;var ParentH;window[OxO694a[7]]=function window_onload(){ParentW=top[OxO694a[10]][OxO694a[9]][OxO694a[8]];ParentH=top[OxO694a[10]][OxO694a[9]][OxO694a[11]];var iframe=idSource[OxO694a[12]];var editdoc=Window_GetDialogArguments(window);var Ox65f;if(Browser_IsWinIE()){Ox65f=editdoc[OxO694a[14]][OxO694a[13]];} else {Ox65f=outerHTML(editdoc.documentElement);} ;var Ox660=Frame_GetContentDocument(iframe);Ox660.open(OxO694a[15],OxO694a[16]);Ox660.write(Ox65f);Ox660.close();ShowSizeInfo();} ;window[OxO694a[17]]=ShowSizeInfo;function ShowSizeInfo(){if(Browser_IsWinIE()){inc_width[OxO694a[18]]=self[OxO694a[19]];inc_height[OxO694a[18]]=self[OxO694a[20]];} else {inc_width[OxO694a[18]]=self[OxO694a[21]];inc_height[OxO694a[18]]=self[OxO694a[22]];} ;} ;function do_Close(){Window_CloseDialog(window);} ;function ResizeThis(Ox59e,Ox7e){if(Browser_IsWinIE()){self[OxO694a[19]]=Ox59e+OxO694a[23];self[OxO694a[20]]=Ox7e+OxO694a[23];var Ox15d=ParentW/2-Ox59e/2;var Ox47a=ParentH/2-Ox7e/2;self[OxO694a[24]]=Ox47a+OxO694a[23];self[OxO694a[25]]=Ox15d+OxO694a[23];} else {if(Browser_IsGecko()){self[OxO694a[21]]=Ox59e;self[OxO694a[22]]=Ox7e;var Ox15d=ParentW/2-Ox59e/2;var Ox47a=ParentH/2-Ox7e/2;self[OxO694a[26]]=Ox47a;self[OxO694a[27]]=Ox15d;} else {window.resizeTo(Ox59e,Ox7e);if((screen[OxO694a[8]]-Ox59e>0)&&(screen[OxO694a[11]]-Ox7e>0)){window.moveTo((screen[OxO694a[8]]-Ox59e)/2,(screen[OxO694a[11]]-Ox7e)/2);} ;} ;} ;switch(Ox59e){case 640:W640[OxO694a[28]]=true;break ;;case 800:W800[OxO694a[28]]=true;break ;;case 1024:W1024[OxO694a[28]]=true;break ;;} ;} ;function Frame_GetContentDocument(Ox580){if(Ox580[OxO694a[29]]){return Ox580[OxO694a[29]];} ;return Ox580[OxO694a[30]];} ;