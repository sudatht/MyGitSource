var OxO6df3=["inp_src","box1","box2","box3","box4","box5","box6","box7","box8","box9","inp_start","CustomBullet","nodeName","LI","parentNode","none","decimal","upper-roman","upper-alpha","lower-alpha","lower-roman","disc","circle","square","listStyleType","style","border","solid 2px #708090","listStyleImage","","value","visibility","hidden","length","start","url(\x27","\x27)","visible","UL","OL","document","firstChild","element","solid 2px #ffffff","solid 2px #ECECF6","onclick"];var inp_src=Window_GetElement(window,OxO6df3[0],true);var box1=Window_GetElement(window,OxO6df3[1],true);var box2=Window_GetElement(window,OxO6df3[2],true);var box3=Window_GetElement(window,OxO6df3[3],true);var box4=Window_GetElement(window,OxO6df3[4],true);var box5=Window_GetElement(window,OxO6df3[5],true);var box6=Window_GetElement(window,OxO6df3[6],true);var box7=Window_GetElement(window,OxO6df3[7],true);var box8=Window_GetElement(window,OxO6df3[8],true);var box9=Window_GetElement(window,OxO6df3[9],true);var inp_start=Window_GetElement(window,OxO6df3[10],true);var CustomBullet=Window_GetElement(window,OxO6df3[11],true);OriginalnodeName=element[OxO6df3[12]];if(element[OxO6df3[12]]&&element[OxO6df3[12]]==OxO6df3[13]){OriginalnodeName=(element[OxO6df3[14]])[OxO6df3[12]];} ;var OriginalnodeName,CurrentnodeName,selectedObject;SyncToView=function SyncToView_LI(){if(element[OxO6df3[12]]==OxO6df3[13]){element=element[OxO6df3[14]];} ;switch((element[OxO6df3[25]][OxO6df3[24]]).toLowerCase()){case OxO6df3[15]:selectedObject=box1;break ;;case OxO6df3[16]:selectedObject=box2;break ;;case OxO6df3[17]:selectedObject=box3;break ;;case OxO6df3[18]:selectedObject=box4;break ;;case OxO6df3[19]:selectedObject=box5;break ;;case OxO6df3[20]:selectedObject=box6;break ;;case OxO6df3[21]:selectedObject=box7;break ;;case OxO6df3[22]:selectedObject=box8;break ;;case OxO6df3[23]:selectedObject=box9;break ;;default:selectedObject=box1;break ;;} ;selectedObject[OxO6df3[25]][OxO6df3[26]]=OxO6df3[27];if(element[OxO6df3[25]][OxO6df3[28]]==OxO6df3[29]){inp_src[OxO6df3[30]]=OxO6df3[29];CustomBullet[OxO6df3[25]][OxO6df3[31]]=OxO6df3[32];} else {var Ox398;Ox398=element[OxO6df3[25]][OxO6df3[28]];Ox398=Ox398.substring(4,Ox398[OxO6df3[33]]-1);inp_src[OxO6df3[30]]=Ox398;} ;} ;SyncTo=function SyncTo_LI(element){switch(selectedObject){case box1:;case box2:;case box3:;case box4:;case box5:;case box6:try{element.setAttribute(OxO6df3[34],inp_start.value);} catch(er){} ;break ;;case box7:;case box8:;case box9:break ;;} ;if(inp_src[OxO6df3[30]]){element[OxO6df3[25]][OxO6df3[28]]=OxO6df3[35]+inp_src[OxO6df3[30]]+OxO6df3[36];} ;} ;function ToggleCustomBullet(){if(CustomBullet[OxO6df3[25]][OxO6df3[31]]==OxO6df3[32]){CustomBullet[OxO6df3[25]][OxO6df3[31]]=OxO6df3[37];} else {CustomBullet[OxO6df3[25]][OxO6df3[31]]=OxO6df3[32];} ;} ;function doClick1(Ox5a9){if(element[OxO6df3[12]]==OxO6df3[13]){element=element[OxO6df3[14]];} ;selectedObject=Ox5a9;switch(selectedObject){case box1:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[15];break ;;case box2:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[16];break ;;case box3:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[17];break ;;case box4:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[18];break ;;case box5:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[19];break ;;case box6:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[20];break ;;case box7:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[21];break ;;case box8:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[22];break ;;case box9:element[OxO6df3[25]][OxO6df3[24]]=OxO6df3[23];break ;;} ;var Ox61a=false;switch(selectedObject){case box1:;case box2:;case box3:;case box4:;case box5:;case box6:if(OriginalnodeName==OxO6df3[38]){OriginalnodeName=OxO6df3[39];Ox61a=true;} ;break ;;case box7:;case box8:;case box9:if(OriginalnodeName==OxO6df3[39]){OriginalnodeName=OxO6df3[38];Ox61a=true;} ;break ;;} ;if(Ox61a){var Ox77a=editwin[OxO6df3[40]].createElement(OriginalnodeName);while(element[OxO6df3[41]]){Ox77a.appendChild(element.firstChild);} ;element[OxO6df3[14]].insertBefore(Ox77a,element);element[OxO6df3[14]].removeChild(element);var arg=Window_FindDialogArguments(window);arg[OxO6df3[42]]=element=Ox77a;} ;box1[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box2[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box3[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box4[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box5[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box6[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box7[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box8[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];box9[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];selectedObject[OxO6df3[25]][OxO6df3[26]]=OxO6df3[27];inp_src[OxO6df3[30]]=OxO6df3[29];SyncTo();} ;function doMouseOut(Ox5a9){if(Ox5a9==selectedObject){Ox5a9[OxO6df3[25]][OxO6df3[26]]=OxO6df3[27];} else {Ox5a9[OxO6df3[25]][OxO6df3[26]]=OxO6df3[43];} ;} ;function doMouseOver(Ox5a9){Ox5a9[OxO6df3[25]][OxO6df3[26]]=OxO6df3[44];} ;btnbrowse[OxO6df3[45]]=function btnbrowse_onclick(){function Ox595(Ox37c){if(Ox37c){inp_src[OxO6df3[30]]=Ox37c;SyncTo(element);} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectImageDialog(Ox595,inp_src.value,inp_src);} else {editor.ShowSelectImageDialog(Ox595,inp_src.value);} ;} ;