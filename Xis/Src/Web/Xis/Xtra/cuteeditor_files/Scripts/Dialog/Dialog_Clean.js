var OxOb456=["ig","\x3C/?[^\x3E]*\x3E","","allhtml","\x3C\x5C?xml[^\x3E]*\x3E","\x3C/?[a-z]+:[^\x3E]*\x3E","(\x3C[^\x3E]+) class=[^ |^\x3E]*([^\x3E]*\x3E)","$1 $2","(\x3C[^\x3E]+) style=\x22[^\x22]*\x22([^\x3E]*\x3E)","\x3Cspan[^\x3E]*\x3E\x3C/span[^\x3E]*\x3E","\x3Cspan\x3E\x3Cspan\x3E","\x3Cspan\x3E","\x3C/span\x3E\x3C/span\x3E","\x3C/span\x3E","[ ]*\x3E","\x3E","word","css","\x3C/?font[^\x3E]*\x3E","font","\x3C/?span[^\x3E]*\x3E","span"];var editor=Window_GetDialogArguments(window);function execRE(Ox1b6,Ox4d1,Ox4d2){var Ox4d3= new RegExp(Ox1b6,OxOb456[0]);return Ox4d2.replace(Ox4d3,Ox4d1);} ;function getContent(){return editor.GetBodyInnerHTML();} ;function setContent(Ox4d2){editor.SetHTML(Ox4d2);} ;function codeCleaner(Ox14c){var Ox4d2=getContent();switch(Ox14c){case OxOb456[3]:Ox4d2=execRE(OxOb456[1],OxOb456[2],Ox4d2);break ;;case OxOb456[16]:Ox4d2=execRE(OxOb456[4],OxOb456[2],Ox4d2);Ox4d2=execRE(OxOb456[5],OxOb456[2],Ox4d2);Ox4d2=execRE(OxOb456[6],OxOb456[7],Ox4d2);Ox4d2=execRE(OxOb456[8],OxOb456[7],Ox4d2);Ox4d2=execRE(OxOb456[9],OxOb456[2],Ox4d2);Ox4d2=execRE(OxOb456[10],OxOb456[11],Ox4d2);Ox4d2=execRE(OxOb456[12],OxOb456[13],Ox4d2);Ox4d2=execRE(OxOb456[14],OxOb456[15],Ox4d2);break ;;case OxOb456[17]:Ox4d2=execRE(OxOb456[6],OxOb456[7],Ox4d2);Ox4d2=execRE(OxOb456[8],OxOb456[7],Ox4d2);break ;;case OxOb456[19]:Ox4d2=execRE(OxOb456[18],OxOb456[2],Ox4d2);break ;;case OxOb456[21]:Ox4d2=execRE(OxOb456[20],OxOb456[2],Ox4d2);break ;;} ;setContent(Ox4d2);} ;