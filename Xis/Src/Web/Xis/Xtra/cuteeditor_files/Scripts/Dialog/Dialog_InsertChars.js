var OxOb541=["Verdana","innerHTML","Unicode","innerText","\x3Cspan style=\x27font-family:","\x27\x3E","\x3C/span\x3E","selfont","length","checked","value","charstable1","charstable2","fontFamily","style","display","block","none"];var editor=Window_GetDialogArguments(window);function getchar(obj){var Ox7e;var Ox7f=getFontValue()||OxOb541[0];if(!obj[OxOb541[1]]){return ;} ;Ox7e=obj[OxOb541[1]];if(Ox7f==OxOb541[2]){Ox7e=obj[OxOb541[3]];} else {if(Ox7f!=OxOb541[0]){Ox7e=OxOb541[4]+Ox7f+OxOb541[5]+obj[OxOb541[1]]+OxOb541[6];} ;} ;editor.PasteHTML(Ox7e);Window_CloseDialog(window);} ;function cancel(){Window_CloseDialog(window);} ;function getFontValue(){var Ox82=document.getElementsByName(OxOb541[7]);for(var i=0;i<Ox82[OxOb541[8]];i++){if(Ox82.item(i)[OxOb541[9]]){return Ox82.item(i)[OxOb541[10]];} ;} ;} ;function sel_font_change(){var Ox84=getFontValue()||OxOb541[0];var Ox5b7=Window_GetElement(window,OxOb541[11],true);var Ox5b8=Window_GetElement(window,OxOb541[12],true);Ox5b7[OxOb541[14]][OxOb541[13]]=Ox84;Ox5b7[OxOb541[14]][OxOb541[15]]=(Ox84!=OxOb541[2]?OxOb541[16]:OxOb541[17]);Ox5b8[OxOb541[14]][OxOb541[15]]=(Ox84==OxOb541[2]?OxOb541[16]:OxOb541[17]);} ;