/*######################## finne formnr #############################################################*/

function finnFormNr(formNavn){
	formene = document.forms;
	for (i=0; i<=formene.length; i++){
		//alert(document.forms[i].name);
		if(formene[i].name == formNavn){
			//alert("Formen " + formNavn + " er funnet og har nummer " + i + " !");
			break;
		}
	}
	return i;
}

//#######################################################################################################
/*########################test av datoformat#############################################################*/
//#######################################################################################################
var boltest=false;
var boltest2=false;
var aktivtFelt;
var formNr0 = 0;

function dateCheck(formNr0, felt)
{
	aktivtFelt=felt;

	formNr0 = formNr0.name;

	var date;
	//alert("felt: " + felt + ", formNr: " + formNr0);
	//alert(document.forms["formEn"].elements["tbxFradato"].value)
	date = document.forms[formNr0].elements[felt].value; //string for å lese tiden fra tekstfeltet
	var ok=false;

	switch (date.length){

		case 0:
					document.forms[formNr0].elements[felt].value=date;		
		break;
	//---------------------------------------------------------------------
		case 6:
				
					switch (date.indexOf(".")){
						case 0:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 1:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 2:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 3:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 4:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 5:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 6:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						
					default:				
						ok=true;	
					}
		break;
	//----------------------------------------------------------------------
		case 8:

					switch (date.indexOf(".")){
						case 0:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;
						case 1:
						alert(date+" er feil format");
						document.forms[formNr0].elements[felt].focus();
						break;

					default:				
					var hjelpevar=date.substring(3,8);			
					switch (date.indexOf(".")&& hjelpevar.indexOf(".")){						
							case 2:						
									var stripPkt = date.substring(0,2)+date.substring(3,5)+date.substring(6,8);
									date = stripPkt;	
									ok=true;							
							break;
					default:
									alert(date+" er feil format \n\n bruk DD.MM.ÅÅ");
									document.forms[formNr0].elements[felt].focus();
					}}
		break;
	//------------------------------------------------------------------------
		default:
					alert(date+" er feil datoformat \n\n bruk DD.MM.ÅÅ");
					//document.forms[formNr0].elements[felt].value="";		
					document.forms[formNr0].elements[felt].select();	
	}
	//-------------------------------------------------------------------------
	switch (ok){
		case true:
			var DD = date.substring(0,2);
			var MM = date.substring(2,4);
			var AA = date.substring(4,6);

			
			
			if(isNaN(DD)||isNaN(MM)||isNaN(AA)){
				alert(DD+" eller "+MM+ " eller "+AA+"\n er ikke et tall!");
				document.forms[formNr0].elements[felt].focus();
				
			}
			else if (DD>31||MM>12||(DD>30&&(MM==4||MM==6||MM==9||MM==11))||(DD>30&&MM==2)){
			alert("Datoen er ugyldig. \nDag eller måned er et for stort tall. \nTast inn en ny dato på formatet ddmmåå.");
			document.forms[formNr0].elements[felt].focus();				
			}
			
			else 
			{  
				document.forms[formNr0].elements[felt].value=DD+"."+MM+"."+AA;
				if (felt=='tbxFraDato')
				{
				boltest=true
				}
				else
				boltest2=true;
			}	
		break;
	}
}
//#########################################################################################

function dateCheck2(formNavn,elementNavn){
	formNr0 = finnFormNr(formNavn);
//alert(formNr0);
	dateCheck(elementNavn);
}

//#########################################################################################
/*##############################test for riktig klokkeslett################################*/
//#########################################################################################

var bol;
bol=false;
var tt;
var mm;
var formNr1 = 0;

function timeCheck(formNr1, felt){

var test;
formNr1 = formNr1.name;



test=document.forms[formNr1].elements[felt].value; //string for å lese tiden fra tekstfeltet
var ok=false;
//alert(test);
switch (test.length){

	case 0:
				document.forms[formNr1].elements[felt].value=test;		
	break;
//---------------------------------------------------------------------
	case 4:
				
				
				ok=true;	
	break;
//----------------------------------------------------------------------
	case 5:
				
				switch (test.indexOf(":")){
						case 2:						
								var stripPkt = test.substring(0,2)+test.substring(3,5);
								test = stripPkt;	
								ok=true;							
						break;
				default:
								alert(test+" er feil format \n\n bruk TT:MM");
								document.forms[formNr1].elements[felt].focus();
								bol=false;
				}
	break;
//------------------------------------------------------------------------
	default:
				alert(test+" er feil tidsformat \n\n bruk TT:MM");	
				document.forms[formNr1].elements[felt].focus();	
				bol=false;
}
//-------------------------------------------------------------------------
switch (ok){
	case true:
		var TT = test.substring(0,2);
		var MM = test.substring(2,4);
		
		
		if(isNaN(MM)||isNaN(TT)){
			alert(TT+" eller "+MM+ " er ikke et tall");
			document.forms[formNr1].elements[felt].focus();
			bol=false;
			
		}
		else if (MM>59||TT>24){
		alert("dette klokkeslettet finnes ikke! \n\n bruk TT:MM");
		document.forms[formNr1].elements[felt].focus();
		bol=false;				
		}
		else if (TT==24){
			TT="00";
			document.forms[formNr1].elements[felt].value=TT+":"+MM;	
			bol=true;
		}
		else {  document.forms[formNr1].elements[felt].value=TT+":"+MM;
			tt=TT;
			mm=MM;
			bol=true;
		}
	break;
}}
//####################################################################################################

function timeCheck2(formNavn,elementNavn){
	formNr1 = finnFormNr(formNavn);
//alert(formNr1);
	timeCheck(elementNavn);
}

//###################################################################################################
//*******************************utregning og maskering av arbeidstid*******************************
//###################################################################################################

var formNr2 = 0;

function workTime(formNr2,felt){
formNr2 = formNr2.name;
//alert(felt + " " + document.forms[formNr2].name + " " + formNr2);
var strTil;
var strFra;
var TTfra;
var MMfra ;
var TTtil;
var MMtil;
var arbeidstid;
var TTlunsj;
var MMlunsj;
var lunsjTid;


if (bol==true){
switch (felt){

	case 'tbxFraKl':
		strLunsj=document.forms[formNr2].elements['tbxLunsj'].value;			
		TTlunsj = strLunsj.substring(0,2);
		MMlunsj = strLunsj.substring(3,5)/60
		var lunsjTid = parseFloat(TTlunsj) + parseFloat(MMlunsj);	
		
		strTil=document.forms[formNr2].elements['tbxTilKl'].value;
		TTtil = strTil.substring(0,2);
		MMtil = strTil.substring(3,5)/60;
		strFra = tt+":"+mm;
		var minFra = mm/60;
		
		var fraTid = parseFloat(tt) + parseFloat(minFra);
		var tilTid = parseFloat(TTtil) + parseFloat(MMtil);
		
	break;
	
	case 'tbxTilKl':
		strLunsj=document.forms[formNr2].elements['tbxLunsj'].value;			
		TTlunsj = strLunsj.substring(0,2);
		MMlunsj = strLunsj.substring(3,5)/60
		var lunsjTid = parseFloat(TTlunsj) + parseFloat(MMlunsj);	

		strFra=document.forms[formNr2].elements['tbxFraKl'].value;	
		TTfra = strFra.substring(0,2);
		MMfra = strFra.substring(3,5)/60;
		strTil = tt+":"+mm;	
		var minTil = mm/60;

		var fraTid = parseFloat(TTfra) + parseFloat(MMfra);
		var tilTid = parseFloat(tt) + parseFloat(minTil);

	break;
	
	case 'tbxLunsj':
		strLunsj=document.forms[formNr2].elements['tbxLunsj'].value;
		strTil=document.forms[formNr2].elements['tbxTilKl'].value;
		strFra=document.forms[formNr2].elements['tbxFraKl'].value;	
		TTfra = strFra.substring(0,2);
		MMfra = strFra.substring(3,5)/60;
		strTil=document.forms[formNr2].elements['tbxTilKl'].value;
		TTtil = strTil.substring(0,2);
		MMtil = strTil.substring(3,5)/60;
		
		var fraTid = parseFloat(TTfra) + parseFloat(MMfra);
		var tilTid = parseFloat(TTtil) + parseFloat(MMtil);
		
		strLunsj = tt+":"+mm;
		var minLunsj = mm/60;
		if (!strLunsj){
			alert("null");
			lunsjTid=0;
		}

		var lunsjTid = parseFloat(tt) + parseFloat(minLunsj);

} //switch

if (isNaN(lunsjTid)){
	lunsjTid=0;
}

if (tilTid>fraTid) {
	var temp = parseFloat(tilTid) - parseFloat(fraTid) - parseFloat(lunsjTid);
}
else if (tilTid==fraTid) {
	var temp = 0;
}
else { // tilTid<FraTid
	var temp = 24 - parseFloat(fraTid) + parseFloat(tilTid) - parseFloat(lunsjTid);
}

arbeidstid = Math.round(temp*100)/100; // two decimals
document.forms[formNr2].elements['tbxTimerPrDag' ].value = arbeidstid;
}
}
//####################################################################################################

function workTime2(formNavn,elementNavn){
	formNr2 = finnFormNr(formNavn);
//alert(formNr2);
	workTime(elementNavn);
}

/*#####################################################################################################
######                     merker all tekst inne i formselementer 			###############	    
######                     denne teksten legges inn i formselementer:                   ###############
######                     onfocus=merkTekst(this.form.name,this.name)                  ###############
######################################################################################################*/
	
function merkTekst(form,felt) {
document.forms[0].elements[felt].select();
//alert(form);
//alert(felt);
}


function dateInterval(formNr0,felt){

	formNr0 = formNr0.name;

	var strTil;
	var strFra;

	F = document.forms[formNr0].elements['tbxFraDato'].value;
	T = document.forms[formNr0].elements['tbxTilDato'].value;

	if (boltest == true && boltest2 == true && F != "" && T != "")
	{

		strFra = document.forms[formNr0].elements['tbxFraDato'].value;
		strTil = document.forms[formNr0].elements['tbxTilDato'].value;

		var fraTest = parseInt(strFra.substring(6,8));
		var tilTest = parseInt(strTil.substring(6,8));

		if(fraTest < 30 && fraTest >= 0)
			var fraAarh = "20"+strFra.substring(6,8);
		else
			var fraAarh = "19"+strFra.substring(6,8);
			
		if(tilTest < 30 && tilTest >=0)
			var tilAarh = "20"+strTil.substring(6,8);
		else
			var tilAarh = "19"+strTil.substring(6,8);
			
		var stripTil = parseInt(tilAarh+strTil.substring(3,5)+strTil.substring(0,2));
		var stripFra = parseInt(fraAarh+strFra.substring(3,5)+strFra.substring(0,2));
		
		if (stripFra > stripTil){
			alert("fraDato er større enn tilDato");
			document.forms[formNr0].elements[aktivtFelt].focus();
		}
	}	

}
/*#####################################################################################################
######                     Display Progress Bar in Accrued report			###############	    
######                                        ###############
######                                       ###############
######################################################################################################*/


function activateProgressBar(imgSource)
{
  ctrl.src=imgSource;
  ctrl.style.display='block';
  document.getElementById("view").disabled =true;
  document.getElementById("tildato").disabled =true;
  document.getElementById("dbxAvdeling").disabled =true;
  document.getElementById("ikke_fakt").disabled =true;
  document.getElementById("ikke_loennet").disabled =true;
  
}

function displayProgressBar()
 {  
  
 //Ltrim the date
 strDate =  document.getElementById("tildato").value.replace(/^\s+/,"");
 
 //Rtrim the date
 strDate =  strDate.replace(/\s+$/,"");

 if(strDate  =="")
 {
   alert('Enter the to date');
 }
 else
 {

  ctrl= document.getElementById("prgBar");

  setTimeout('activateProgressBar("../images/Export_Wait.gif")');
 
 } 
 }