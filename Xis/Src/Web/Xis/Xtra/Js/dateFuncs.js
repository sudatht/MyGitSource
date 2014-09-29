var blnErrorOccured = false;

/*####################################################################################################
########################sjekker at fra-dato er før til-dato###########################################
####################################################################################################*/


function dateInterval(formNr0, felt)
{
	if (blnErrorOccured)
	{
		return;
	}

	formNr0 = formNr0.name;

	var strTil;
	var strFra;

	F = document.forms[formNr0].elements['tbxFraDato'].value;
	T = document.forms[formNr0].elements['tbxTilDato'].value;
		
	if  (F!="" && T!="")
	{

		strFra = document.forms[formNr0].elements['tbxFraDato'].value;
		strTil = document.forms[formNr0].elements['tbxTilDato'].value;

		var fraTest = parseInt(strFra.substring(6,8));
		var tilTest = parseInt(strTil.substring(6,8));

		if(fraTest < 30 && fraTest >=0)
			var fraAarh = "20" + strFra.substring(6,8);
		else
			var fraAarh = "19" + strFra.substring(6,8);
			
		if(tilTest < 30 && tilTest >=0)
			var tilAarh = "20" + strTil.substring(6,8);
		else
			var tilAarh = "19" + strTil.substring(6,8);
			
		var stripTil = parseInt(tilAarh + strTil.substring(3,5) + strTil.substring(0,2));
		var stripFra = parseInt(fraAarh + strFra.substring(3,5) + strFra.substring(0,2));

			if (stripFra > stripTil)
			{
				alert("fraDato er større enn tilDato");
				if (blnErrorOccured == false)
				{
					document.forms[formNr0].elements['tbxTilDato'].focus();
				}	
				blnErrorOccured = true;
				return false;
			}
	}	
	blnErrorOccured = false;
	return true;
}