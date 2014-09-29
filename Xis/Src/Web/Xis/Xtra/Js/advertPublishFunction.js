
function advertNavigation(oppdragID) {

if(document.all.cboCVChoice.value !="")
{
          

		if (document.all.cboCVChoice.value =="view")
		{            window.open('WebUI\\PreviewAdd.aspx?CommId=' + oppdragID, '','menubar=yes,toolbar=yes,status=yes,scrollbars=yes');
			

		
		}
                else if(document.all.cboCVChoice.value =="unpublish")
   		{
                    document.all.cboCVChoice.disabled=true;
                    window.location = 'WebUI\\WizardAP.aspx?CommId=' + oppdragID + '&unpublish=YES';		   
	        }

                
                else
		{
                 document.all.cboCVChoice.disabled=true;
		 window.location = 'WebUI\\WizardAP.aspx?CommId=' + oppdragID + '&unpublish=';
	        }

}

}