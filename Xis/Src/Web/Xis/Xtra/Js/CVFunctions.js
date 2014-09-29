	function Vis_CV(nVikarID) {
		if (document.all.cboCVChoice.value=="1")
		{
			//window.location = 'vikarCVutskrift.asp?VikarID='+nVikarID;
			window.open('vikarCVutskrift.asp?VikarID=' + nVikarID, 'CV');
			document.all.cboCVChoice.selectedIndex = 0;
		}
		else if (document.all.cboCVChoice.value == "2")
		{
			window.location = 'vikarCVnyPersonalia.asp?VikarID=' + nVikarID;
		}
		else if (document.all.cboCVChoice.value=="3")
		{
			window.location = 'vikarCVvis.asp?VikarID=' + nVikarID;
		}
	}

