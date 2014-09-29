		// Menu in the contentHead1
		function menuOver(which)
		{
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "#ffffff";
			oSelected.style.borderTop = "1px solid #000000";
			oSelected.style.borderRight = "1px solid #999999";
			oSelected.style.borderBottom = "1px solid #999999";
			oSelected.style.borderLeft = "1px solid #000000";
		}
		
		function menuOut(which)
		{
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "transparent";
			oSelected.style.borderTop = "1px solid #999999";
			oSelected.style.borderRight = "1px solid #000000";
			oSelected.style.borderBottom = "1px solid #000000";
			oSelected.style.borderLeft = "1px solid #999999";
		}

		// Menu2 after the contentHead1
		function menu2Over(which)
		{
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "#ffffff";
			oSelected.style.borderTop = "1px solid #737371";
			oSelected.style.borderRight = "1px solid #737371";
			oSelected.style.borderLeft = "1px solid #737371";
		}
		
		function menu2Out(which)
		{
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "transparent";
			oSelected.style.borderTop = "1px solid #999999";
			oSelected.style.borderRight = "1px solid #999999";
			oSelected.style.borderLeft = "1px solid #999999";
		}
		
		// styles for the page elements (closing, opens)
		function toggleOver(which)
		{
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "#a5b2c0";
			oSelected.style.cursor = "hand";
		}
		
		function toggleOut(which){
			var oSelected = document.getElementById(which);
			oSelected.style.backgroundColor = "#d6dee7";
		}
		
		function toggle(sID,sBarID){
			var oSelected = document.getElementById(sID);
			var oParent = document.getElementById(sBarID);
			if (oSelected.style.display == "none")
			{
					oParent.className = "contentHead contentHeadOpen";
					oSelected.style.display = "";
			}
			else
			{
				oParent.className = "contentHead contentHeadClosed";
				oSelected.style.display = "none";
			}
		}
		

