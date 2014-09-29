function ajax(url, controlId, callbackFunction)
{
 
        if (window.XMLHttpRequest)
         {
           var request = new XMLHttpRequest();
        } else {
            var request = new ActiveXObject("MSXML2.XMLHTTP.3.0");
        }
         
        request.open("POST", url, true);
        request.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 
 
        request.onreadystatechange = function()
        {
 
                if (request.readyState == 4 && request.status == 200)
                 {
                    
                        if (request.responseText)
                        {
                             
                               callbackFunction(request.responseText);
                        }
                  }
                  else
                  {
                  	 callbackFunction("");
                  }
          }
                      
          txtbox1=document.getElementById(controlId);
      
          request.send( "PostNr=" + txtbox1.value);
}


function fillPostOffice(PostOffice)
{
	txtbox=document.getElementById("lnk11");
	txtbox.value=PostOffice;
}