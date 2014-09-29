var xmlHttp;

function asynchronousCall(urlPath,cbFunc)
{
	if (window.ActiveXObject) 
	{
		xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
	}
	else if (window.XMLHttpRequest) 
	{
		xmlHttp = new XMLHttpRequest();
	}
	
	xmlHttp.open("POST", urlPath, true);
	xmlHttp.setRequestHeader("Content-Type", "text/html; Charset=ISO-8859-1"); 
	
	xmlHttp.onreadystatechange = function()
	{
		if(xmlHttp.readyState == 4 && xmlHttp.status == 200) 
		{
			if(xmlHttp.responseText)
				cbFunc(xmlHttp.responseText);
		}
	};
	
	xmlHttp.send(null);
}
