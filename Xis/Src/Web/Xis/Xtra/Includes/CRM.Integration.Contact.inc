<%
	Function GetCRMContactPersonName(SOpersonID)
		dim aXmlHTTP
		dim personRs
		
		Set aXmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		aXmlHTTP.Open "POST", "http://xis/Xtra/CRMIntegration/Responder.aspx?PageMode=GetCustomerContactName&Sopeid=" + Cstr(SOpersonID) , false, Application("Xtra_Domain") +"\" + Application("Xtra_User"), Application("Xtra_Password")
		aXmlHTTP.send ""

		GetCRMContactPersonName = aXmlHTTP.responseText

	End Function
	
%>