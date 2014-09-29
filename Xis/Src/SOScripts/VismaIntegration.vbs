' --------------------------------------------------------------------------
' Copyright(c) 2000-2006 Eurocenter DDC.
' No 65, Walukarama Road, Colombo 3, Sri Lanka
' All rights reserved.
'
' This software is the confidential and proprietary information of 
' Eurocenter DDC (Confidential Information). You shall not disclose such
' Confidential Information and shall use it only in accordance with the
' terms of the license agreement you entered into with Eurocenter.
'
' Description       : This script will update customer in Visma when a company change occurs in 'SuperOffice.
' Author            : PDF
' Created Timestamp : 15/06/2007
' -------------------------------------------------------------------------- 

'Called when a new customer is saved for the first time.
Sub OnCurrentContactCreated()
	CustomerUpdate()
End Sub

'Called when an existing company is saved after having been changed.'
Sub OnCurrentContactSaved
	CustomerUpdate()
End Sub

'Called after contact person of a company changed
Sub OnCurrentPersonSaved
	ContactPersonUpdate()
End Sub

'Called when a new contact person is saved for the first time.
Sub OnCurrentPersonCreated()
	ContactPersonUpdate()
End Sub

'Update Visma customer
Sub CustomerUpdate()
	Dim dataXml
	Dim methodXml
'MsgBox "OnCurrentContactSaved"
	dataXml = GetCustomerXml()
'MsgBox dataXml
	methodXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<UpdateCustomer xmlns=""http://www.xtra.no/xtra/WebUI/VismaIntegration/VismaServices.asmx"">" & _
             "<cusXml><![CDATA["& dataXml &"]]></cusXml>" & _ 
             "</UpdateCustomer>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
'MsgBox methodXml
	CallWebService "UpdateCustomer", methodXml
End Sub

'Update Visma contact person
Sub ContactPersonUpdate()
	Dim dataXml
	Dim methodXml
'MsgBox "Visma integration : OnCurrentPersonSaved", vbInformation + vbOkOnly, "Visma"
	dataXml = GetContactPersonXml()
'MsgBox dataXml
	methodXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<UpdateContactPerson xmlns=""http://www.xtra.no/xtra/WebUI/VismaIntegration/VismaServices.asmx"">" & _
             "<personXml><![CDATA["& dataXml &"]]></personXml>" & _ 
             "</UpdateContactPerson>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
	CallWebService "UpdateContactPerson", methodXml
End Sub

'web service call
function CallWebService(method, arg)
	Dim xmlDOC
	Dim bOK
	Dim HTTP
	Dim objDom
	Dim user
	Dim password
	Dim url
'MsgBox "getResults called"

	Set HTTP = CreateObject("MSXML2.XMLHTTP")
'MsgBox "getResults called 1"
	If IsObject(HTTP) Then
'MsgBox "getResults called 1.1"
		url = GetVismaWebServiceUrl() '"http://xis.dev.ec/xtra/WebUI/VismaIntegration/VismaServices.asmx"
'MsgBox url
		user = GetXisUser()
		password = GetXisUserPassword()
'MsgBox "User: "& user& " Password: "& password
		HTTP.Open "POST", url, false, user, password
			
'MsgBox "getResults called 2"
		HTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		HTTP.setRequestHeader "SoapAction","http://www.xtra.no/xtra/WebUI/VismaIntegration/VismaServices.asmx/" & method
	
		HTTP.send arg
		
'MsgBox "Web service response: " & HTTP.responseText
		
		Set HTTP = nothing
	Else
		MsgBox "Couldn't send data to Visma"
	End If

	
End Function



Sub CustomerCreate
'MsgBox "Visma integration: OnCurrentContactSaved", vbInformation + vbOkOnly, "Visma"
	On Error Resume Next	          
        Dim SOClient
	Set SOClient = CreateObject("SuperOffice.Application")
	
	if IsObject(SOClient) then
	
		Dim cur
'MsgBox "OnCurrentContactSaved 1", vbInformation + vbOkOnly, "Visma" 
		Set cur = SOClient.CurrentContact
		if err.number<>0 then
'MsgBox "OnCurrentContactSaved: SOClient.CurrentContact Error: " & Err.description
			Exit Sub
		End If
'MsgBox "OnCurrentContactSaved 2", vbInformation + vbOkOnly, "Visma" 
'MsgBox "The current contact id is: " & cur.Identity & " The current name is: " & cur.Name, vbInformation + vbOkOnly, "Visma"
	else
		MsgBox "Unable to connect to super office"
	end if

'MsgBox "OnCurrentContactSaved >> Done!", vbInformation + vbOkOnly, "Visma"
	CollectGarbage
End Sub

'Return customer xml for given SO customer object
Function GetCustomerXml()
	
	'xml variables
	Dim objDom
	Dim objRoot
	Dim objChild1
	Dim objChild2,objChild3,objAddr
	Dim objPI
	'SO variables	        
        Dim SOClient
        
        On Error Resume Next	  
                
	Set SOClient = CreateObject("SuperOffice.Application")
	
	if IsObject(SOClient) then
	
		Dim cur
		Set cur = SOClient.CurrentContact
	
		'Create customer xml from SO Contact object
		Set objDom = CreateObject("Microsoft.XMLDOM")
		Set objRoot = objDom.createElement("ROOT")
		objDom.appendChild objRoot
		
		'Company Id
		AddDomElement objDom, objRoot, "COMPANYID", cur.Number2 'rsContact("FirmaID")
		'Firma Id from Xis
		AddDomElement objDom, objRoot, "NAME", cur.Name
		'Company code
		AddDomElement objDom, objRoot, "CODE", cur.Number1
		'Organization number
		AddDomElement objDom, objRoot, "ORGNO", cur.OrgNr
		
		'Postal address
		Set objChild1 = AddDomElement(objDom, objRoot, "POSTALADDRESS", "")
		'Address 1
		AddDomElement objDom, objChild1, "ADDRESS1", cur.Postaladdress.Address1
		'Address 2
		AddDomElement objDom, objChild1, "ADDRESS2", cur.Postaladdress.Address2
		'Post code
		AddDomElement objDom, objChild1, "POSTCODE", cur.Postaladdress.ZipCode
		'Post area
		AddDomElement objDom, objChild1, "POSTAREA", cur.Postaladdress.City
		
		'Invoice address
		Set objChild2 = AddDomElement(objDom, objRoot, "INVOICEADDRESS", "")
		'Address 1
		AddDomElement objDom, objChild2, "ADDRESS1", cur.UDef.ByName("Fakturaadresse:").Value
		'Post code
		AddDomElement objDom, objChild2, "POSTCODE", cur.UDef.ByName("Postnr:").Value
		'Post area
		AddDomElement objDom, objChild2, "POSTAREA", cur.UDef.ByName("Sted:").Value
		
		'County
		AddDomElement objDom, objRoot, "COUNTRY", cur.Country
		'Xis responsible person for customer
		AddDomElement objDom, objRoot, "CONTACTPERSON", cur.Associate.FullName
		'Category
		AddDomElement objDom, objRoot, "CATEGORY", cur.Category.text
		'Business
		AddDomElement objDom, objRoot, "BUSINESS", cur.Business.text
		'Telephone
		If cur.phones.Exists (1) then
		      AddDomElement objDom, objRoot, "TELEPHONE", cur.Phones.Item(1).number
		End if
		'Fax
		Dim no
		Dim counter
		no = cur.Phones.Count
		For counter = 1 to no
			If ( cur.Phones.Item(counter).Type=3 ) Then
				AddDomElement objDom, objRoot, "FAX", cur.Phones.Item(counter).number
				break
			End If
		Next
		
		'Email
		If cur.Emails.Exists(1) then
			AddDomElement objDom, objRoot, "EMAIL", cur.Emails.Item(1).address
		End if
		
		'Set objPI = objDom.createProcessingInstruction("xml","version='1.0'")
	
		'objDom.insertBefore objPI, objDom.childNodes(0)
		
		'Return customer xml
		GetCustomerXml = objDom.xml
		'test 			
		'objDom.Save "H:\SO_ARC\Scripts\customer.xml"
            	
	else
		MsgBox "Unable to connect to super office"
	end if
	
'MsgBox "GetCustomerXml >> Done!"
		
	CollectGarbage
	
End Function

'Return customer xml for given SO customer object
Function GetContactPersonXml()
	
	'xml variables
	Dim objDom
	Dim objRoot
	Dim objChild1
	Dim objChild2,objChild3,objAddr
	Dim objPI
	'SO variables	        
        Dim SOClient
        
        On Error Resume Next	  
        
        Set SOClient = CreateObject("SuperOffice.Application")
       
	if IsObject(SOClient) then
		
		    	
	    	Set objperson = SOClient.CurrentPerson
	    	
	    	
		'Create customer xml from SO Contact object
		Set objDom = CreateObject("Microsoft.XMLDOM")
		Set objRoot = objDom.createElement("ROOT")
		objDom.appendChild objRoot
		
		'Person ID
		AddDomElement objDom, objRoot, "PERSONID", objperson.Identity
		'Person first name
		AddDomElement objDom, objRoot, "FNAME", objperson.FirstName
		'Person middle name
		AddDomElement objDom, objRoot, "MNAME", objperson.MiddleName
		'Person last name
		AddDomElement objDom, objRoot, "LNAME", objperson.LastName
		
		'Company Id  - number2 filed of customer this person belongs to
		AddDomElement objDom, objRoot, "COMPANYID", objperson.Contact.Number2
		
		'Is person main contavt fo the customer - true/false
		AddDomElement objDom, objRoot, "ISMAINCONTACT", "false"
		'Address 1
		AddDomElement objDom, objRoot, "ADDRESS1", objperson.Address.Address1
		'Address 2
		AddDomElement objDom, objRoot, "ADDRESS2", objperson.Address.Address2
		'Address 3
		AddDomElement objDom, objRoot, "ADDRESS3", objperson.Address.Address3
		'postcode
		AddDomElement objDom, objRoot, "POSTCODE", objperson.Address.ZipCode
		'postoffice
		AddDomElement objDom, objRoot, "POSTOFFICE", objperson.Address.City
		
		
		'SO Phones type - Type is 1=phone, 3=fax, 5=mobile
		Set objChild1 = AddDomElement(objDom, objRoot, "PHONES", "")
		Dim no,teleNo,faxNo,mobileNo
		teleNo = 1
		faxNo = 1
		mobileNo = 1
		Dim counter
		no = objperson.Phones.Count
		For counter = 1 to no
			If ( objperson.Phones.Item(counter).Type=1 ) Then
				AddDomElement objDom, objChild1, "TELEPHONE" & teleNo, objperson.Phones.Item(counter).number
				teleNo = teleNo+1
			End If
			
			If ( objperson.Phones.Item(counter).Type=3 ) Then
				AddDomElement objDom, objChild1, "FAX" & faxNo, objperson.Phones.Item(counter).number
				faxNo=faxNo+1
			End If
			
			If ( objperson.Phones.Item(counter).Type=5 ) Then
				AddDomElement objDom, objChild1, "MOBILE" & mobileNo, objperson.Phones.Item(counter).number
				mobileNo=mobileNo+1
			End If
		Next
		
		'Email
		Set objChild2 = AddDomElement(objDom, objRoot, "EMAILS", "")
		no = objperson.Emails.Count
		For counter = 1 to no
			AddDomElement objDom, objChild2, "EMAIL" & counter, objperson.Emails.Item(counter).address
		Next
						
		'department no
		AddDomElement objDom, objRoot, "DEPNO", objperson.Department
		'title no
		AddDomElement objDom, objRoot, "TITLENO", objperson.Title
		'Set objPI = objDom.createProcessingInstruction("xml","version='1.0'")
	
		'objDom.insertBefore objPI, objDom.childNodes(0)
		
		'Return contact person xml
		GetContactPersonXml = objDom.xml
		'test 			
		'objDom.Save "C:\SO_ARC\Scripts\customer.xml"
            				
	else
		MsgBox "Unable to connect to super office"
	end if
	
'MsgBox "GetCustomerXml >> Done!"
		
	CollectGarbage
	

End Function

'Add element to node in xml 
Function AddDomElement(objParentNode, objRoot, strName, strValue)
	Dim objNode
	Set objNode = objParentNode.createElement(strName)
	objNode.Text = strValue
	objRoot.appendChild objNode
	Set AddDomElement = objNode
End Function

'Add attribute to a tag in xml
Sub AddNodeAttribute(objDom, objNode, strName, strValue)
	Dim objAttrib
	Set objAttrib = objDOM.createAttribute(strName)
	objAttrib.Text =strValue
	objNode.Attributes.setNamedItem objAttrib
	objDOM.documentElement.appendChild objNode
End Sub

'Get Visma web service url from registry
Function GetVismaWebServiceUrl()
	Dim objShell
	Dim url
	Set objShell = CreateObject("WScript.Shell")
	url = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\Settings\VismaWebService")
	GetVismaWebServiceUrl = url
End Function


'Get username to log on to xis from registry
Function GetXisUser()
	 
	 Dim objShell
	 Dim xtra_Username
	 Set objShell = CreateObject("WScript.Shell")
	 
	 xtra_Username = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Electric Farm\Xtra\Xis\Username")
 	 
	 GetXisUser = trim(xtra_Username)
End Function

'Get password to log on to xis from registry
Function GetXisUserPassword()
	 
	 Dim objShell
	 Dim xtra_Password
	 Set objShell = CreateObject("WScript.Shell")
	 
 	 xtra_Password = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Electric Farm\Xtra\Xis\Password")
 	 
	 GetXisUserPassword = trim(xtra_Password)
End Function