Sub OnCurrentPersonSaved
Dim XMLString
XMLString = GetPersonXml()
CallWebService "Save","XMLString",XMLString

End Sub

'web service call
function CallWebService(method, Parameter1Name,XMLString)
	Dim xmlDOC
	Dim bOK
	Dim HTTP
	Dim objDom
	MsgBox "getResults called"
	
	Set HTTP = CreateObject("MSXML2.XMLHTTP")
	
	HTTP.Open "POST","http://localhost/TestContactPerson/ContactPerson.asmx", false
	
	
	HTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	HTTP.setRequestHeader "SoapAction","http://tempuri.org/" & method

	XMLString = "<ROOT><Table1><COMPANYID>1110846</COMPANYID><DeliveryName>Brandvold, Grete</DeliveryName><CODE>10024</CODE><DeliveryAddress1>dfgdfgdfsdfgsfghgdfgdfgdfdsf</DeliveryAddress1><DeliveryAddress2></DeliveryAddress2><DeliveryAddress3></DeliveryAddress3><PostCode>2334</PostCode><PostOffice>RQWSOMEDAL</PostOffice><TELEPHONE>23103673</TELEPHONE><FAX>23103681</FAX><EMAIL>grete.brandvold@xtra.no</EMAIL></Table1></ROOT>"
	
	strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<" & method & " xmlns=""http://tempuri.org/"">" & _
              "<" & Parameter1Name & "><![CDATA[" & XMLString & "]]></" & Parameter1Name & ">" & _
             "</" & method & ">" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
         
	HTTP.send strXml
	
	MsgBox "Web service response: " & HTTP.responseText
	
	Set HTTP = nothing
	
End Function




'Return customer xml for given SO customer object
Function GetPersonXml() 
	
	'xml variables
	Dim objDom
	Dim objRoot
        Dim objTable1
	Dim objChild1
	Dim objChild2,objChild3,objAddr
	Dim objPI
        Dim curContact
	'SO variables	        
        Dim SOPerson
        
        On Error Resume Next	  
        
	Set SOPerson = CreateObject("SuperOffice.Application")
	
	if IsObject(SOPerson) then
	
		Dim cur
		Set cur = SOPerson.CurrentPerson
                Set curContact=SOPerson.CurrentContact
		'get xtra firma id for SO contact
		Dim strCon
         	strCon = GetConnectionString() 		
	
	 	Dim Conn 
	 	Set Conn = CreateConnection(strCon)
	 	
	 	'Check if contact exists in x|is
                strSQL = "SELECT [FirmaID] FROM [Firma] WHERE [Socuid] = " & cur.Identity
            	Set rsContact = GetRS(strSQL, Conn)
            	
            	'if contact exist
            	If (Not rsContact.EOF) Then
            		'create xml
     		
			'Create customer xml from SO Contact object
			Set objDom = CreateObject("Microsoft.XMLDOM")
			Set objRoot = objDom.createElement("ROOT")
			objDom.appendChild objRoot
			
			'Company Id
			AddDomElement objDom, objRoot, "COMPANYID", rsContact("FirmaID")
		
            		'create xml
     		
			'Create customer xml from SO Contact object
			Set objDom = CreateObject("Microsoft.XMLDOM")
			Set objRoot = objDom.createElement("ROOT")
			objDom.appendChild objRoot

                        'Table Attribute
                        Set objTable1 = AddDomElement(objDom, objRoot, "Table1", "")                       

 			'Company Id
			AddDomElement objDom, objTable1 , "COMPANYID",curContact.Number2
			'Firma Id from Xis
			AddDomElement objDom, objTable1 , "DeliveryName", cur.FullName
			'Company code
			AddDomElement objDom, objTable1 , "CODE", cur.Number
			
			
			
			'Address 1
			AddDomElement objDom, objTable1 , "DeliveryAddress1", cur.Address.Address1
			'Address 2
			AddDomElement objDom, objTable1 , "DeliveryAddress2", cur.Address.Address2
                        
                        'Address 3
			AddDomElement objDom, objTable1 , "DeliveryAddress3", cur.Address.Address3

			'Post code

			AddDomElement objDom, objTable1 , "PostCode", cur.Address.ZipCode 
        		                 

			'Post area
			AddDomElement objDom, objTable1 , "PostOffice", cur.Address.City
			
			
			
			'County
			AddDomElement objDom, objTable1 , "COUNTRY", cur.Country
			
			'Telephone
			If cur.phones.Exists (1) then
			      AddDomElement objDom, objTable1 , "TELEPHONE", cur.Phones.Item(1).number
			End if
			'Fax
			Dim no
			Dim counter
			no = cur.Phones.Count
			For counter = 1 to no
				If ( cur.Phones.Item(counter).Type=3 ) Then
					AddDomElement objDom, objTable1 , "FAX", cur.Phones.Item(counter).number
					break
				End If
			Next
			'Email
			If cur.Emails.Exists(1) then
				AddDomElement objDom, objTable1 , "EMAIL", cur.Emails.Item(1).address
			End if

			Set objPI = objDom.createProcessingInstruction("xml","version='1.0'")
		
			objDom.insertBefore objPI, objDom.childNodes(0)
			'test 			
			objDom.Save "C:\SO_ARC\Scripts\person1.xml"
                        GetPersonXml= objDom.XML
                ELSE
                    MsgBox "No record"
            	End If
            	rsContact.Close
	        Set rsContact = Nothing
	        
	 	'close connection
	 	CloseConnection(Conn)
					
	else
		MsgBox "Unable to connect to super office"
	end if
	
	MsgBox "GetCustomerXml >> Done!"
		
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

'	*********************************************		'
'		Functions from companyupdate vbs
'	*********************************************		'

'Returns a forward readonly recordset
Function GetRS(strSQL, Conn)
		Dim oRS 
		'MsgBox "GetRS"	
		Set oRS	= CreateObject("ADODB.Recordset")
		'Prepare firehose RS for dataretrieval
		oRS.LockType =	1 	'adLockReadOnly
		oRS.CursorType = 0	'adOpenForwardOnly 
		on error resume next
		oRS.Open strSQL, Conn
		If ( Conn.Errors.Count > 0 ) Then
			on error goto 0
			MsgBox "Error in :" & strSQL
			set GetRS = nothing
			Exit Function
		End If
		on error goto 0
		Set GetRS = oRS
End Function

'Executes a crud SQL statement (Create, Update, Delete and Insert)
Function ExecuteSQL(strSQL, Conn)
		Dim oRS 
		Dim Cmd
		'MsgBox "ExecuteSQL"
		ExecuteSQL = false		
		
		set Cmd = CreateObject("ADODB.Command")
		set Cmd.ActiveConnection = Conn
		Cmd.CommandText = strSQL
		Cmd.CommandType = &H0001 'adCmdText
		on error resume next
		Cmd.Execute 

		If ( Conn.Errors.Count > 0 ) Then
			MsgBox "Error in :" & strSQL
		else
			ExecuteSQL = true
			
		End If
		on error goto 0
		set Cmd.ActiveConnection = nothing
		set Cmd = nothing
End Function

'Read the connection details of the Xis db from the registry
Function GetConnectionString()
	 
	 Dim objShell
	 Set objShell = CreateObject("WScript.Shell")
	 xtra_Provider = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Provider")
	 xtra_Username = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Username")
 	 xtra_Password = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Password")
 	 xtra_DBServer = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\DBServer")
 	 xtra_DBName = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\DBNAME")
	 xtra_Options = objShell.RegRead("HKLM\SOFTWARE\Electric Farm\Xtra\Xtraweb\ConnectionStrings\XtraDefault\Options")
	 	 
	 Dim strConnection
	 	 
         'xtra_Provider = "SQLOLEDB.1"
         'xtra_Username = "xtrawebuser"
         'xtra_Password = "xtrawebuser"
         'xtra_DBServer = "eccoldev04vm02\xtra_integration"
         'xtra_DBName = "Xis"
         'xtra_Options =  "Persist Security Info=True"
         strConnection = "Provider=" & xtra_Provider & ";" &_
				"User ID=" & xtra_Username & ";" &_
				"Password=" & xtra_Password & ";" &_
				"Data Source=" & xtra_DBServer & ";" &_
				"Initial Catalog=" & xtra_DBName  & ";" &_
				xtra_Options
				
	 'MsgBox "Connection : " & strConnection

	 GetConnectionString = trim(strConnection)
End Function

'Creates the db connection
Function CreateConnection(sConString)
	Dim Conn
	Set Conn = CreateObject("adodb.Connection")
	Conn.Open(sConString)
	Set CreateConnection = Conn
End Function

'Closes the db connection
Sub CloseConnection(byref Conn)
		'MsgBox "Close connection"
		If (IsObject(Conn)) Then
			If (not Conn is nothing) Then
				If ( Conn.State <> 0) Then
					Conn.Close()
				End If
			end if
		End If
End Sub


