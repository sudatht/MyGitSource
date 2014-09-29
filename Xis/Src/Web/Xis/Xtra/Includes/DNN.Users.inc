<%
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
' Solution Name     : XtraDNN
' Project Name      : UserService
' Coding Standard   : EC Coding Standards
' Author            : cmu
' Created Timestamp : 31/05/2006 10:57  AM
' --------------------------------------------------------------------------

Function CreateUserDom(iApp, sUserRef, sUserName, sPassword, sFirstName, sLastName, sEmail,sUserType)
    Dim objXml 
    Dim objElem
    Dim objRoot
        
    Set objXml = SErver.CreateObject("Microsoft.XMLDOM")
  
    Set objRoot = objXml.appendChild(objXml.createElement("user"))
    Set objElem = objRoot.appendChild(objXml.createElement("application"))
    objElem.Text = iApp
    Set objElem = objRoot.appendChild(objXml.createElement("userRef"))
    objElem.Text = sUserRef
    Set objElem = objRoot.appendChild(objXml.createElement("externalPassword"))
    objElem.Text = ""
    Set objElem = objRoot.appendChild(objXml.createElement("externalUserName"))
    objElem.Text = ""    
    Set objElem = objRoot.appendChild(objXml.createElement("userName"))
    objElem.nodeTypedValue = sUserName
    Set objElem = objRoot.appendChild(objXml.createElement("password"))
    objElem.Text = sPassword
    Set objElem = objRoot.appendChild(objXml.createElement("firstName"))
    objElem.Text = sFirstName
    Set objElem = objRoot.appendChild(objXml.createElement("lastName"))
    objElem.Text = sLastName
    Set objElem = objRoot.appendChild(objXml.createElement("email"))
    objElem.Text = sEmail
    Set objElem = objRoot.appendChild(objXml.createElement("UserType"))
    objElem.Text = sUserType
   
    Set CreateUserDom = objXml
End Function

Function AppendRoles(objUserDom, arrRoles)
    Dim objRoot
    Dim objElem
    Dim i
    Dim cnXis
    
    Set cnXis = GetClientConnection(GetConnectionstring(XIS, ""))
    Set objRoot = objUserDom.documentElement
    Set objRoot = objRoot.appendChild(objUserDom.createElement("roles"))
    For i = 0 To UBound(arrRoles)
		Set rsRole = GetFirehoseRS("SELECT Navn FROM TJENESTEOMRADE WHERE TomID =" & arrRoles(i) , cnXis)
		If (HasRows(rsRole)) Then
			Set objElem = objRoot.appendChild(objUserDom.createElement("role"))
    		objElem.Text = rsRole("Navn").Value
		End If
    Next
    Set AppendRoles = objUserDom
End Function

%>