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
' Description       : This script will display start page for user on application start up
'		      
' Author            : SKA
' Created Timestamp : 18/01/2007
' -------------------------------------------------------------------------- 


'display start page for user on application start up
Sub OnStartup()
		
	'SOMessageBox "Welcome to SuperOffice"
    
    	Dim objSO
	Set objSO = CreateObject("SuperOffice.Application")
	If not (objSO is nothing) Then
    		objSO.ShowUrl "superoffice:browserpanel.x-is"
	else
    		MsgBox "Unable to connect"
	end if
	set objSO = Nothing
	
End Sub 
