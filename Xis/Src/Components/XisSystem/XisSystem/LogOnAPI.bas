Attribute VB_Name = "LogOnAPI"
Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Public Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long

Public Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


