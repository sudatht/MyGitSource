VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0
Private Const CREDENTIAL_HIVE = "SOFTWARE\Electric Farm\Xtra\Xis"

Public Function EnsurePathExists(ByVal path As String) As Boolean
    Dim objFSO As Scripting.FileSystemObject
    Call Logon
    On Error GoTo LblAccessError
        
    Set objFSO = New Scripting.FileSystemObject

    If (Not objFSO.FolderExists(path)) Then
        objFSO.CreateFolder (path)
    End If
    
    Set objFSO = Nothing
    EnsurePathExists = True
    On Error GoTo 0
    Call Logoff
    Exit Function
LblAccessError:
    On Error GoTo 0
    EnsurePathExists = False
End Function
    
Public Sub Logon()
    Dim lngTokenHandle, lngLogonType, lngLogonProvider As Long
    Dim blnResult As Boolean
    
    lngLogonType = LOGON32_LOGON_INTERACTIVE
    lngLogonProvider = LOGON32_PROVIDER_DEFAULT
    
    blnResult = RevertToSelf()
    
    blnResult = LogonUser(GetCredentialFromRegistry("Username"), GetCredentialFromRegistry("Domain"), GetCredentialFromRegistry("Password"), lngLogonType, lngLogonProvider, lngTokenHandle)
    
    blnResult = ImpersonateLoggedOnUser(lngTokenHandle)
    CloseHandle (lngTokenHandle)
End Sub

Public Sub Logoff()
    Dim blnResult As Boolean

    blnResult = RevertToSelf()
End Sub


Private Function GetCredentialFromRegistry(keyName As String) As String
    Dim regXis As XisUtils.Registry
    
    Set regXis = New XisUtils.Registry
    GetCredentialFromRegistry = regXis.GetKeyValue(CREDENTIAL_HIVE, keyName)
    Set regXis = Nothing
End Function

