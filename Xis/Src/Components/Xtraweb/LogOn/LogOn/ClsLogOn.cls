VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function LogOn(ByVal strConnection As String, ByVal StrUserID As String, StrPwd As String) As String
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As adodb.Command
    Dim ObjParam As adodb.Parameter
    Dim ObjOutputCmd As adodb.Command
    
    'strConnection = "Provider=SQLOLEDB.1;User ID=ebusiness;Password=ebusiness;Data Source=PRAHA;Initial Catalog=XtraImp"
    
    If ((Trim(StrUserID) <> "") And (Trim(StrPwd) <> "")) Then
        Set ObjDAL = New ClsDataAccess
        Set ObjCmd = New adodb.Command
        
        'Add parameters to command object
            Set ObjParam = New Parameter
            ObjParam.Name = "strUserName"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adLongVarChar
            ObjParam.Value = StrUserID
            ObjParam.Size = Len(StrUserID)
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "StrInPassword"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adLongVarChar
            ObjParam.Value = StrPwd
            ObjParam.Size = Len(StrPwd)
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "intUserID"
            ObjParam.Direction = adParamOutput
            ObjParam.Type = adInteger
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "strType"
            ObjParam.Direction = adParamOutput
            ObjParam.Type = adVarChar
            ObjParam.Size = 15
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "int_xtra_id"
            ObjParam.Direction = adParamOutput
            ObjParam.Type = adInteger
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                
            Set ObjParam = New Parameter
            ObjParam.Name = "intchanged_password"
            ObjParam.Direction = adParamOutput
            ObjParam.Type = adInteger
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "strstatus"
            ObjParam.Direction = adParamOutput
            ObjParam.Type = adVarChar
            ObjParam.Size = 20
            ObjCmd.Parameters.Append ObjParam
       
        
        ' Getuserids
        ObjDAL.FetchReturnValues strConnection, "Getuserids", ObjCmd
        Set ObjDAL = Nothing
        For Each ObjParam In ObjCmd.Parameters
            If ObjParam.Direction = adParamOutput Then
                  LogOn = LogOn & ObjParam.Name & ", " & ObjParam.Value & ", "
            End If
        Next ObjParam
        LogOn = Mid(LogOn, 1, Len(LogOn) - 2)
        Set ObjCmd = Nothing
    Else
        LogOn = ""
    End If
End Function


