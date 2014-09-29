VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IdataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Fetches a disconnected rc using stored procedure with parameters passed through parameter collection
'for stored procedures that return many values
Public Function FetchRCSP(ByVal StrCon As String, ByVal strStoredProcedure As String, colparams As adodb.Parameters) As adodb.Recordset
End Function

'Returns a parameter collection using stored procedure with parameters passed through parameter collection
' For stored procedures that only returns a couple of values (counts ++)
Public Sub FetchReturnValues(ByVal StrCon As String, ByVal strStoredProcedure As String, ByRef Objcmd As adodb.Command)
End Sub

'Fetches a disconnected rc using SQL
'for stored procedures that return many values
Public Function FetchRC(ByVal StrCon As String, ByVal strSQL As String) As adodb.Recordset
End Function

'Executes create update, delete (CRUD) using stored procedure with parameters passed through parameter collection
Public Function ExecuteCRUDSP(ByVal StrCon As String, ByVal strStoredProcedure As String, colparams As adodb.Parameters) As Boolean
End Function
'Executes create update, delete (CRUD) using SQL statement (no parameters )
Public Function ExecuteCRUD(ByVal StrCon As String, ByVal strSQL As String) As Boolean
End Function

'Reconnects a disconnected recordset and batchupdates
Public Function BatchUpdate(ByVal StrCon As String, ByVal TargetRC As adodb.Recordset) As Boolean
End Function






