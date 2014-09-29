VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ClsDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Object Interface
Implements IdataAccess
'MTS
Implements ObjectControl

'variables used to access database
Private m_objCon As adodb.Connection
Private m_objCmd As adodb.Command
Private m_objRs As adodb.Recordset

'MTS Object context
Private m_objCtx As ObjectContext

'Error constants, Error consts for this class starts at 700
Private Const ERR_LOGON_NUMBER As Integer = 700
Private Const ERR_LOGON_DESC As String = "Error during logon. "
Private Const ERR_NOFPARAMETER_NUMBER As Integer = 701
Private Const ERR_NOFPARAMETER_DESC As String = "invalid number of parameters! "
Private Const ERR_PARAMETER_NUMBER As Integer = 702
Private Const ERR_PARAMETER_DESC As String = "Error during parameter transfer. "
Private Const ERR_FETCHDATA_NUMBER As Integer = 703
Private Const ERR_FETCHDATA_DESC As String = "Error during data retrieval. "
Private Const ERR_CRUD_NUMBER As Integer = 704
Private Const ERR_CRUD_DESC As String = "Error during Crud-operation. "
Private Const ERR_BATCHUPDATE_NUMBER As Integer = 705
Private Const ERR_BATCHUPDATE_DESC = "Error during batchupdate. "

'Returns a parameter collection using stored procedure with parameters passed through command object
'For stored procedures that only returns a couple of values
Public Sub IdataAccess_FetchReturnValues(ByVal StrCon As String, ByVal strStoredProcedure As String, ByRef Objcmd As adodb.Command)
    'Initialize ado-objects
    Set m_objCtx = GetObjectContext()
    
    On Error GoTo FetchData_ErrHandler
    'initialize the command object with type
    With Objcmd
        .CommandText = strStoredProcedure
        .CommandType = adCmdStoredProc
        'Connect to the database, JIT
        Call Logon(StrCon)
        Set .ActiveConnection = m_objCon
        'Retrieve the data
        .Execute
        Set .ActiveConnection = Nothing
    End With
    Call CleanUp
    'Tell MTS that operation was successful, vote OK
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetComplete
    End If
    Exit Sub

FetchData_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_FETCHDATA_NUMBER, App.EXEName, ERR_FETCHDATA_DESC & Err.Description

End Sub



'***************************************************************************
'NAME: IdataAccess_FetchData()
'
'PURPOSE: Executes the stored procedure specified in storedProc using the
'         parameters in Vparams(), if any. Returns a disconnected populated
'         adodb.recordset if sucessfull, a recordset set to nothing otherwise.
'
'PARAMETERS:
'StrCon - Valid Connection string
'StoredProc - Name of stored procedure to execute
'
'***************************************************************************
Public Function IdataAccess_FetchRCSP(ByVal StrCon As String, ByVal strStoredProcedure As String, colparams As adodb.Parameters) As adodb.Recordset
   Dim ObjParam As adodb.Parameter
   'Initialize ado-objects
    Set m_objCtx = GetObjectContext()
    Set m_objCmd = New adodb.Command
    Set m_objRs = New adodb.Recordset
    
    On Error GoTo FetchData_ErrHandler
    'initialize the command object with type and parameters
    With m_objCmd
        .CommandText = strStoredProcedure
        .CommandType = adCmdStoredProc
        For Each ObjParam In colparams
            .Parameters.Append ObjParam
        Next ObjParam
    End With
        
    'Connect to the database, JIT
    Call Logon(StrCon)
    Set m_objCmd.ActiveConnection = m_objCon
    
    'Retrieve the data
    Set m_objRs = m_objCmd.Execute()
    
    Set IdataAccess_FetchRCSP = m_objRs.Clone
    'Disconnect the recordsets
    Set IdataAccess_FetchRCSP.ActiveConnection = Nothing
    Set m_objRs.ActiveConnection = Nothing
    Call CleanUp
    'Tell MTS that operation was successful, vote OK
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetComplete
    End If
    Exit Function

FetchData_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_FETCHDATA_NUMBER, App.EXEName, ERR_FETCHDATA_DESC & Err.Description
End Function

'***************************************************************************
'NAME: IdataAccess_FetchData()
'
'PURPOSE: Executes a SQL string and Returns a disconnected
'         populated adodb.recordset if sucessfull, a recordset set to nothing otherwise.
'
'PARAMETERS:
'StrCon - Valid Connection string
'StoredProc - Name of stored procedure to execute
'VParams - Array of parameters clustered in 4's - name, type, value, direction
'
'LIMITATIONS - Doesn't currently support return parameters, executes
'              stored procedures only, not SQL-strings.
'***************************************************************************
Public Function IdataAccess_FetchRC(ByVal StrCon As String, ByVal strSQL As String) As adodb.Recordset

    On Error GoTo FetchData_ErrHandler
   'Initialize ado-objects
    Set m_objCtx = GetObjectContext()
    Set m_objRs = New adodb.Recordset
    
    'Connect to the database, JIT
    Call Logon(StrCon)
    
    'Retrieve the data
     m_objRs.Open strSQL, m_objCon, , adLockBatchOptimistic
    
    Set IdataAccess_FetchRC = m_objRs.Clone()
    'Disconnect the recordsets
    Set IdataAccess_FetchRC.ActiveConnection = Nothing
    Set m_objRs.ActiveConnection = Nothing
    Call CleanUp
    'Tell MTS that operation was successful, vote OK
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetComplete
    End If
    Exit Function

FetchData_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_FETCHDATA_NUMBER, App.EXEName, ERR_FETCHDATA_DESC & Err.Description
End Function


'***************************************************************************
'NAME: IdataAccess_ExecuteCRUDSP() (Create,Update,delete -> CRUD)
'
'PURPOSE: Add the specified parameters in the colparams collection to the
'         global m_objCmd and executes the stored procedure specified in
'         storedProc. Returns true if sucessfull, false otherwise.
'
'PARAMETERS:
'StrCon - Valid Connection string
'strStoredProcedure  - Name of stored procedure to execute
'colparams - collection of parameters to append to command object
'
'***************************************************************************

Public Function IdataAccess_ExecuteCRUDSP(ByVal StrCon As String, ByVal strStoredProcedure As String, colparams As adodb.Parameters) As Boolean
    Dim ObjParam As adodb.Parameter
    
    'Initialize ado-objects
    Set m_objCmd = New adodb.Command
    
    On Error GoTo ExecuteCRUD_ErrHandler
    
    'set type to stored procedure and Add the parameters
    With m_objCmd
        .CommandText = strStoredProcedure
        .CommandType = adCmdStoredProc
        For Each ObjParam In colparams
            .Parameters.Append ObjParam
        Next ObjParam
    End With
    
    m_objCmd.Parameters.Refresh
    'Connect to the database, JIT
    Call Logon(StrCon)
    Set m_objCmd.ActiveConnection = m_objCon
    'Perform the CRUD operation
    m_objCmd.Execute
    'Free up the connection object
    IdataAccess_ExecuteCRUDSP = True
    Call CleanUp
    'Tell MTS that operation was successful, vote OK
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetComplete
    End If
    Exit Function
    
ExecuteCRUD_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    IdataAccess_ExecuteCRUDSP = False
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_CRUD_NUMBER, App.EXEName, ERR_CRUD_DESC & Err.Description
End Function

'Executes create update, delete (CRUD) using SQL statement (no parameters )
Public Function IdataAccess_ExecuteCRUD(ByVal StrCon As String, ByVal strSQL As String) As Boolean

    'Initialize ado-objects
    Set m_objCmd = New adodb.Command
    
    On Error GoTo ExecuteCRUD_ErrHandler
    
    With m_objCmd
        .CommandText = strSQL
        .CommandType = adCmdText
    End With
    'Connect to the database, JIT
    Call Logon(StrCon)
    Set m_objCmd.ActiveConnection = m_objCon
    'Perform the CRUD operation
    m_objCmd.Execute
    'Free up the connection object
    IdataAccess_ExecuteCRUD = True
    Call CleanUp
    'Tell MTS that operation was successful, vote OK
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetComplete
    End If
    Exit Function
    
ExecuteCRUD_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    IdataAccess_ExecuteCRUD = False
    If Not m_objCtx Is Nothing Then
        m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_CRUD_NUMBER, App.EXEName, ERR_CRUD_DESC & Err.Description

End Function


'***************************************************************************
'NAME: IdataAccess_batchUpdate()
'
'PURPOSE: reconnects and Batchupdates TargetRC, returns true if succesfull,
'         false if an error occurs.
'
'PARAMETERS:
' StrCon - Valid Connection string
' TargetRC - Adodb.recordset to be updated
'***************************************************************************

Public Function IdataAccess_batchUpdate(ByVal StrCon As String, ByVal TargetRC As adodb.Recordset) As Boolean
  On Error GoTo Err_batchUpdate
  Call Logon(StrCon)
  'reconnect the recordset
  Set TargetRC.ActiveConnection = m_objCon
  TargetRC.UpdateBatch
  IdataAccess_batchUpdate = True
  Set TargetRC = Nothing
  Call CleanUp
  If Not m_objCtx Is Nothing Then
    m_objCtx.SetComplete
  End If
  Exit Function

Err_batchUpdate:
  'Tell MTS that operation was unsuccessful, vote ABORT
  IdataAccess_batchUpdate = False
  If Not m_objCtx Is Nothing Then
    m_objCtx.SetAbort
  End If
  Call CleanUp
  Err.Raise vbObjectError + ERR_BATCHUPDATE_NUMBER, App.EXEName, ERR_BATCHUPDATE_DESC & Err.Description
End Function

'***************************************************************************
'NAME: Logon()
'
'PURPOSE: Logs on the database using the SrtrCon parameter, connection is
'prepared for disconnected rc's.
'
'PARAMETERS:
' StrCon - Valid Connection string
'***************************************************************************

Private Sub Logon(StrCon As String)
  On Error GoTo logon_ErrHandler
    Set m_objCon = New adodb.Connection
    'Use client side cursorlocation for disconnected RC's
    m_objCon.CursorLocation = adUseClient
    m_objCon.ConnectionString = StrCon
    'Open connection
    m_objCon.Open
    Exit Sub
    
logon_ErrHandler:
    'Tell MTS that operation was unsuccessful, vote ABORT
    If Not m_objCtx Is Nothing Then
      m_objCtx.SetAbort
    End If
    Call CleanUp
    Err.Raise vbObjectError + ERR_LOGON_NUMBER, App.EXEName, ERR_LOGON_DESC & Err.Description
End Sub

Private Sub ObjectControl_Activate()
    'Get the context of which the parent application is running under
    Set m_objCtx = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    'This object cannot be pooled (No Visual Basic COM-objects can)
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objCtx = Nothing
End Sub

'***************************************************************************
'NAME: CleanUp()
'
'PURPOSE: Cleans up by Setting the global variables to nothing
'***************************************************************************
Private Sub CleanUp()
    If Not m_objCon Is Nothing Then
        Set m_objCon = Nothing
    End If
    Set m_objRs = Nothing
    Set m_objCmd = Nothing
End Sub
