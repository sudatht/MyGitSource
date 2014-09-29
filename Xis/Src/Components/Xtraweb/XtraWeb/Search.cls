VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStrXtraCon As String

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Public Property Get XtraConString() As String
    XtraConString = mStrXtraCon
End Property

Public Property Let XtraConString(strConString As String)
    mStrXtraCon = strConString
End Property

Public Function GetEmployees(ByVal lLocationID As Long) As ADODB.Recordset
On Error GoTo err_GetEmployees
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As Command
    Dim ObjParam As ADODB.Parameter
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjCmd = New ADODB.Command
    
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "lLocationID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = lLocationID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjDAL = New ClsDataAccess

    Set GetEmployees = ObjDAL.FetchRCSP(mStrXtraCon, "GetEmployees", ObjCmd.Parameters)
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

err_GetEmployees:
    WriteLog "Search_GetEmployees", Error$, EF_LogEventType_Error
End Function

Public Function GetEmployee(ByVal Lvikarid As Long) As ADODB.Recordset
On Error GoTo err_GetEmployee
    Dim strSQL As String
    Dim ObjDAL As IdataAccess

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess

    strSQL = "Select v.* from vikar v Where v.vikarid = " & Lvikarid & " and v.statusid = 4"

    Set GetEmployee = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    Set ObjDAL = Nothing
    Exit Function

err_GetEmployee:
    WriteLog "Search_GetEmployee", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
End Function

Public Function GetLast10Jobs() As ADODB.Recordset
On Error GoTo err_GetLast10Jobs
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As Command
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command

    Set GetLast10Jobs = ObjDAL.FetchRCSP(mStrXtraCon, "getlast10jobs", ObjCmd.Parameters)
    
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
err_GetLast10Jobs:
    WriteLog "Search_GetLast10Jobs", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetJobsInInterval(ByVal lStart As Long, ByVal lInterval As Long, ByVal lLocationID As Long) As Collection
    Dim strSQL As String
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim objJob As Job
    Dim fldData As ADODB.Field
    Dim mCol As Collection
    Dim i As Integer
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If
    
If lStart > 0 And lInterval > 0 Then
        Set ObjDAL = New ClsDataAccess
    
        Set rs = New ADODB.Recordset
    
        strSQL = "Exec getfreejobs " & lLocationID
    
        Set rs = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    
        Set mCol = New Collection
        If rs.RecordCount > 0 Then
            rs.Move lStart - 1
            i = 0
            Do Until rs.EOF
                If i >= lInterval Then
                    Exit Do
                End If
                Set objJob = objContext.CreateInstance("xtraweb.Job")
                For Each fldData In rs.Fields
                    objJob.DataValues(fldData.Name) = fldData.Value
                Next
                mCol.Add objJob
                rs.MoveNext
                i = i + 1
            Loop
        End If
        Set GetJobsInInterval = mCol
    Else
        Set GetJobsInInterval = Nothing
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
End Function


Public Function GetJobLocations() As ADODB.Recordset
On Error GoTo err_GetJobLocations
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As Command
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    'Dummy parameter
    Set ObjCmd = New ADODB.Command
    Set GetJobLocations = ObjDAL.FetchRCSP(mStrXtraCon, "GetJobLocations", ObjCmd.Parameters)
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

err_GetJobLocations:
    WriteLog "Search_GetJobLocations", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetJob(ByVal lOppdragID As Long) As ADODB.Recordset
On Error GoTo err_GetJob
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As Command
    Dim ObjParam As ADODB.Parameter
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    Set ObjCmd = New ADODB.Command
    
    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "lOppdragID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = lOppdragID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjDAL = New ClsDataAccess

    Set GetJob = ObjDAL.FetchRCSP(mStrXtraCon, "GetJob", ObjCmd.Parameters)
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

err_GetJob:
    WriteLog "Search_GetJob", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub


