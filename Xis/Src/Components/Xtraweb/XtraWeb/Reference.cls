VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Reference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjOwner As CV

Public Property Set Owner(ByVal vData As CV)
    Set mobjOwner = vData
    'Set mobjDataValues.Owner = SafeRef(Me)
End Property

Public Property Get Owner() As CV
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Function Save() As Boolean
On Error GoTo err_Save
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim fldData As ADODB.Field
    Dim objValue As DataValue
    Dim lCVId As Long
    Dim lRefID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = mobjOwner.DataValues("CVId").Value
    
    ' check if JobType already exist on this cv
    
    If Not ExistsInDataValues(mobjDataValues, "ReferenceID") Then
        'New
        Trace "inside if"
        strSQL = "Select * from cv_references where 1 = null"
        Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
        Trace "record count : " & rs.RecordCount
        rs.AddNew
        rs("cvid") = lCVId
        Trace "lCVId :" & lCVId
        For Each objValue In mobjDataValues
            Trace " objValue.ValueName : " & objValue.ValueName
            Trace " objValue.Value : " & objValue.Value
            rs(objValue.ValueName) = objValue.Value
        Next
    Else
        'change
        Trace "inside else"
        
        lRefID = mobjDataValues("ReferenceID").Value
        
        Trace "lCVId :" & lCVId
        Trace "ReferenceID :" & lRefID
        
        
        strSQL = "Select * from cv_references where cvid = " & lCVId & " and ReferenceID = " & lRefID
        Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
        Trace "record count : " & rs.RecordCount
        
        For Each objValue In mobjDataValues
            Trace "inside for"
            If LCase(objValue.ValueName) <> "referenceid" Then
                Trace " objValue.ValueName : " & objValue.ValueName
                Trace " objValue.Value : " & objValue.Value
                rs(objValue.ValueName) = objValue.Value
            End If
        Next
    End If
    
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Set rs = Nothing
    Set ObjDAL = Nothing
    Save = True
    Exit Function
    
err_Save:
    Save = False
    WriteLog "Reference_Save", Error$, EF_LogEventType_Error
End Function


Public Function Delete() As Boolean

Trace "inside Reference.Delete()"
On Error GoTo err_Delete
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim lRefID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = mobjOwner.DataValues("CVId").Value
    lRefID = mobjDataValues("ReferenceID").Value
    
    strSQL = "Select * from cv_references where cvid = " & lCVId & " and ReferenceID = " & lRefID
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If Not rs.EOF Then
        rs.Delete
        ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
        rs.Close
    End If
    Set ObjDAL = Nothing
    Delete = True
    Exit Function

err_Delete:
    Delete = False
    WriteLog "Reference_Delete", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Public Function SaveChanged() As Boolean
On Error GoTo err_Savechanged
    
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim RsOriginal As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim strSQL As String
    Dim LreferenceID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
       
    If ExistsInDataValues(mobjDataValues, "ReferenceID") Then
        LreferenceID = mobjDataValues("ReferenceID").Value
    End If
    strSQL = "Select * from cv_vikar_references_temp where ReferenceID = " & LreferenceID
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "referenceid" Then
                rs.Fields(objValue.ValueName).Value = objValue.Value
            End If
        End If
    Next
    rs("CVId") = mobjOwner.DataValues("CvID").Value
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Set rs = Nothing
    Set ObjDAL = Nothing
    'Make sure CV is marked as changed
    If mobjOwner.DataValues("cv_changed").Value = 0 Then
        mobjOwner.SaveChanged
    End If
    SaveChanged = True
    Exit Function

err_Savechanged:
    SaveChanged = False
    WriteLog "Reference_SaveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set ObjDAL = Nothing
End Function


Public Function DeleteChanged() As Boolean
On Error GoTo err_DeleteChanged
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strColName As String
    Dim lCVId As Long
    Dim LreferenceID As Long
        
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
        
    lCVId = mobjOwner.DataValues("CVId").Value
    LreferenceID = mobjDataValues("ReferenceID").Value
    
    strSQL = "Select * from cv_vikar_references_temp where ReferenceID = " & LreferenceID & " and CVId = " & lCVId
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    If Not rs.EOF Then
        If ((rs.Fields("OriginaldataID").Value = "") Or (IsNull(rs.Fields("OriginaldataID").Value))) Then
            rs.Delete
        Else
            rs.Fields("Type").Value = "SLETTET"
        End If
        ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
        rs.Close
    End If
    Set ObjDAL = Nothing
    DeleteChanged = True
    Exit Function
    
err_DeleteChanged:
    DeleteChanged = False
    WriteLog "Reference_DeleteChanged", Error$, EF_LogEventType_Error
End Function


