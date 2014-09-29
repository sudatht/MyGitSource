VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Experience"
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
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lDataID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    
    Set rs = New ADODB.Recordset
   
    If ExistsInDataValues(mobjDataValues, "dataid") Then
        lDataID = mobjDataValues("dataid").Value
    End If
           
    strSQL = "Select * from cv_data where dataid = " & lDataID
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "dataid" Then
                rs(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    
    rs("CVId") = mobjOwner.DataValues("CvID").Value
    rs("FieldType") = "ARB"
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Set rs = Nothing
    Save = True
    Exit Function

err_Save:
    Save = False
    WriteLog "Experience_Save", Error$, EF_LogEventType_Error
    Set rs = Nothing
End Function

Public Function Delete() As Boolean
On Error GoTo err_Delete
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strColName As String
    Dim lCVId As Long
    Dim lDataID As Long
        
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = mobjDataValues("CVId").Value
    lDataID = mobjDataValues("dataid").Value
        
    strSQL = "Select * from cv_data where DataId = " & lDataID & " and CVId = " & lCVId
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    If Not rs.EOF Then
        rs.Delete
        ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
        rs.Close
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Delete = True
    Exit Function

err_Delete:
    Delete = False
    WriteLog "Experience_Delete", Error$, EF_LogEventType_Error
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
    Dim lDataID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
       
    Set ObjDAL = New ClsDataAccess
    
    Set rs = New ADODB.Recordset
   
    If ExistsInDataValues(mobjDataValues, "dataid") Then
        lDataID = mobjDataValues("dataid").Value
    End If
              
    strSQL = "Select * from CV_vikar_Data_temp where dataid = " & lDataID
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "dataid" Then
                rs(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    rs("CVId") = mobjOwner.DataValues("CvID").Value
    rs("FieldType") = "ARB"
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Set rs = Nothing
    'Make sure CV is marked as changed
    If mobjOwner.DataValues("cv_changed").Value = 0 Then
        mobjOwner.SaveChanged
    End If
    SaveChanged = True
    Exit Function

err_Savechanged:
    SaveChanged = False
    WriteLog "Experience_SaveChanged", Error$, EF_LogEventType_Error
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
    Dim lDataID As Long
            
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = mobjOwner.DataValues("CVId").Value
    lDataID = mobjDataValues("dataid").Value
    
    strSQL = "Select * from CV_vikar_Data_temp where dataid = " & lDataID & " and CVId = " & lCVId
    
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
    Set rs = Nothing
    Set ObjDAL = Nothing
    DeleteChanged = True
    Exit Function
    
err_DeleteChanged:
    DeleteChanged = False
    WriteLog "Experience_DeleteChanged", Error$, EF_LogEventType_Error
End Function
