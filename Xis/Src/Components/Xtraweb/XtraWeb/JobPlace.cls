VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobPlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjOwner As JobApplication
Private mobjDataValues As DataValues

Public Function Save() As Boolean
On Error GoTo err_Save

    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lJobPlaceID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
     'Trace "JobPlaceID :" & JobPlaceID
    
    If ExistsInDataValues(mobjDataValues, "JobPlaceID") Then
        'new
        lJobPlaceID = mobjDataValues("JobPlaceID").Value
        'Trace "inside if ExistsInDataValues"
        'Trace "JobPlaceID :" & lJobPlaceID
    End If
    strSQL = "SELECT * FROM [JobPlace] WHERE [JobPlaceID] = " & lJobPlaceID
    
    'Trace "sql:" & strSQL
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    'Trace "rs.RecordCount : " & rs.RecordCount
    
    If rs.RecordCount = 0 Then
        'new
        'Trace "inside rs.RecordCount : "
        rs.AddNew
    End If
    
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "JobPlaceID" Then
                'Trace "objValue.Value :" & objValue.Value
                rs(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    rs("JobApplicationID") = mobjOwner.DataValues("JobApplicationID").Value
    'Trace "JobApplicationID" + mobjOwner.DataValues("JobApplicationID").Value
    
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Call SaveJobPlaceEx
    Save = True
    Exit Function
err_Save:
    Save = False
    WriteLog "JobPlace_Save", Error$, EF_LogEventType_Error
    
End Function

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Property Set Owner(ByVal vData As JobApplication)
    Set mobjOwner = vData
    'Set mobjDataValues.Owner = SafeRef(Me)
End Property

Public Property Get Owner() As JobApplication
    Set Owner = mobjOwner
End Property

Private Function SaveJobPlaceEx() As Boolean
    'Saves to table v_suspect_avdkontor, this is
    'necessary to show the suspect in the internal system
    
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lSuspID As Long
    
    lSuspID = mobjOwner.Owner.DataValues("Suspectid").Value
    
    strSQL = "SELECT * FROM [v_suspect_avdkontor] WHERE [suspectid] = " & lSuspID
    Set ObjDAL = New ClsDataAccess
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    If rs.RecordCount = 0 Then
        'new
        rs.AddNew
    End If
    
    rs("avdkontor") = mobjDataValues("placename").Value
    rs("suspectid") = lSuspID
    rs("overfoert") = 0
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs

End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    'destroys collection when this class is terminated
    'Set mobjDataValues.Owner = Nothing
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub


