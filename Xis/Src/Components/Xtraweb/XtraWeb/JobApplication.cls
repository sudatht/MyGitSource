VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mStrXtraCon As String
Private mobjDataValues As DataValues
Private mobjOwner As Object
Private mobjJobPlaces As JobPlaces

Public Property Get XtraConString() As String
    XtraConString = mStrXtraCon
End Property

Public Property Let XtraConString(strConString As String)
    mStrXtraCon = strConString
End Property

Public Property Set JobPlaces(ByVal vData As JobPlaces)
    Set mobjJobPlaces = vData
End Property

Public Property Get JobPlaces() As JobPlaces
    Set JobPlaces = mobjJobPlaces
    If Not mobjJobPlaces.Owner Is Nothing Then
        mobjJobPlaces.Refresh
    End If
End Property

Public Property Set Owner(ByVal vData As Object)
    Set mobjOwner = vData
    'If Not vData Is Nothing Then
    '    Set mobjDataValues.Owner = SafeRef(Me)
    'End If
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Dim objValue As XtraWeb.DataValue
    Trace "Set DataValues 1"
    Set mobjDataValues = vData
    Trace "Set DataValues 2"
    For Each objValue In mobjDataValues
        Trace "name : " & objValue.ValueName & "value : " & objValue.Value
    Next
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
    'Refresh
End Property

Public Function Save() As Boolean
On Error GoTo err_Save

    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
       Trace "Error"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
        
    strSQL = "Select * from JobApplication where 1 = null"
    Trace "strSQL" & strSQL
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    If Not ExistsInDataValues(mobjDataValues, "JobApplicationID") Then
        'new
        rs.AddNew
        Trace "New"
    Else
        strSQL = "Select * from jobapplication where JobApplicationID = " & mobjDataValues("JobApplicationID").Value
        Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
        Trace "Else " & mobjDataValues("JobApplicationID").Value
    End If
    
    For Each objValue In mobjDataValues
        Trace "name : " & objValue.ValueName & "value : " & objValue.Value
        If Not IsNull(objValue.Value) Then
            Trace "1"
            If LCase(objValue.ValueName) <> "jobapplicationid" Then
                Trace "2"
                If LCase(objValue.ValueName) <> "suspectid" Then
                    rs(objValue.ValueName) = objValue.Value
                    Trace objValue.Value
                End If
            End If
        End If
    Next
    
    Trace mobjOwner.DataValues("SuspectID").Value
    rs("SuspectID") = mobjOwner.DataValues("SuspectID").Value
    
    ObjDAL.BatchUpdate mobjOwner.XtraConString, rs
    rs.Close
    Set rs = Nothing
    Set ObjDAL = Nothing
    Save = True
    Exit Function
err_Save:
    Save = False
    WriteLog "JobApplication_Save", Error$, EF_LogEventType_Error
End Function

Public Sub Refresh()
On Error GoTo err_Refresh

    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim fldData As ADODB.Field
    Dim strSQL As String
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    strSQL = "Select * from JobApplication where suspectid = " & mobjOwner.DataValues("suspectid").Value
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If Not rs.EOF Then
        For Each fldData In rs.Fields
            mobjDataValues(fldData.Name) = fldData.Value
        Next
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Exit Sub
err_Refresh:
    WriteLog "JobApplication_Refresh", Error$, EF_LogEventType_Error
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjJobPlaces = objContext.CreateInstance("xtraweb.JobPlaces")
    Set mobjJobPlaces.Owner = SafeRef(Me)
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Sub CleanUp()
    Set mobjDataValues = Nothing
    Set mobjJobPlaces.Owner = Nothing
    Set mobjJobPlaces = Nothing
    Set mobjOwner = Nothing
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjJobPlaces.Owner = Nothing
    Set mobjJobPlaces = Nothing
    Set mobjOwner = Nothing
End Sub

