VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "CV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjExperiences As Experiences
Private mobjEducations As Educations
Private mobjDataValues As DataValues
Private mobjOwner As Object
Private mobjJobGroups As JobGroups
Private mobjReferences As References
Private mobjProductGroups As ProductGroups
Private mStrXtraCon As String
Private mStrXtraDataShapeCon As String
Private mStrCVType As String
Private mobjUploadFiles As UploadFiles

Public Property Get CvType() As String
    CvType = mStrCVType
End Property

Public Property Get isLocked() As Boolean
    If CInt(mobjDataValues("Locked").Value) = -1 Then
        isLocked = True
    Else
        isLocked = False
    End If
End Property

Public Function LockCV() As Boolean
    If mStrCVType = "ORIGINAL" Then
        Call Refresh
    ElseIf mStrCVType = "CHANGED" Then
        Call RefreshChanged
    End If
    mobjDataValues("Locked").Value = "1"
    Me.SaveChanged
End Function

Public Function UnLockCV() As Boolean
    If mStrCVType = "ORIGINAL" Then
        Call Refresh
    ElseIf mStrCVType = "CHANGED" Then
        Call RefreshChanged
    End If
    mobjDataValues("Locked").Value = "0"
    Me.SaveChanged
End Function

Public Property Get XtraConString() As String
    XtraConString = mStrXtraCon
End Property

Public Property Let XtraConString(strConString As String)
    mStrXtraCon = strConString
End Property

Public Property Get XtraDataShapeConString() As String
    XtraDataShapeConString = mStrXtraDataShapeCon
End Property

Public Property Let XtraDataShapeConString(strConString As String)
    mStrXtraDataShapeCon = strConString
End Property

Public Property Set ProductGroups(ByVal vData As ProductGroups)
    Set mobjProductGroups = vData
End Property

Public Property Get ProductGroups() As ProductGroups
    Set ProductGroups = mobjProductGroups
    If mStrCVType = "ORIGINAL" Then
        mobjProductGroups.Refresh
    ElseIf mStrCVType = "CHANGED" Then
        mobjProductGroups.RefreshChanged
    End If
End Property

Public Property Set References(ByVal vData As References)
    Set mobjReferences = vData
End Property

Public Property Get References() As References

Trace "inside CV.Get References() "
    Set References = mobjReferences
    If mStrCVType = "ORIGINAL" Then
        Trace "inside ORIGINAL "
        mobjReferences.Refresh
    ElseIf mStrCVType = "CHANGED" Then
        Trace "inside CHANGED "
        mobjReferences.RefreshChanged
    End If
End Property

Public Property Set JobGroups(ByVal vData As JobGroups)
    Set mobjJobGroups = vData
End Property

Public Property Get JobGroups() As JobGroups
    Set JobGroups = mobjJobGroups
    If mStrCVType = "ORIGINAL" Then
        mobjJobGroups.Refresh
    ElseIf mStrCVType = "CHANGED" Then
        mobjJobGroups.RefreshChanged
    End If
End Property

Public Property Set Owner(ByVal vData As Object)
    Set mobjOwner = vData
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Property Get Educations() As Educations
    If mobjEducations Is Nothing Then
        Set mobjEducations = objContext.CreateInstance("xtraweb.Educations")
    End If
    If mStrCVType = "ORIGINAL" Then
        mobjEducations.Refresh
    ElseIf mStrCVType = "CHANGED" Then
        mobjEducations.RefreshChanged
    End If
    Set Educations = mobjEducations
End Property

Public Property Set Educations(vData As Educations)
    Set mobjEducations = vData
End Property

Public Property Get Experiences() As Experiences
    If mobjExperiences Is Nothing Then
        Set mobjExperiences = objContext.CreateInstance("xtraweb.Experiences")
    End If
    If mStrCVType = "ORIGINAL" Then
        mobjExperiences.Refresh
    ElseIf mStrCVType = "CHANGED" Then
        mobjExperiences.RefreshChanged
    End If
    Set Experiences = mobjExperiences
End Property

Public Property Set Experiences(vData As Experiences)
    Set mobjExperiences = vData
End Property


Public Property Set UplaodFiles(ByVal vData As UploadFiles)
    Set mobjUploadFiles = vData
End Property

Public Property Get UplaodFiles() As UploadFiles

    Trace "Enter property CV.Get UplaodFiles() "
    'mobjUploadFiles.Refresh
    Set UplaodFiles = mobjUploadFiles
    'If mStrCVType = "ORIGINAL" Then
        mobjUploadFiles.Refresh
    'ElseIf mStrCVType = "CHANGED" Then
        'mobjUploadFiles.RefreshChanged
    'End If
End Property



Public Function Save() As Boolean
On Error GoTo err_Save
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rsChanged As ADODB.Recordset
    Dim rsCVData As ADODB.Recordset
    Dim strSQL As String
    Dim strColName As String
    Dim strCVtype As String
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra) !"
    End If
    
    Set ObjDAL = New ClsDataAccess
                 
    If TypeOf mobjOwner Is Suspect Then
        strColName = "suspectid"
        strCVtype = "S"
    ElseIf TypeOf mobjOwner Is Consultant Then
        strColName = "VikarID"
        strCVtype = "C"
    End If
        
    strSQL = "Select * from cv where consultantid = " & mobjOwner.DataValues(strColName).Value & " and type ='" & strCVtype & "'"
    
    'Check if cv exists
    Set rsCVData = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    
    If rsCVData.EOF Then
        rsCVData.AddNew
    End If
        
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "cvid" Then
                rsCVData(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    
    rsCVData("ConsultantID") = mobjOwner.DataValues(strColName).Value
    rsCVData("Type") = strCVtype

    ObjDAL.BatchUpdate mStrXtraCon, rsCVData
    rsCVData.Close
    Set rsCVData = Nothing
    Set ObjDAL = Nothing
    Call Refresh
    Save = True
    Exit Function

err_Save:
    Save = False
    WriteLog "CV_Save", Error$ & " StrSQL:" & strSQL & ". StrCon:" & mStrXtraCon, EF_LogEventType_Error
    Set ObjCmd = Nothing
    Set rsCVData = Nothing
    Set rsChanged = Nothing
End Function

Public Function DeleteChanged() As Boolean
On Error GoTo err_Save
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rsChanged As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
        
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra) !"
    End If
    
    Set ObjDAL = New ClsDataAccess
        
    If TypeOf mobjOwner Is Consultant Then
        'if there are changed_cv data..
        strSQL = "Select CVid from cv where consultantid = " & mobjOwner.DataValues("VikarID").Value & " and type ='C'"
        Set rsChanged = ObjDAL.FetchRC(mStrXtraCon, strSQL)
        If Not rsChanged.EOF Then
            lCVId = rsChanged("cvid").Value

            'Update & delete values from CV data..
            strSQL = "update CV set cv_changed = 0 where CVid=" & lCVId
            If Not ObjDAL.ExecuteCRUD(mStrXtraCon, strSQL) Then
                Err.Raise 701, "", "Could not update CV values!"
            End If
            
            Set ObjCmd = New Command
            Set ObjParam = New Parameter
            ObjParam.Name = "CVid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = lCVId
            ObjCmd.Parameters.Append ObjParam
                
            Set ObjParam = Nothing

            If Not ObjDAL.ExecuteCRUDSP(mStrXtraCon, "DeleteConsultantsChangedCV", ObjCmd.Parameters) Then
                Err.Raise 702, "", "Could not delete changed CV values!"
            End If
            Set ObjCmd = Nothing
        End If
        Set rsChanged = Nothing
    End If
    Set ObjDAL = Nothing
    Exit Function
    
err_Save:
    WriteLog "CV_deleteChanged", Error$ & " StrSQL:" & strSQL & ". StrCon:" & mStrXtraCon, EF_LogEventType_Error
    Set ObjCmd = Nothing
    Set rsChanged = Nothing
End Function

Public Function SaveChanged() As Boolean
On Error GoTo err_Savechanged
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rsChanged As ADODB.Recordset
    Dim strSQL As String

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "CV_saveChanged", " No connection string (Xtra) !"
    End If
    
    Set ObjDAL = New ClsDataAccess
                         
    strSQL = "Select * from cv where consultantid = " & mobjOwner.DataValues("VikarID").Value & " and type ='C'"
    
    'Check if cv exists
    Set rsChanged = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    
    If rsChanged.EOF Then
        rsChanged.AddNew
    End If
        
    'Transfer all values
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "cvid" Then
                rsChanged(objValue.ValueName) = objValue.Value
            End If
        End If
    Next
    
    rsChanged.Fields("ConsultantID").Value = mobjOwner.DataValues("VikarID").Value
    rsChanged.Fields("Type").Value = "C"
    rsChanged.Fields("cv_changed").Value = 1

    ObjDAL.BatchUpdate mStrXtraCon, rsChanged
    rsChanged.Close
    Set rsChanged = Nothing
    Call RefreshChanged
    SaveChanged = True
    Exit Function

err_Savechanged:
    SaveChanged = False
    WriteLog "CV_SaveChanged", Error$ & " StrSQL:" & strSQL, EF_LogEventType_Error
    Set rsChanged = Nothing
End Function

Public Sub RefreshChanged()
On Error GoTo err_RefreshChanged
    ' ##BD Denne funksjonen refresher kolleksjonen
    
    Dim ObjDAL As IdataAccess
    Dim rsCV As ADODB.Recordset
    Dim strSQL As String
    Dim lConsultantID As Long
    Dim objEducation As Education
    Dim fldCol As ADODB.Field
    Dim strPropName As String
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "CV_RefreshChanged", " No connection string (Xtra) !"
    End If
        
    Set ObjDAL = New ClsDataAccess
    Set rsCV = New Recordset
    Set mobjDataValues = Nothing
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    
    'Dersom "konsulent"..
    If TypeOf mobjOwner Is Consultant Then
        lConsultantID = mobjOwner.DataValues("VikarID").Value
        'Check if consultant has changed his CV..
        strSQL = "Select * from cv where ConsultantID = " & lConsultantID & " and type = 'C'"
        Set rsCV = ObjDAL.FetchRC(mStrXtraCon, strSQL)
        If Not rsCV.EOF Then
            If Not rsCV.EOF Then
                For Each fldCol In rsCV.Fields
                    mobjDataValues.Add fldCol.Name, fldCol.Value
                Next
            End If
            Set rsCV = Nothing
            Exit Sub
        End If
        Set rsCV = Nothing
    ElseIf TypeOf mobjOwner Is Suspect Then
        Err.Raise 701, "", " Wrong owner type!"
    End If
    Set rsCV = Nothing
    Set ObjDAL = Nothing
    mStrCVType = "CHANGED"
    Exit Sub
    
err_RefreshChanged:
    WriteLog "cv_RefreshChanged", Error$, EF_LogEventType_Error
    Set rsCV = Nothing
    Set mobjDataValues = Nothing
End Sub

Public Sub Refresh()
On Error GoTo err_Refresh
    
    Dim ObjDAL As IdataAccess
    Dim rsCV As ADODB.Recordset
    Dim strSQL As String
    Dim lConsultantID As Long
    Dim fldCol As ADODB.Field
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "CV_Refresh", " No connection string (Xtra) !"
    End If
        
    Set ObjDAL = New ClsDataAccess
    Set rsCV = New Recordset
    Set mobjDataValues = Nothing
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    
    'Dersom "konsulent"..
    If TypeOf mobjOwner Is Consultant Then
        lConsultantID = mobjOwner.DataValues("VikarID").Value
        strSQL = "Select * from CV where ConsultantID = " & lConsultantID & " and type = 'C'"
        Set rsCV = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    'Dersom "Søker fra web"..
    ElseIf TypeOf mobjOwner Is Suspect Then
        lConsultantID = mobjOwner.DataValues("Suspectid").Value
        strSQL = "Select * from cv where ConsultantID = " & lConsultantID & " and type = 'S'"
        Set rsCV = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    End If
    Set ObjDAL = Nothing
    If Not rsCV.EOF Then
        For Each fldCol In rsCV.Fields
            mobjDataValues.Add fldCol.Name, fldCol.Value
        Next
    End If
    Set rsCV = Nothing
    mStrCVType = "ORIGINAL"
    Exit Sub
err_Refresh:
    WriteLog "CV_Refresh", Error$, EF_LogEventType_Error
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjEducations = objContext.CreateInstance("xtraweb.Educations")
    Set mobjExperiences = objContext.CreateInstance("xtraweb.Experiences")
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjJobGroups = objContext.CreateInstance("xtraweb.JobGroups")
    Set mobjReferences = objContext.CreateInstance("xtraweb.References")
    Set mobjProductGroups = objContext.CreateInstance("xtraweb.ProductGroups")
    Set mobjUploadFiles = objContext.CreateInstance("xtraweb.UploadFiles")
    Set mobjProductGroups.Owner = SafeRef(Me)
    Set mobjJobGroups.Owner = SafeRef(Me)
    Set mobjEducations.Owner = SafeRef(Me)
    Set mobjExperiences.Owner = SafeRef(Me)
    Set mobjReferences.Owner = SafeRef(Me)
    Set mobjUploadFiles.Owner = SafeRef(Me)
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjProductGroups.Owner = Nothing
    Set mobjJobGroups.Owner = Nothing
    Set mobjEducations.Owner = Nothing
    Set mobjExperiences.Owner = Nothing
    Set mobjReferences.Owner = Nothing
    Set mobjUploadFiles.Owner = Nothing
    Set mobjEducations = Nothing
    Set mobjExperiences = Nothing
    Set mobjDataValues = Nothing
    Set mobjJobGroups = Nothing
    Set mobjReferences = Nothing
    Set mobjProductGroups = Nothing
    Set mobjUploadFiles = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Public Sub CleanUp()
    Set mobjProductGroups.Owner = Nothing
    Set mobjJobGroups.Owner = Nothing
    Set mobjEducations.Owner = Nothing
    Set mobjExperiences.Owner = Nothing
    Set mobjReferences.Owner = Nothing
    Set mobjUploadFiles.Owner = Nothing
    Set mobjEducations = Nothing
    Set mobjExperiences = Nothing
    Set mobjDataValues = Nothing
    Set mobjJobGroups = Nothing
    Set mobjReferences = Nothing
    Set mobjProductGroups = Nothing
    Set mobjUploadFiles = Nothing
    Set mobjOwner = Nothing
End Sub

