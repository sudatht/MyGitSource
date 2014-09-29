VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjCV As CV
Private mblnChecked As Boolean

Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property

Public Property Let Checked(ByVal blnChecked As Boolean)
    mblnChecked = blnChecked
End Property

Public Property Get Owner() As CV
    Set Owner = mobjCV
End Property

Public Property Set Owner(ByRef robjCV As CV)
    Set mobjCV = robjCV
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
    Dim lCVId As Long
    Dim lJobTypeID As Long
    Dim Lvikarid As Long
       
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    lCVId = mobjCV.DataValues("CVId").Value
    lJobTypeID = mobjDataValues("K_TittelID").Value
           
    If TypeOf mobjCV.Owner Is Consultant Then
        Lvikarid = CLng(mobjCV.Owner.DataValues("vikarid").Value)
    ElseIf TypeOf mobjCV.Owner Is Suspect Then
        Lvikarid = 0
    End If
        
    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
   
    ' check if JobType already exist on this cv
    strSQL = "SELECT COUNT(*) FROM [vikar_kompetanse] WHERE [K_TypeID] = 4 AND [cvid] = " & lCVId & " AND [K_TittelID] = " & lJobTypeID
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If rs(0) = 0 Then
        'no jobtypes exist
        strSQL = "SELECT * FROM [vikar_kompetanse] WHERE 1 = NULL"
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        'Add new
        rs.AddNew
        'Jobbønske
        rs("K_TypeID") = 4
        'Not rated
        rs("K_LevelID") = 0
        rs("vikarid") = Lvikarid
        rs("cvid") = lCVId
        rs("K_TittelID") = lJobTypeID
        ObjDAL.BatchUpdate mobjCV.XtraConString, rs
        rs.Close
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Save = True
    Exit Function

err_Save:
    Save = False
    WriteLog "JobType_Save", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set rs = Nothing
End Function

'deprecated
Public Function SaveChanged() As Boolean
On Error GoTo err_Savechanged
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim RsOriginal As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim lJobTypeID As Long
    Dim strChecked As String
    Dim StrType As String
    Dim BlnSave As Boolean
    Dim BlnUpdate As Boolean
    Dim val As DataValue
        
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    lCVId = mobjCV.DataValues("CVId").Value
    lJobTypeID = mobjDataValues("k_TittelID").Value
    strChecked = mobjDataValues("Checked").Value
        
    BlnSave = False
    BlnUpdate = False
    
    Set ObjDAL = New ClsDataAccess
    ' check if JobType already exist on this cv
    strSQL = "Select * from CV_vikar_jobwishes_temp where cvid = " & lCVId & " and WishID = " & lJobTypeID

    Set RsOriginal = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    'If jobwish is checked -
    If strChecked = "CHECKED" Then
        'If the jobwish doesn't exist
        If RsOriginal.EOF Then
            'Save as new
            StrType = "NY"
            BlnSave = True
        Else
            If RsOriginal.Fields("type").Value = "SLETTET" Then
                StrType = "ORIGINAL"
                BlnUpdate = True
            End If
        End If
    Else
        'If not checked and the jobwish exists in the original CV
        If Not RsOriginal.EOF Then
            ' Mark as ready to be deleted
            If RsOriginal.Fields("Original_value") Then
                StrType = "SLETTET"
                BlnUpdate = True
            'If not checked and jobwish is not part of original CV, delete at once..
            Else
                strSQL = "delete from CV_vikar_jobwishes_temp where cvid =" & lCVId & " and wishid=" & lJobTypeID
                If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
                    Err.Raise 701 + vbObjectError, "", "Couldn't delete jobwish! CVid:" & lCVId & ", Jobwish:" & lJobTypeID
                End If
            End If
        Else
            SaveChanged = False
            Exit Function
        End If
    End If
    If BlnSave Then
        strSQL = "SELECT * FROM CV_vikar_jobwishes_temp where 1 = null"
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        rs.AddNew
        rs("cvid") = lCVId
        rs("wishid") = lJobTypeID
        rs("type") = StrType
        ObjDAL.BatchUpdate mobjCV.XtraConString, rs
        rs.Close
    End If
    If BlnUpdate Then
        RsOriginal("type") = StrType
        ObjDAL.BatchUpdate mobjCV.XtraConString, RsOriginal
        RsOriginal.Close
    End If
    Set rs = Nothing
    Set RsOriginal = Nothing
    Set ObjDAL = Nothing
    'Make sure CV is marked as changed
    If mobjCV.DataValues("cv_changed").Value = 0 Then
        mobjCV.SaveChanged
    End If
    SaveChanged = True
    Exit Function
    
err_Savechanged:
    SaveChanged = False
    WriteLog "JobType_SaveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set RsOriginal = Nothing
    Set ObjDAL = Nothing
End Function


Public Function Delete() As Boolean
On Error GoTo err_Delete
    Dim ObjDAL As IdataAccess
    Dim rsJobWish As ADODB.Recordset
    Dim strSQL As String
    
    Dim lCVId As Long
    Dim lJobTypeID As Long
    
    Set ObjDAL = New ClsDataAccess
    Set rsJobWish = New ADODB.Recordset
    
    lCVId = mobjCV.DataValues("CVId").Value
    lJobTypeID = mobjDataValues("K_TittelID").Value
    
    strSQL = "SELECT K_TittelID FROM [vikar_kompetanse] WHERE [cvid] = " & lCVId & " AND [K_TypeID] = 4 AND [K_TittelID] = " & lJobTypeID
    Set rsJobWish = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If Not rsJobWish.EOF Then
        rsJobWish.Delete
        ObjDAL.BatchUpdate mobjCV.XtraConString, rsJobWish
    End If
    Set rsJobWish = Nothing
    Delete = True
    Exit Function

err_Delete:
    Delete = False
    WriteLog "JobType_Delete", Error$, EF_LogEventType_Error
    Set rsJobWish = Nothing
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
    Set mobjCV = Nothing
    Set objContext = Nothing
End Sub
