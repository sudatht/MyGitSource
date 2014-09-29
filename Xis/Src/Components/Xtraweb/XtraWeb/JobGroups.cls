VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjOwner As CV

Public Property Get Owner() As CV
    Set Owner = mobjOwner
End Property

Public Property Set Owner(ByRef robjCV As CV)
    Set mobjOwner = robjCV
End Property

Public Function Add(ByVal objJobGroup As JobGroup, Optional sKey As String) As JobGroup
    
    If Len(sKey) = 0 Then
        mCol.Add objJobGroup
    Else
        mCol.Add objJobGroup, sKey
    End If
    'return the object created
    Set Add = objJobGroup
    Set objJobGroup = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As JobGroup
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Refresh()
On Error GoTo err_Refresh
    Dim ObjDAL          As IdataAccess
    Dim rsMain          As ADODB.Recordset
    Dim fldData         As ADODB.Field
    Dim strSQL          As String
    Dim objJobGroup     As JobGroup
    Dim strID           As String
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    If Trim(Len(mobjOwner.XtraDataShapeConString)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " CV has no Datashape connection string (Xtra)!"
    End If

    Set mCol = New Collection
    Set ObjDAL = New ClsDataAccess
        
    'strSQL = "SELECT DISTINCT [t1].* FROM [tjenesteomrade] AS [t1], [h_komp_tittel] AS [t2] WHERE [t1].[tomid] = [t2].[tomid]"
    strSQL = "SELECT DISTINCT [t1].[FagID], [t1].[fagOmrade] FROM [H_KOMP_FAGOMRADE] AS [t1], [h_komp_tittel] AS [t2] WHERE [t1].[FagID] = [t2].[fagID] ORDER BY [t1].[fagOmrade]"
    
    Set rsMain = ObjDAL.FetchRC(mobjOwner.XtraDataShapeConString, strSQL)
    Set ObjDAL = Nothing
    While Not rsMain.EOF
        Set objJobGroup = objContext.CreateInstance("xtraweb.JobGroup")
        For Each fldData In rsMain.Fields
            objJobGroup.DataValues(fldData.Name) = fldData.Value
        Next
        strID = "ID" & objJobGroup.DataValues("FagID").Value
        Add objJobGroup, strID
        rsMain.MoveNext
    Wend
    Set rsMain = Nothing
    Exit Sub
    
err_Refresh:
    WriteLog "JobGroups_Refresh", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set rsMain = Nothing
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjOwner = Nothing
End Sub

Public Sub RefreshChanged()
    Call Refresh
End Sub

Public Function ApproveChanged(ByVal LWishID As Long) As Boolean
On Error GoTo err_ApproveChanged
    
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim StrType As String
    Dim LWishOrgId As Long
    Dim lCVId As Long
    Dim Lvikarid As Long
    Dim LOriginalValue As Long
    Dim strSQL As String
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter

    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
                                            
    If IsNull(mobjOwner.DataValues("cvid").Value) Then
        ApproveChanged = False
        Exit Function
    Else
        lCVId = mobjOwner.DataValues("cvid").Value
    End If

    If val(mobjOwner.Owner.DataValues("vikarid").Value) < 1 Then
       Err.Raise 701 + vbObjectError, "", " Missing or invalid vikarid!"
    Else
        Lvikarid = mobjOwner.Owner.DataValues("vikarid").Value
    End If

    Set ObjDAL = New ClsDataAccess
    strSQL = "Select * from CV_vikar_jobwishes_temp where wishID = " & LWishID & " and cvid=" & lCVId
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        ApproveChanged = False
        Set ObjDAL = Nothing
        Exit Function
    End If
    
    StrType = rs.Fields("Type").Value
    
    If IsNull(rs.Fields("Original_value").Value) Then
        LOriginalValue = 0
    Else
        LOriginalValue = rs.Fields("Original_value").Value
    End If
    
    Set rs = Nothing
   
    If (LCase(StrType) = "ny") And (LOriginalValue = 0) Then
        Set ObjCmd = New ADODB.Command

        'Add input parameters to command object
        Set ObjParam = New Parameter
        ObjParam.Name = "Lcvid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = lCVId
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                           
        Set ObjParam = New Parameter
        ObjParam.Name = "LWishid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LWishID
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        Set ObjParam = New Parameter
        ObjParam.Name = "Lvikarid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = Lvikarid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        'Oppdaterer hovedtabell for CV data
        If Not ObjDAL.ExecuteCRUDSP(mobjOwner.XtraConString, "InsertWishMain", ObjCmd.Parameters) Then
            Err.Raise 701 + vbObjectError, "", "Could not insert wishes into main data! "
        End If
        Set ObjCmd = Nothing
        Set ObjDAL = Nothing
               
    ElseIf (LCase(StrType) = "slettet") And (LOriginalValue <> 0) Then
        Set ObjCmd = New ADODB.Command
        
        'Add input parameters to command object
        Set ObjParam = New Parameter
        ObjParam.Name = "Lcvid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = lCVId
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                           
        Set ObjParam = New Parameter
        ObjParam.Name = "LWishid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LWishID
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                                                                                  
        Set ObjParam = New Parameter
        ObjParam.Name = "Lvikarid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = Lvikarid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        'Oppdaterer hovedtabell for CV data
        If Not ObjDAL.ExecuteCRUDSP(mobjOwner.XtraConString, "DeleteTempAndOrgWish", ObjCmd.Parameters) Then
            Err.Raise 701 + vbObjectError, "", "Could not Delete wishes from tables! Procedure: DeleteTempAndOrgWish."
        End If
        Set ObjCmd = Nothing
        Set ObjDAL = Nothing
    Else
        ApproveChanged = False
        Exit Function
    End If
    ApproveChanged = True
    Exit Function

err_ApproveChanged:
    ApproveChanged = False
    WriteLog "JobGroups_ApproveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function


Public Function RejectChanges(ByVal LWishID As Long) As Boolean
On Error GoTo err_RejectChanges
    
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim StrType As String
    Dim lCVId As Long
    Dim LOriginalValue As Long
    Dim strSQL As String

    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
        
    Set ObjDAL = New ClsDataAccess
                                                                                 
    If IsNull(mobjOwner.DataValues("cvid").Value) Then
        RejectChanges = False
        Exit Function
    Else
        lCVId = mobjOwner.DataValues("cvid").Value
    End If
    
    strSQL = "Select * from CV_vikar_jobwishes_temp where wishID = " & LWishID & " and cvid=" & lCVId
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Set ObjDAL = Nothing
        RejectChanges = False
        Exit Function
    End If
    
    StrType = rs.Fields("Type").Value
    
    If IsNull(rs.Fields("Original_value").Value) Then
        LOriginalValue = 0
    Else
        LOriginalValue = rs.Fields("Original_value").Value
    End If
                
    Set rs = Nothing
    
    If (LCase(StrType) = "ny") And (LOriginalValue = 0) Then
        strSQL = "Delete from CV_vikar_jobwishes_temp where wishID = " & LWishID & " and cvid=" & lCVId
        'Sletter forekomst fra temp tabellen
        If Not ObjDAL.ExecuteCRUD(mobjOwner.XtraConString, strSQL) Then
            Err.Raise 701 + vbObjectError, "", "Could delete wishes from temp data! "
        End If
    ElseIf (LCase(StrType) = "slettet") And (LOriginalValue <> 0) Then
        strSQL = "update CV_vikar_jobwishes_temp set Type='ORIGINAL' where wishID = " & LWishID & " and cvid=" & lCVId
        'Sletter forekomst fra temp tabellen
        If Not ObjDAL.ExecuteCRUD(mobjOwner.XtraConString, strSQL) Then
            Err.Raise 701 + vbObjectError, "", "Could delete wishes from temp data! "
        End If
    Else
        RejectChanges = False
        Set ObjDAL = Nothing
        Exit Function
    End If
    Set ObjDAL = Nothing
    RejectChanges = True
    Exit Function

err_RejectChanges:
    RejectChanges = False
    WriteLog "JobGroups_RejectChanges", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
End Function


