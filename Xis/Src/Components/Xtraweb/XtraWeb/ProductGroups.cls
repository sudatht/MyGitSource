VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "ProductGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjCV As CV

Public Property Get Owner() As CV
    Set Owner = mobjCV
End Property

Public Property Set Owner(ByRef objCV As CV)
    Set mobjCV = objCV
End Property

Public Function Add(ByVal objProductGroup As ProductGroup, Optional sKey As String) As ProductGroup
        
    If Len(sKey) = 0 Then
        mCol.Add objProductGroup
    Else
        mCol.Add objProductGroup, sKey
    End If

    'return the object created
    Set Add = objProductGroup
    Set objProductGroup = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ProductGroup
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

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjCV = Nothing
End Sub

Public Sub Refresh()
On Error GoTo err_Refresh
    Dim ObjDAL              As IdataAccess
    Dim rsMain              As ADODB.Recordset
    Dim fldData             As ADODB.Field
    Dim strSQL              As String
    Dim objProductGroup     As ProductGroup
    Dim objProduct          As Product
    Dim strID               As String
    Dim i                   As Integer
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    
    strSQL = "SELECT * FROM [H_KOMP_AREA]"
    
    Set rsMain = ObjDAL.FetchRC(mobjCV.XtraDataShapeConString, strSQL)
    Set ObjDAL = Nothing
    Set mCol = New Collection
    
    While Not rsMain.EOF
        Set objProductGroup = objContext.CreateInstance("xtraweb.ProductGroup")
        With objProductGroup
            For Each fldData In rsMain.Fields
                objProductGroup.DataValues(fldData.Name) = fldData.Value
            Next
            strID = "ID" & .DataValues("ProdOmradeID").Value
        End With
        Add objProductGroup, strID
        rsMain.MoveNext
    Wend
    'close & Clean up
    Set rsMain = Nothing
    Exit Sub

err_Refresh:
    WriteLog "ProductGroups_Refresh", Error$, EF_LogEventType_Error
    Set rsMain = Nothing
    Set ObjDAL = Nothing
End Sub

Public Sub RefreshChanged()
    Call Refresh
End Sub

Public Function RejectChanges(ByVal LProdID As Long) As Boolean
On Error GoTo err_RejectChanges
    
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim StrType As String
    Dim lCVId As Long
    Dim LOriginalValue As Long
    Dim strSQL As String

    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
                                                                                
    If IsNull(mobjCV.DataValues("cvid").Value) Then
        RejectChanges = False
        Exit Function
    Else
        lCVId = mobjCV.DataValues("cvid").Value
    End If
    
    strSQL = "Select * from CV_vikar_productknowledge_temp where ProdID = " & LProdID & " and cvid=" & lCVId
    Set ObjDAL = New ClsDataAccess
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
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
        strSQL = "Delete from CV_vikar_productknowledge_temp where ProdID = " & LProdID & " and cvid=" & lCVId
        'Sletter forekomst fra temp tabellen
        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
            Err.Raise 701 + vbObjectError, "", "Could not delete Products from temp data! StrSQL:" & strSQL
        End If
    ElseIf (LCase(StrType) = "slettet") And (LOriginalValue <> 0) Then
        strSQL = "update CV_vikar_productknowledge_temp set Type='ORIGINAL' where ProdID = " & LProdID & " and cvid=" & lCVId
        'Sletter forekomst fra temp tabellen
        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
            Err.Raise 701 + vbObjectError, "", "Could not delete Product from temp data! StrSQL:" & strSQL
        End If
    Else
        Set ObjDAL = Nothing
        RejectChanges = False
        Exit Function
    End If
    Set ObjDAL = Nothing
    RejectChanges = True
    Exit Function

err_RejectChanges:
    RejectChanges = False
    WriteLog "ProductGroups_RejectChanges", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
End Function

Public Function ApproveChanged(ByVal LProdID As Long) As Boolean
On Error GoTo err_ApproveChanged
    
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim StrType As String
    Dim lCVId As Long
    Dim LOriginalValue As Long
    Dim strSQL As String
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long

    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    If val(mobjCV.Owner.DataValues("vikarid").Value) < 1 Then
       Err.Raise 701 + vbObjectError, "", " Missing or invalid vikarid!"
    Else
        Lvikarid = mobjCV.Owner.DataValues("vikarid").Value
    End If
                                                                                  
    If IsNull(mobjCV.DataValues("cvid").Value) Then
        ApproveChanged = False
        Exit Function
    Else
        lCVId = mobjCV.DataValues("cvid").Value
    End If

    Set ObjDAL = New ClsDataAccess
       
    strSQL = "Select * from CV_vikar_productknowledge_temp where ProdID = " & LProdID & " and cvid=" & lCVId
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
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
        ObjParam.Name = "LProdID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LProdID
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
        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "InsertProductMain", ObjCmd.Parameters) Then
            Err.Raise 701 + vbObjectError, "", "Could not insert Product into main data! "
        End If
        Set ObjCmd = Nothing
               
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
        ObjParam.Name = "LProdID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LProdID
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
        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "DeleteTempAndOrgProduct", ObjCmd.Parameters) Then
            Err.Raise 701 + vbObjectError, "", "Could not Delete Products from tables! Procedure: DeleteTempAndOrgProduct."
        End If
        Set ObjCmd = Nothing
    Else
        ApproveChanged = False
        'Clean up
        Set ObjDAL = Nothing
        Exit Function
    End If
    ApproveChanged = True
    'Clean up
    Set ObjDAL = Nothing
    Exit Function

err_ApproveChanged:
    ApproveChanged = False
    WriteLog "JobGroups_ApproveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

