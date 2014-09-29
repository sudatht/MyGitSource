VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Educations"
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

Public Property Set Owner(ByRef robjCV As CV)
    Set mobjCV = robjCV
End Property

Public Function Add(ByVal objEducation As Education) As Education
    Dim strKey As String
    strKey = "ID" & Trim(objEducation.DataValues("dataid").Value)
    mCol.Add objEducation, strKey
    'return the object created
    Set Add = objEducation
    Set objEducation = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Education
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
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim objEducation As Education
    Dim fldCol As ADODB.Field
        
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    lCVId = mobjCV.DataValues("CVid").Value
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New Recordset
    
    strSQL = "Select * from cv_data where cvid = " & lCVId & " and FieldType = 'UTD' order by FromYear desc, FromMonth desc"
    
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    
    Set mCol = New Collection
    
    While Not rs.EOF
        Set objEducation = objContext.CreateInstance("xtraweb.Education")
        Set objEducation.Owner = mobjCV
        For Each fldCol In rs.Fields
            objEducation.DataValues(fldCol.Name) = fldCol.Value
        Next
        Add objEducation
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set ObjDAL = Nothing
    Exit Sub
err_Refresh:
    WriteLog "Educations_Refresh", Error$, EF_LogEventType_Error
End Sub

Public Sub RefreshChanged()
On Error GoTo err_Refresh
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim objEducation As Education
    Dim fldCol As ADODB.Field
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
        
    lCVId = mobjCV.DataValues("CVid").Value
    
    Set ObjDAL = New ClsDataAccess
    
    strSQL = "Select * from CV_vikar_Data_temp where FieldType = 'UTD' and cvid = " & lCVId & " order by FromYear desc, FromMonth desc"
    
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If rs.EOF Then
        strSQL = "Select CVId, Place, Country, Description, Title, FieldType, FromMonth, FromYear, ToMonth, ToYear, Type = 'ORIGINAL',  OriginaldataID = dataid from CV_Data where cvid = " & lCVId & " and FieldType = 'UTD' order by FromYear desc, FromMonth desc"
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        If Not rs.EOF Then
            While Not rs.EOF
                With rs
                    strSQL = "Insert into CV_vikar_Data_temp(CVId, Place, Country, Description, Title, FieldType, FromMonth, FromYear, ToMonth, ToYear, Type, OriginaldataID) values(" & .Fields("cvid").Value & ",'" & .Fields("place").Value & "','" & .Fields("country").Value & "','" & .Fields("description").Value & "','" & .Fields("Title").Value & "','UTD','" & .Fields("FromMonth").Value & "','" & .Fields("FromYear").Value & "','" & .Fields("ToMonth").Value & "','" & .Fields("ToYear").Value & "','" & .Fields("Type").Value & "','" & .Fields("OriginaldataID").Value & "')"
                    If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
                        Err.Raise 701 + vbObjectError, "", "Error during insert! (Educations) StrSQL:" & strSQL
                    End If
                End With
                rs.MoveNext
            Wend
            strSQL = "Select * from CV_vikar_Data_temp where cvid = " & lCVId & " and FieldType = 'UTD' order by FromYear desc, FromMonth desc"
            Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        End If
    End If
    If Not rs.EOF Then
        Set mCol = New Collection
        While Not rs.EOF
            Set objEducation = objContext.CreateInstance("xtraweb.Education")
            Set objEducation.Owner = mobjCV
            For Each fldCol In rs.Fields
                objEducation.DataValues(fldCol.Name) = fldCol.Value
            Next
            Add objEducation
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Exit Sub
err_Refresh:
    WriteLog "Educations_RefreshChanged", Error$, EF_LogEventType_Error
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
    Set mobjCV = Nothing
    Set objContext = Nothing
End Sub

Public Function ApproveChanged(ByVal lDataID As Long) As Boolean
On Error GoTo err_ApproveChanged
    
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim RsOriginal As ADODB.Recordset
    Dim strSQL As String
    Dim StrType As String
    Dim LOrgId As Long
    Dim lCVId As Long
    Dim LtempInt As Long
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter

    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
                                         
    strSQL = "Select * from CV_vikar_Data_temp where fieldType='UTD' and dataid = " & lDataID
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        ApproveChanged = False
        Set ObjDAL = Nothing
        Exit Function
    End If
    
    StrType = rs.Fields("Type").Value
    If IsNull(rs.Fields("OriginaldataID").Value) Then
        LOrgId = 0
    Else
        LOrgId = rs.Fields("OriginaldataID").Value
    End If
    
    If IsNull(mobjCV.DataValues("cvid").Value) Then
        ApproveChanged = False
        Exit Function
    Else
        lCVId = mobjCV.DataValues("cvid").Value
    End If
        
    If (LCase(StrType) = "endret") And (LOrgId > 0) Then
        Set ObjCmd = New ADODB.Command
        With rs
            'Add input parameters to command object
            Set ObjParam = New Parameter
            ObjParam.Name = "Lcvid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = lCVId
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                               
            Set ObjParam = New Parameter
            ObjParam.Name = "Ldataid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = LOrgId
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                                                              
            Set ObjParam = New Parameter
            ObjParam.Name = "strplace"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("place"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                
            Set ObjParam = New Parameter
            ObjParam.Name = "strcountry"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("country"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
            
            Set ObjParam = New Parameter
            ObjParam.Name = "strdescription"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("description"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "strTitle"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("Title"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "strFieldType"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = "UTD"
            ObjParam.Size = 3
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                
            Set ObjParam = New Parameter
            ObjParam.Name = "iFromMonth"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("FromMonth"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                            
            Set ObjParam = New Parameter
            ObjParam.Name = "iFromYear"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("FromYear"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                                                       
            Set ObjParam = New Parameter
            ObjParam.Name = "iToMonth"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("ToMonth"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "iToYear"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("ToYear"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "LTempid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = lDataID
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        End With
        'Oppdaterer hovedtabell for CV data
        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "UpdateCVDataMain", ObjCmd.Parameters) Then
            Err.Raise 701 + vbObjectError, "", "Could not update main data! StrSQL:" & strSQL
        End If
        Set ObjCmd = Nothing
        Set ObjDAL = Nothing
    
    ElseIf (LCase(StrType) = "slettet") And (LOrgId > 0) Then
        
        'Sletter erfarings-forekomst fra temptabell
        strSQL = "delete from CV_vikar_Data_temp where fieldType='UTD' and dataid=" & lDataID & " and cvid=" & lCVId
        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
            Err.Raise 703 + vbObjectError, "", "Could not delete temp data! StrSQL:" & strSQL
        End If
        'Sletter erfarings-forekomst fra hovedtabell
        strSQL = "delete from CV_Data where fieldType='UTD' and dataid=" & LOrgId & " and cvid=" & lCVId
        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
            Err.Raise 704 + vbObjectError, "", "Could not delete original data! StrSQL:" & strSQL
        End If
        Set ObjDAL = Nothing
    
    ElseIf (LCase(StrType) = "ny") And (LOrgId < 1) Then
        Set ObjCmd = New ADODB.Command
        With rs
            'Add input parameters to command object
            Set ObjParam = New Parameter
            ObjParam.Name = "lcvid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = .Fields("cvid").Value
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                               
            Set ObjParam = New Parameter
            ObjParam.Name = "strplace"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("place"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                
            Set ObjParam = New Parameter
            ObjParam.Name = "strcountry"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("country"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
            
            Set ObjParam = New Parameter
            ObjParam.Name = "strdescription"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("description"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "strTitle"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = ReturnString(.Fields.Item("Title"))
            ObjParam.Size = Len(ObjParam.Value)
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "strFieldType"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = "UTD"
            ObjParam.Size = 3
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                                
            Set ObjParam = New Parameter
            ObjParam.Name = "iFromMonth"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("FromMonth"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                            
            Set ObjParam = New Parameter
            ObjParam.Name = "iFromYear"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("FromYear"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
                                                       
            Set ObjParam = New Parameter
            ObjParam.Name = "iToMonth"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("ToMonth"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "iToYear"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = ReturnVal(.Fields("ToYear"))
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        
            Set ObjParam = New Parameter
            ObjParam.Name = "iTempDataid"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adInteger
            ObjParam.Value = lDataID
            ObjCmd.Parameters.Append ObjParam
            Set ObjParam = Nothing
        End With
        
        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "InsertintoCVDataMain", ObjCmd.Parameters) Then
            Err.Raise 705 + vbObjectError, "", "Could not update temp & main data! Procedure InsertintoCVDataMain."
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
    WriteLog "Educations_ApproveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function RejectChanges(ByVal lDataID As Long) As Boolean
On Error GoTo err_RejectChanges
    
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim RsOriginal As ADODB.Recordset
    Dim strSQL As String
    Dim StrType As String
    Dim LOrgId As Long
    Dim lCVId As Long
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
                     
    strSQL = "Select * from CV_vikar_Data_temp where FieldType='UTD' and dataid = " & lDataID
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        RejectChanges = False
        Exit Function
    End If
    
    StrType = rs.Fields("Type").Value
    If IsNull(rs.Fields("OriginaldataID").Value) Then
        LOrgId = 0
    Else
        LOrgId = rs.Fields("OriginaldataID").Value
    End If
    
    If IsNull(mobjCV.DataValues("cvid").Value) Then
        RejectChanges = False
        Exit Function
    Else
        lCVId = mobjCV.DataValues("cvid").Value
    End If
            
    If ((LCase(StrType) = "slettet") Or (LCase(StrType) = "endret")) And (LOrgId > 0) Then
        Set ObjCmd = New ADODB.Command
        Set rs = Nothing
        
        strSQL = "Select * from CV_Data where FieldType= 'UTD' and dataid = " & LOrgId & " and Cvid = " & lCVId
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        If Not rs.EOF Then
            With rs
                'Add input parameters to command object
                Set ObjParam = New Parameter
                ObjParam.Name = "Lcvid"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = .Fields("cvid").Value
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                   
                Set ObjParam = New Parameter
                ObjParam.Name = "LOrgDataid"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = LOrgId
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                   
                Set ObjParam = New Parameter
                ObjParam.Name = "LTempDataid"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = lDataID
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                   
                Set ObjParam = New Parameter
                ObjParam.Name = "strplace"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adVarChar
                ObjParam.Value = ReturnString(.Fields.Item("place"))
                ObjParam.Size = Len(ObjParam.Value)
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                    
                Set ObjParam = New Parameter
                ObjParam.Name = "strcountry"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adVarChar
                ObjParam.Value = ReturnString(.Fields.Item("country"))
                ObjParam.Size = Len(ObjParam.Value)
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                
                Set ObjParam = New Parameter
                ObjParam.Name = "strdescription"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adVarChar
                ObjParam.Value = ReturnString(.Fields.Item("description"))
                ObjParam.Size = Len(ObjParam.Value)
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
            
                Set ObjParam = New Parameter
                ObjParam.Name = "strTitle"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adVarChar
                ObjParam.Value = ReturnString(.Fields.Item("Title"))
                ObjParam.Size = Len(ObjParam.Value)
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
            
                Set ObjParam = New Parameter
                ObjParam.Name = "strFieldType"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adVarChar
                ObjParam.Value = "UTD"
                ObjParam.Size = 3
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                        
                Set ObjParam = New Parameter
                ObjParam.Name = "iFromMonth"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = ReturnVal(.Fields("FromMonth"))
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                
                Set ObjParam = New Parameter
                ObjParam.Name = "iFromYear"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = ReturnVal(.Fields("FromYear"))
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
                                                           
                Set ObjParam = New Parameter
                ObjParam.Name = "iToMonth"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = ReturnVal(.Fields("ToMonth"))
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
            
                Set ObjParam = New Parameter
                ObjParam.Name = "iToYear"
                ObjParam.Direction = adParamInput
                ObjParam.Type = adInteger
                ObjParam.Value = ReturnVal(.Fields("ToYear"))
                ObjCmd.Parameters.Append ObjParam
                Set ObjParam = Nothing
            
            End With
            
            If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "InsertintoCVDataTempOrg", ObjCmd.Parameters) Then
                Err.Raise 705 + vbObjectError, "", "Couldn't update temp data! Procedure InsertintoCVDataTempOrg."
            End If
            Set rs = Nothing
            Set ObjCmd = Nothing
            Set ObjDAL = Nothing
        Else
            Set rs = Nothing
            RejectChanges = False
            Exit Function
        End If
    ElseIf (LCase(StrType) = "ny") And (LOrgId < 1) Then
        'Sletter erfarings-forekomst fra temptabell
        strSQL = "delete from CV_vikar_Data_temp where FieldType='UTD' and dataid=" & lDataID & " and cvid=" & lCVId
        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
            Err.Raise 701 + vbObjectError, "", "Could not delete temp data! StrSQL:" & strSQL
        End If
        Set ObjDAL = Nothing
    Else
        RejectChanges = False
    End If
    
    RejectChanges = True
    Exit Function

err_RejectChanges:
    RejectChanges = False
    WriteLog "Educations_RejectChanges", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set mCol = Nothing
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function
