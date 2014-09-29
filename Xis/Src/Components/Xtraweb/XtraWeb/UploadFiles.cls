VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "UploadFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' --------------------------------------------------------------------------
' Copyright(c) 2000-2006 Eurocenter DDC.
' No 65, Walukarama Road, Colombo 3, Sri Lanka
' All rights reserved.
'
' This software is the confidential and proprietary information of
' Eurocenter DDC (Confidential Information). You shall not disclose such
' Confidential Information and shall use it only in accordance with the
' terms of the license agreement you entered into with Eurocenter.
'
' Solution          : Xtra web integration
' Description       : This Class keeps a collection of UploadFile objects relevant to a CV object
' Author            : ska
' Created Timestamp : 29/06/2006
' --------------------------------------------------------------------------

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
' --------------------------------------------------------------------------
' Description       : Adds a UploadFile object to the collection
' Parameters        : UploadFile
' Author            : ska
' --------------------------------------------------------------------------
Public Function Add(ByVal objUploadFile As UploadFile) As UploadFile
    Dim strKey As String
    strKey = "ID" & Trim(objUploadFile.DataValues("UploadId").Value)
    mCol.Add objUploadFile, strKey
    'return the object created
    Set Add = objUploadFile
    Set objUploadFile = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As UploadFile
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property
' --------------------------------------------------------------------------
' Description       : Removes a UploadFile object from the collection
' Parameters        : Index
' Author            : ska
' --------------------------------------------------------------------------
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property
' --------------------------------------------------------------------------
' Description       : Fills the collcetion with cv_upload table relevant to a cv
' Author            : ska
' --------------------------------------------------------------------------
Public Sub Refresh()
On Error GoTo err_Refresh
    Trace "Enter method UploadFiles.Refresh()"

    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim objUploadFile As UploadFile
    Dim fldCol As ADODB.Field
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    lCVId = mobjCV.DataValues("CVid").Value
    
    Trace "cvid: " & lCVId
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New Recordset
    
    strSQL = "Select * from cv_upload where cvid = " & lCVId
    
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    
    Set mCol = New Collection
    
    While Not rs.EOF
        Set objUploadFile = objContext.CreateInstance("xtraweb.UploadFile")
        Set objUploadFile.Owner = mobjCV
        For Each fldCol In rs.Fields
            objUploadFile.DataValues(fldCol.Name) = fldCol.Value
        Next
        Add objUploadFile
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set ObjDAL = Nothing
    Trace "Exit method UploadFiles.Refresh()"
    Exit Sub
    
err_Refresh:
    WriteLog "UploadFiles_Refresh", Error$, EF_LogEventType_Error
End Sub

Private Sub ObjectControl_Activate()
    Trace "UploadFiles.Activate"
    
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

'Public Sub RefreshChanged()
'On Error GoTo err_Refresh
'    Dim ObjDAL As IdataAccess
'    Dim rs As ADODB.Recordset
'    Dim strSQL As String
'    Dim lCVId As Long
'    Dim objReference As Reference
'    Dim fldCol As ADODB.Field
'
'    If Trim(Len(mobjCV.XtraConString)) = 0 Then
'       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
'    End If
'
'    Set ObjDAL = New ClsDataAccess
'
'    lCVId = mobjCV.DataValues("CVid").Value
'    strSQL = "Select * from cv_vikar_references_temp where cvid = " & lCVId
'
'    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'
'    If rs.EOF Then
'        strSQL = "Select CVId, Name, Title, Tel, Comment, Firma, Type = 'ORIGINAL', OriginaldataID = ReferenceID from cV_references where cvid = " & lCVId
'
'        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'        If Not rs.EOF Then
'            While Not rs.EOF
'                With rs
'                    strSQL = "Insert into cv_vikar_references_temp( CVId, Name, Title, Tel, Comment, Firma, Type, OriginaldataID) values(" & .Fields("cvid").Value & ",'" & .Fields("Name").Value & "','" & .Fields("Title").Value & "','" & .Fields("Tel").Value & "','" & .Fields("Comment").Value & "','" & .Fields("firma").Value & "','ORIGINAL','" & .Fields("OriginaldataID").Value & "')"
'                    If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
'                        Err.Raise 701 + vbObjectError, "", "Error during insert! (References) StrSQL:" & strSQL
'                    End If
'                End With
'                rs.MoveNext
'            Wend
'            strSQL = "Select * from cv_vikar_references_temp where cvid = " & lCVId
'            Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'        End If
'    End If
'    If Not rs.EOF Then
'        Set mCol = New Collection
'        While Not rs.EOF
'            Set objReference = objContext.CreateInstance("xtraweb.reference")
'            Set objReference.Owner = mobjCV
'            For Each fldCol In rs.Fields
'                objReference.DataValues(fldCol.Name) = fldCol.Value
'            Next
'            Add objReference
'            rs.MoveNext
'        Wend
'    End If
'    Set rs = Nothing
'    Set ObjDAL = Nothing
'    Exit Sub
'
'err_Refresh:
'    WriteLog "References_RefreshChanged", Error$, EF_LogEventType_Error
'End Sub

'Public Function RejectChanges(ByVal LreferenceID As Long) As Boolean
'On Error GoTo err_RejectChanges
'
'    Dim ObjDAL As IdataAccess
'    Dim objValue As XtraWeb.DataValue
'    Dim rs As ADODB.Recordset
'    Dim RsOriginal As ADODB.Recordset
'    Dim strSQL As String
'    Dim StrType As String
'    Dim LOrgId As Long
'    Dim lCVId As Long
'    Dim ObjCmd As ADODB.Command
'    Dim ObjParam As ADODB.Parameter
'
'    If Trim(Len(mobjCV.XtraConString)) = 0 Then
'       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
'    End If
'
'    Set ObjDAL = New ClsDataAccess
'
'    strSQL = "Select * from cv_vikar_references_temp where ReferenceID = " & LreferenceID
'    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'    If rs.RecordCount = 0 Then
'        Set rs = Nothing
'        RejectChanges = False
'        Exit Function
'    End If
'
'    StrType = rs.Fields("Type").Value
'    If IsNull(rs.Fields("OriginaldataID").Value) Then
'        LOrgId = 0
'    Else
'        LOrgId = rs.Fields("OriginaldataID").Value
'    End If
'
'    If IsNull(mobjCV.DataValues("cvid").Value) Then
'        RejectChanges = False
'        Exit Function
'    Else
'        lCVId = mobjCV.DataValues("cvid").Value
'    End If
'
'    If ((LCase(StrType) = "slettet") Or (LCase(StrType) = "endret")) And (LOrgId > 0) Then
'        Set ObjCmd = New ADODB.Command
'        Set rs = Nothing
'
'        strSQL = "Select * from cv_references where ReferenceID = " & LOrgId & " and Cvid = " & lCVId
'        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'        If Not rs.EOF Then
'            With rs
'                'Add input parameters to command object
'                Set ObjParam = New Parameter
'                ObjParam.Name = "Lcvid"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adInteger
'                ObjParam.Value = .Fields("cvid").Value
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "LOrgReferenceid"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adInteger
'                ObjParam.Value = LOrgId
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "LtempReferenceid"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adInteger
'                ObjParam.Value = LreferenceID
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "strName"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adVarChar
'                ObjParam.Value = ReturnString(.Fields.Item("name"))
'                ObjParam.Size = Len(ObjParam.Value)
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "StrTitle"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adVarChar
'                ObjParam.Value = ReturnString(.Fields.Item("Title"))
'                ObjParam.Size = Len(ObjParam.Value)
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "StrTel"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adVarChar
'                ObjParam.Value = ReturnString(.Fields.Item("Tel"))
'                ObjParam.Size = Len(ObjParam.Value)
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "StrFirma"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adVarChar
'                ObjParam.Value = ReturnString(.Fields.Item("firma"))
'                ObjParam.Size = Len(ObjParam.Value)
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'                Set ObjParam = New Parameter
'                ObjParam.Name = "StrComment"
'                ObjParam.Direction = adParamInput
'                ObjParam.Type = adVarChar
'                ObjParam.Value = ReturnString(.Fields.Item("Comment"))
'                ObjParam.Size = Len(ObjParam.Value)
'                ObjCmd.Parameters.Append ObjParam
'                Set ObjParam = Nothing
'
'            End With
'
'            If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "InsertintoCVReferenceTempOrg", ObjCmd.Parameters) Then
'                Err.Raise 705 + vbObjectError, "", "Couldn't update temp data! Procedure InsertintoCVReferenceTempOrg."
'            End If
'            Set rs = Nothing
'            Set ObjCmd.ActiveConnection = Nothing
'            Set ObjCmd = Nothing
'            Set ObjDAL = Nothing
'        Else
'            Set rs = Nothing
'            Set ObjDAL = Nothing
'            RejectChanges = False
'            Exit Function
'        End If
'    ElseIf (LCase(StrType) = "ny") And (LOrgId < 1) Then
'        'Sletter erfarings-forekomst fra temptabell
'        strSQL = "delete from cv_vikar_references_temp where ReferenceID =" & LreferenceID & " and cvid=" & lCVId
'        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
'            Err.Raise 701 + vbObjectError, "", "Could not delete temp Reference data! StrSQL:" & strSQL
'        End If
'        Set ObjDAL = Nothing
'    Else
'        RejectChanges = False
'    End If
'
'    RejectChanges = True
'    Exit Function
'
'err_RejectChanges:
'    RejectChanges = False
'    WriteLog "References_RejectChanges", Error$, EF_LogEventType_Error
'    Set rs = Nothing
'    Set mCol = Nothing
'    Set ObjDAL = Nothing
'    Set ObjCmd = Nothing
'End Function


'Public Function ApproveChanged(ByVal LreferenceID As Long) As Boolean
'On Error GoTo err_ApproveChanged
'
'    Dim ObjDAL As IdataAccess
'    Dim objValue As XtraWeb.DataValue
'    Dim rs As ADODB.Recordset
'    Dim RsOriginal As ADODB.Recordset
'    Dim strSQL As String
'    Dim StrType As String
'    Dim LOrgId As Long
'    Dim lCVId As Long
'    Dim LtempInt As Long
'    Dim ObjCmd As ADODB.Command
'    Dim ObjParam As ADODB.Parameter
'
'    If Trim(Len(mobjCV.XtraConString)) = 0 Then
'       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
'    End If
'
'    Set ObjDAL = New ClsDataAccess
'
'    strSQL = "Select * from  cv_vikar_references_temp where ReferenceID = " & LreferenceID
'    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
'    If rs.RecordCount = 0 Then
'        Set rs = Nothing
'        Set ObjDAL = Nothing
'        ApproveChanged = False
'        Exit Function
'    End If
'
'    StrType = rs.Fields("Type").Value
'    If IsNull(rs.Fields("OriginaldataID").Value) Then
'        LOrgId = 0
'    Else
'        LOrgId = rs.Fields("OriginaldataID").Value
'    End If
'
'    If IsNull(mobjCV.DataValues("cvid").Value) Then
'        ApproveChanged = False
'        Exit Function
'    Else
'        lCVId = mobjCV.DataValues("cvid").Value
'    End If
'
'    If (LCase(StrType) = "endret") And (LOrgId > 0) Then
'        Set ObjCmd = New ADODB.Command
'        With rs
'            Set ObjParam = New Parameter
'            ObjParam.Name = "Lcvid"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adInteger
'            ObjParam.Value = .Fields("cvid").Value
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "LOrgReferenceid"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adInteger
'            ObjParam.Value = LOrgId
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "LtempReferenceid"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adInteger
'            ObjParam.Value = LreferenceID
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "strName"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("name"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrTitle"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Title"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrTel"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Tel"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrFirma"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("firma"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrComment"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Comment"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'        End With
'        'Oppdaterer hovedtabell for CV data
'        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "UpdateCVReferenceMain", ObjCmd.Parameters) Then
'            Err.Raise 701 + vbObjectError, "", "Could not update main data! Stored procedure: UpdateCVReferenceMain."
'        End If
'        Set ObjCmd = Nothing
'        Set ObjDAL = Nothing
'
'    ElseIf (LCase(StrType) = "slettet") And (LOrgId > 0) Then
'
'        'Sletter erfarings-forekomst fra temptabell
'        strSQL = "delete from cv_vikar_references_temp where ReferenceID=" & LreferenceID & " and cvid=" & lCVId
'        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
'            Err.Raise 702 + vbObjectError, "", "Could not delete temp reference data! StrSQL:" & strSQL
'        End If
'        'Sletter erfarings-forekomst fra hovedtabell
'        strSQL = "delete from cv_references where ReferenceID=" & LOrgId & " and cvid=" & lCVId
'        If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
'            Err.Raise 703 + vbObjectError, "", "Could not delete original reference data! StrSQL:" & strSQL
'        End If
'    ElseIf (LCase(StrType) = "ny") And (LOrgId < 1) Then
'        Set ObjCmd = New ADODB.Command
'        With rs
'            Set ObjParam = New Parameter
'            ObjParam.Name = "Lcvid"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adInteger
'            ObjParam.Value = .Fields("cvid").Value
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "LtempReferenceid"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adInteger
'            ObjParam.Value = LreferenceID
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "strName"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("name"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrTitle"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Title"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrTel"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Tel"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrFirma"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("firma"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'
'            Set ObjParam = New Parameter
'            ObjParam.Name = "StrComment"
'            ObjParam.Direction = adParamInput
'            ObjParam.Type = adVarChar
'            ObjParam.Value = ReturnString(.Fields.Item("Comment"))
'            ObjParam.Size = Len(ObjParam.Value)
'            ObjCmd.Parameters.Append ObjParam
'            Set ObjParam = Nothing
'        End With
'
'        If Not ObjDAL.ExecuteCRUDSP(mobjCV.XtraConString, "InsertintoCVReferenceMain", ObjCmd.Parameters) Then
'            Err.Raise 705 + vbObjectError, "", "Could not update temp & main data! Procedure InsertintoCVReferenceMain."
'        End If
'        Set ObjCmd.ActiveConnection = Nothing
'        Set ObjCmd = Nothing
'    Else
'        ApproveChanged = False
'        Exit Function
'    End If
'    Set ObjDAL = Nothing
'    ApproveChanged = True
'    Exit Function
'
'err_ApproveChanged:
'    ApproveChanged = False
'    WriteLog "References_ApproveChanged", Error$, EF_LogEventType_Error
'    Set rs = Nothing
'    Set mCol = Nothing
'    Set ObjDAL = Nothing
'    Set ObjCmd = Nothing
'End Function



