VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Product"
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

Public Property Set Owner(ByRef objCV As CV)
    Set mobjCV = objCV
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
    Dim LProdID As Long
    Dim Lvikarid As Long
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = CLng(mobjCV.DataValues("CVId").Value)
    LProdID = CLng(mobjDataValues("ProdID").Value)
    
    If TypeOf mobjCV.Owner Is Consultant Then
        Lvikarid = CLng(mobjCV.Owner.DataValues("vikarid").Value)
    ElseIf TypeOf mobjCV.Owner Is Suspect Then
        Lvikarid = 0
    End If
    
    ' check if product already exist on this cv
    strSQL = "SELECT COUNT(*) FROM [vikar_kompetanse] WHERE [K_TypeID] = 3 AND [cvid] = " & lCVId & " AND [K_TittelID] = " & LProdID
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If rs(0) = 0 Then
        'product does not exist
        'Create empty RS
        strSQL = "SELECT * FROM [vikar_kompetanse] WHERE 1 = NULL"
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        'Add new
        rs.AddNew
        'Kompetanse
        rs("K_TypeID") = 3
        'Not rated
        rs("K_LevelID") = 0
        rs("vikarid") = Lvikarid
        rs("cvid") = lCVId
        rs("K_TittelID") = LProdID
        ObjDAL.BatchUpdate mobjCV.XtraConString, rs
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Save = True
    Exit Function

err_Save:
    Save = False
    WriteLog "Product_Save", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set rs = Nothing
End Function


Public Function Delete() As Boolean
On Error GoTo err_Delete
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim LProdID As Long
    
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    lCVId = mobjCV.DataValues("CVId").Value
    LProdID = mobjDataValues("ProdID").Value
    
    strSQL = "Select K_TittelID from vikar_kompetanse where K_TypeID = 3 and cvid = " & lCVId & " and K_TittelID = " & LProdID
    Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    If Not rs.EOF Then
        rs.Delete
        ObjDAL.BatchUpdate mobjCV.XtraConString, rs
    End If
    Delete = True
    
    Exit Function
err_Delete:
    Delete = False
    WriteLog "Product_Delete", Error$, EF_LogEventType_Error
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

Public Function SaveChanged() As Boolean
On Error GoTo err_Savechanged
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim RsOriginal As ADODB.Recordset
    Dim strSQL As String
    Dim lCVId As Long
    Dim LProdID As Long
    Dim strChecked As String
    Dim StrType As String
    Dim BlnSave As Boolean
    Dim BlnUpdate As Boolean
        
    If Trim(Len(mobjCV.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If
    
    lCVId = mobjCV.DataValues("CVId").Value
    LProdID = mobjDataValues("K_TittelID").Value
    strChecked = mobjDataValues("Checked").Value
    
    BlnSave = False
    BlnUpdate = False
    
    Set ObjDAL = New ClsDataAccess
    ' check if Product already exist on this cv
    strSQL = "Select * from CV_vikar_productknowledge_temp where cvid = " & lCVId & " and ProdID = " & LProdID
    Set RsOriginal = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
    'If product is checked -
    If strChecked = "CHECKED" Then
        'If the product doesn't exist
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
    'If not checked and the product exists in the original CV
        If Not RsOriginal.EOF Then
            ' Mark as ready to be deleted
            If RsOriginal.Fields("Original_value") Then
                StrType = "SLETTET"
                BlnUpdate = True
            'If not checked and product is not part of original CV, delete at once..
            Else
                strSQL = "delete from CV_vikar_productknowledge_temp where cvid =" & lCVId & " and ProdID=" & LProdID
                If Not ObjDAL.ExecuteCRUD(mobjCV.XtraConString, strSQL) Then
                    Err.Raise 701 + vbObjectError, "", "Couldn't delete product! CVid:" & lCVId & ", prodID:" & LProdID
                End If
            End If
        Else
            Set RsOriginal = Nothing
            Set ObjDAL = Nothing
            SaveChanged = False
            Exit Function
        End If
    End If
    If BlnSave Then
        strSQL = "Select * from CV_vikar_productknowledge_temp where 1 = null"
        Set rs = ObjDAL.FetchRC(mobjCV.XtraConString, strSQL)
        rs.AddNew
        rs("cvid") = lCVId
        rs("Prodid") = LProdID
        rs("type") = StrType
        ObjDAL.BatchUpdate mobjCV.XtraConString, rs
    End If
    
    If BlnUpdate Then
        RsOriginal("type") = StrType
        ObjDAL.BatchUpdate mobjCV.XtraConString, RsOriginal
        RsOriginal.Close
    End If
    Set RsOriginal = Nothing
    Set rs = Nothing
    Set ObjDAL = Nothing
    'Make sure CV is marked as changed
    If mobjCV.DataValues("cv_changed").Value = 0 Then
        mobjCV.SaveChanged
    End If
    SaveChanged = True
    Exit Function
    
err_Savechanged:
    SaveChanged = False
    WriteLog "Product_SaveChanged", Error$, EF_LogEventType_Error
    Set rs = Nothing
    Set RsOriginal = Nothing
    Set ObjDAL = Nothing
End Function

