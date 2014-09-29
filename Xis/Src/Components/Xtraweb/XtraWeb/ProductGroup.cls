VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "ProductGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjProducts As Products

Public Property Set Products(ByVal vData As Products)
    Set mobjProducts = vData
End Property

Public Property Get Products() As Products
    Set Products = mobjProducts
End Property

Public Sub RefreshChanged(strXtraCon As String, lCVId As Long)
On Error GoTo err_Refresh
    Dim ObjDAL              As IdataAccess
    Dim rsProds             As ADODB.Recordset
    Dim fldData             As ADODB.Field
    Dim strSQL              As String
    Dim Lvikarid            As Long
    Dim objProduct          As Product
    Dim IProductGroupID     As Long

    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If

    IProductGroupID = mobjDataValues("ProdOmradeID").Value
        
    strSQL = "Select count(*) from CV_vikar_productknowledge_temp " & _
            "Where cvid = " & lCVId

    Set ObjDAL = New ClsDataAccess
    Set rsProds = ObjDAL.FetchRC(strXtraCon, strSQL)
    If rsProds(0) = 0 Then
        strSQL = "Select k_tittelID from vikar_kompetanse where k_typeid = 3 and cvid = " & lCVId
        Set rsProds = ObjDAL.FetchRC(strXtraCon, strSQL)
        If Not rsProds.EOF Then
            With rsProds
                While Not .EOF
                        strSQL = "Insert into CV_vikar_productknowledge_temp(CVId, ProdID, Original_value, Type) values(" & lCVId & "," & .Fields("k_tittelID").Value & ", 1, 'ORIGINAL')"
                        If Not ObjDAL.ExecuteCRUD(strXtraCon, strSQL) Then
                            Err.Raise 701 + vbObjectError, "", "Error during insert! (Products) StrSQL:" & strSQL
                        End If
                    rsProds.MoveNext
                Wend
            End With
        End If
    End If
    Set rsProds = Nothing
        
    strSQL = "SELECT [kt].[k_tittelID], [kt].[Ktittel], [kt].[prodomradeID], [vp].[Type], [vp].[Original_value], " & _
    "Checked = CASE WHEN [vp].[CVid] IS NULL THEN '' ELSE 'checked' END " & _
    "FROM [CV_vikar_productknowledge_temp] AS [vp], [H_KOMP_TITTEL] AS [kt] " & _
    "WHERE [vp].[cvid] = " & lCVId & " " & _
    "AND [kt].[k_tittelid] *= [vp].[prodid] " & _
    "AND [kt].[web_komp_visible] = 1 " & _
    "AND [kt].[k_typeID] = 3 " & _
    "AND [kt].[prodomradeid] = " & IProductGroupID & " " & _
    "ORDER BY [ktittel] "

    Set mobjProducts = Nothing
    Set mobjProducts = objContext.CreateInstance("xtraweb.Products")
    Set rsProds = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    While Not rsProds.EOF
        Set objProduct = objContext.CreateInstance("xtraweb.Product")
        For Each fldData In rsProds.Fields
            objProduct.DataValues(fldData.Name).Value = fldData.Value
        Next
        If LCase(objProduct.DataValues("Checked").Value) = "checked" Then
            objProduct.Checked = True
        End If
        mobjProducts.Add objProduct
        rsProds.MoveNext
    Wend
    Set rsProds = Nothing
    Set ObjDAL = Nothing
    Exit Sub

err_Refresh:
    WriteLog "ProductGroup_RefreshChanged", Error$, EF_LogEventType_Error
End Sub

Public Sub Refresh(strXtraCon As String, lCVId As Long)
On Error GoTo err_Refresh
    Dim ObjDAL              As IdataAccess
    Dim rsProds             As ADODB.Recordset
    Dim fldData             As ADODB.Field
    Dim strSQL              As String
    Dim Lvikarid            As Long
    Dim objProduct          As Product
    Dim IProductGroupID     As Long

    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess

    IProductGroupID = mobjDataValues("ProdOmradeID").Value
    strSQL = "SELECT [kt].[k_tittelID], [kt].[Ktittel], [kt].[prodomradeID], " & _
            "[Checked] = CASE WHEN [vk].[CVid] IS NULL THEN '' ELSE 'checked' END " & _
            "FROM [H_KOMP_TITTEL] AS [kt], [H_KOMP_AREA] AS [ka], [vikar_kompetanse] AS [vk] " & _
            "WHERE [kt].[web_komp_visible] = 1 " & _
            "AND [vk].[k_typeid] =* [kt].[k_typeID] " & _
            "AND [vk].[K_TittelID] =* [kt].[K_TittelID] " & _
            "AND [kt].[prodomradeid] = " & IProductGroupID & " " & _
            "AND [kt].[k_typeID] = 3 " & _
            "AND [kt].[ProdOmradeID] = [ka].[ProdOmradeID] " & _
            "AND  [vk].[cvid] = " & lCVId & " " & _
            "ORDER BY [ktittel] "
                        
    Set rsProds = ObjDAL.FetchRC(strXtraCon, strSQL)
    While Not rsProds.EOF
        Set objProduct = objContext.CreateInstance("xtraweb.Product")
        For Each fldData In rsProds.Fields
            objProduct.DataValues(fldData.Name).Value = fldData.Value
        Next
        If LCase(objProduct.DataValues("Checked").Value) = "checked" Then
            objProduct.Checked = True
        End If
        mobjProducts.Add objProduct
        rsProds.MoveNext
    Wend
    Set rsProds = Nothing
    Set ObjDAL = Nothing
    Exit Sub

err_Refresh:
    WriteLog "ProductGroup_Refresh", Error$, EF_LogEventType_Error
End Sub

Public Function GetAllNonApproved(strXtraCon As String, lCVId As Long) As ADODB.Recordset
On Error GoTo err_GetAllNonApproved
    Dim ObjDAL              As IdataAccess
    Dim strSQL              As String
    
    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If
                
    strSQL = "SELECT vp.type, ka.Produktomrade, kt.Ktittel, kt.k_tittelid " & _
        "FROM H_KOMP_TITTEL kt, H_KOMP_AREA ka, CV_vikar_productknowledge_temp vp " & _
        "Where kt.web_komp_visible = 1 " & _
        "and vp.ProdID = kt.K_TittelID " & _
        "and kt.k_typeID = 3 " & _
        "and vp.type <> 'ORIGINAL' " & _
        "and kt.ProdOmradeID = ka.ProdOmradeID " & _
        "and vp.cvid = " & lCVId & " " & _
        "order by ka.Produktomrade, kt.Ktittel "
   
    Set ObjDAL = New ClsDataAccess
    Set GetAllNonApproved = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    Exit Function

err_GetAllNonApproved:
    WriteLog "ProductGroup_GetAllNonApproved", Error$, EF_LogEventType_Error
End Function

Public Function GetAllApproved(strXtraCon As String, lCVId As Long) As ADODB.Recordset
On Error GoTo err_GetAllApproved
    Dim ObjDAL              As IdataAccess
    Dim strSQL              As String
    
    If lCVId < 1 Then
       Err.Raise 700 + vbObjectError, "", " Missing CVid parameter!"
    End If
    
    If Trim(Len(strXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " Missing connection string parameter!"
    End If
                
    strSQL = "SELECT ka.Produktomrade, kt.Ktittel, kt.k_tittelid, k_levelid, vk.Rangering " & _
        "FROM H_KOMP_TITTEL kt, H_KOMP_AREA ka, vikar_kompetanse vk " & _
        "Where kt.web_komp_visible = 1 " & _
        "and vk.K_TittelID = kt.K_TittelID " & _
        "and kt.k_typeID = 3 " & _
        "and kt.ProdOmradeID = ka.ProdOmradeID " & _
        "and vk.cvid = " & lCVId & " " & _
        "order by ka.Produktomrade, kt.Ktittel "
    
   
    Set ObjDAL = New ClsDataAccess
    Set GetAllApproved = ObjDAL.FetchRC(strXtraCon, strSQL)
    Set ObjDAL = Nothing
    Exit Function

err_GetAllApproved:
    WriteLog "ProductGroup_GetAllApproved", Error$, EF_LogEventType_Error
End Function

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjProducts = objContext.CreateInstance("xtraweb.Products")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjProducts = Nothing
    Set objContext = Nothing
End Sub
