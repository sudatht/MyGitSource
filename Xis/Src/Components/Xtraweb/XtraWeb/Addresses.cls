VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mCol As Collection 'local variable to hold collection
Private mobjOwner As Object

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Public Property Set Owner(ByVal vData As Object)
    Set mobjOwner = vData
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Function Add(ByVal objaddress As Address, Optional strKey As String) As Address

    If Len(strKey) = 0 Then
        mCol.Add objaddress
    Else
        mCol.Add objaddress, strKey
    End If

    'return the object created
    Set Add = objaddress
    Set objaddress = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Address
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
    ' ##BD Refreshes the address collection
On Error GoTo err_Refresh
    
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim fldData As ADODB.Field
    Dim rsAddress As ADODB.Recordset
    Dim lOwnerID As Long
    Dim strTable As String
    Dim objaddress As Address
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " Owner has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rsAddress = New ADODB.Recordset
    
    If TypeOf mobjOwner Is Suspect Then
        strTable = "v_suspect_adresse"
        lOwnerID = mobjOwner.DataValues("suspectid").Value
    ElseIf TypeOf mobjOwner Is Consultant Then
        strTable = "adresse"
        lOwnerID = mobjOwner.DataValues("vikarid").Value
    ElseIf TypeOf mobjOwner Is Customer Then
        strTable = "adresse"
        lOwnerID = mobjOwner.DataValues("FirmaID").Value
    End If
    strSQL = "Select * from " & strTable & " where adresserelid = " & lOwnerID
    
    Set mCol = New Collection
    
    Set rsAddress = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    While Not rsAddress.EOF
        Set objaddress = objContext.CreateInstance("xtraweb.Address")
        Set objaddress.Owner = mobjOwner
        For Each fldData In rsAddress.Fields
            objaddress.DataValues(fldData.Name) = fldData.Value
        Next
        Add objaddress
        rsAddress.MoveNext
    Wend
    Set rsAddress = Nothing
    Set ObjDAL = Nothing
    Exit Sub
err_Refresh:
    WriteLog "Addresses_Refresh", Error$, EF_LogEventType_Error
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
    Set objContext = Nothing
End Sub
Public Sub CleanUp()
    Set mCol = Nothing
    Set mobjOwner = Nothing
End Sub

