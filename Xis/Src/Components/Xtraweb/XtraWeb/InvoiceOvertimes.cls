VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "InvoiceOvertimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjInvoice As Invoice

Public Property Get Owner() As Invoice
    Set Owner = mobjInvoice
End Property

Public Property Set Owner(ByRef objInvoice As Invoice)
    Set mobjInvoice = objInvoice
End Property

Public Function Add(ByVal ObjOvertime As InvoiceOvertime) As InvoiceOvertime
    Dim strKey As String
    strKey = "ID" & Trim(ObjOvertime.DataValues("overtidType").Value)
    mCol.Add ObjOvertime, strKey
    
    'return the object created
    Set Add = ObjOvertime
    Set ObjOvertime = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As InvoiceOvertime
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjInvoice = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

