VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "invoiceCustomerMonthSums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjCustomerMonth As InvoiceCustomerMonth

Public Property Get Owner() As InvoiceCustomerMonth
    Set Owner = mobjCustomerMonth
End Property

Public Property Set Owner(ByRef objInvoiceMonth As InvoiceCustomerMonth)
    Set mobjCustomerMonth = objInvoiceMonth
End Property

Public Function Add(ByVal ObjMonthSum As InvoiceCustomerMonthSum) As InvoiceCustomerMonthSum
    mCol.Add ObjMonthSum
    'return the object created
    Set Add = ObjMonthSum
    Set ObjMonthSum = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As InvoiceCustomerMonthSum
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
    Set mobjCustomerMonth = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

