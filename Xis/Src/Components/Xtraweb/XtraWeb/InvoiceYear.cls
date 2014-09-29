VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection
Private mobjInvoice As Invoicebasis

Public Property Get Owner() As Invoicebasis
    Set Owner = mobjInvoice
End Property

Public Property Set Owner(ByRef ObjInvoice As Invoicebasis)
    Set mobjInvoice = ObjInvoice
End Property

Public Function Add(ByRef ObjMonth As InvoiceMonth) As InvoiceMonth
    Dim strKey As String
    'strKey = "ID" & Trim(ObjMonth.DataValues("dataid").Value)
    mCol.Add ObjMonth, strKey
    
    'return the object created
    Set Add = ObjMonth
    Set ObjMonth = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Experience
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

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mobjInvoice = Nothing
End Sub

Public Sub Refresh()
End Sub

