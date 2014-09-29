VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "InvoiceCustomerMonthSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjOwner As invoiceCustomerMonthSums

Public Property Set Owner(ByVal vData As invoiceCustomerMonthSums)
    Set mobjOwner = vData
End Property

Public Property Get Owner() As invoiceCustomerMonthSums
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property
  
Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjOwner = Nothing
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

