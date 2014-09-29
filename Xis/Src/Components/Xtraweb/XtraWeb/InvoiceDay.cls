VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "InvoiceDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjOwner As Invoice
Private mIweekdayno As Integer

Public Property Get WeekDayNo() As Integer
    WeekDayNo = DatePart("w", CDate(mobjDataValues("dagdato").Value), vbMonday, vbFirstFourDays)
End Property

Public Property Set Owner(ByVal objInvoice As Invoice)
    Set mobjOwner = objInvoice
End Property

Public Property Get Owner() As Invoice
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
    Set mobjOwner = objContext.CreateInstance("xtraweb.Invoice")
End Sub

Private Sub ObjectControl_Deactivate()
    'Set mobjDataValues.Owner = Nothing
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

