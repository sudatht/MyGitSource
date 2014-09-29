VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DataValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mstrValueName As String
Private mvValue As Variant

Public Property Let Value(ByVal vData As Variant)
    mvValue = vData
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = mvValue
End Property

Public Property Let ValueName(ByVal vData As String)
    mstrValueName = vData
End Property

Public Property Get ValueName() As String
    ValueName = mstrValueName
End Property

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub



