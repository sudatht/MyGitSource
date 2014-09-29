VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InitValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mobjdataValues As DataValues

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjdataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjdataValues
End Property

Private Sub Class_Initialize()
    Set mobjdataValues = New DataValues
End Sub

Private Sub Class_Terminate()
    Set mobjdataValues = Nothing
End Sub



