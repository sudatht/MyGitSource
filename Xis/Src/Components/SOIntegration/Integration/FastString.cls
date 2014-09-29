VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim stringArray() As String
Dim growthRate As Integer
Dim numItems As Long

Private Sub Class_Initialize()
    growthRate = 50
    numItems = 0
    ReDim stringArray(growthRate)
End Sub

Public Sub Append(ByVal strValue)
    ' next line prevents type mismatch error if strValue is null. Performance hit is negligible.
    strValue = strValue & ""
    If numItems > UBound(stringArray) Then
        ReDim Preserve stringArray(UBound(stringArray) + growthRate)
    End If
    stringArray(numItems) = strValue
    numItems = numItems + 1
End Sub

Public Sub Reset()
    Erase stringArray
    Class_Initialize
End Sub

Public Function ToString()
    ReDim Preserve stringArray(numItems)
    ToString = Join(stringArray, "")
End Function

