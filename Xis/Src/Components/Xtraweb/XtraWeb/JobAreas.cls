VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjCV As CV

Public Property Set CV(ByVal vData As CV)
    Set mobjCV = vData
End Property

Public Property Get CV() As CV
    Set CV = mobjCV
End Property

Public Function Add(ByVal objJobArea As JobArea, Optional strKey As String) As JobArea
    
    If Len(strKey) = 0 Then
        mCol.Add objJobArea
    Else
        mCol.Add objJobArea, strKey
    End If

    'return the object created
    Set Add = objJobArea
    Set objJobArea = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As JobArea
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

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjCV = Nothing
    Set objContext = Nothing
End Sub

