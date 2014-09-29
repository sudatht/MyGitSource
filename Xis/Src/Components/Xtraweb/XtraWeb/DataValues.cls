VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DataValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##CD Denne kolleksjonen er kjernen i hele objekt systemet. Denne kolleksjonen inneholder _
      %DataValue% objekter. For å refere til en verdi i et bestemt objekt bruker man syntaksen _
      objekt.DataValues(kolonne).value. Eks. 'objConsultant.DataValues("etternavn").value' henter ut _
      etternavnet til en konsulent.
      
'MTS
Implements ObjectControl
Private objContext As ObjectContext

'local variable to hold collection
Private mCol As Collection
'Private mobjOwner As Object
Private mstrTableName As String

Public Property Get TableName() As String
    TableName = mstrTableName
End Property

Public Property Let TableName(ByVal vNewValue As String)
    mstrTableName = vNewValue
End Property

Public Function Add(ByVal strValueName As String, Optional vValue As Variant) As DataValue
    Dim objNewMember As DataValue
    
    Set objNewMember = objContext.CreateInstance("xtraweb.DataValue")

    'set the properties passed into the method
    objNewMember.ValueName = strValueName
    objNewMember.Value = vValue
    mCol.Add objNewMember, strValueName
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As DataValue
Attribute Item.VB_UserMemId = 0
    
    If ExistsInCol(mCol, vntIndexKey) Then
        Set Item = mCol(vntIndexKey)
    Else
        Set Item = Add(vntIndexKey)
    End If
        
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Trace "inside DataValues.Remove vntIndexKey : " & vntIndexKey
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
Dim ObjVal As DataValue
    
    If mCol.Count > 0 Then
        For Each ObjVal In mCol
            Set ObjVal = Nothing
        Next ObjVal
    End If
    Set mCol = Nothing
    Set objContext = Nothing
End Sub

