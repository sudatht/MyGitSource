VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "JobPlaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

'local variable to hold collection
Private mCol As Collection
Private mobjOwner As Object

Public Property Set Owner(ByVal vData As Object)
    Set mobjOwner = vData
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Function Add(ByVal objJobPlace As JobPlace, Optional strKey As String) As JobPlace
    
    If Len(strKey) = 0 Then
        mCol.Add objJobPlace
    Else
        mCol.Add objJobPlace, strKey
    End If

    'return the object created
    Set Add = objJobPlace
    Set objJobPlace = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As JobPlace
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

Public Sub Refresh()
On Error GoTo err_Refresh
    ' ##BD Refreshes the Jobplaces collection
    
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim fldData As ADODB.Field
    Dim rs As ADODB.Recordset
    Dim lOwnerID As Long
    Dim strTable As String
    Dim objJobPlace As JobPlace
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    lOwnerID = mobjOwner.DataValues("JobApplicationID").Value
    strSQL = "SELECT * FROM [jobplace] WHERE [JobApplicationid] = " & lOwnerID
    
    Set mCol = New Collection
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    While Not rs.EOF
        Set objJobPlace = objContext.CreateInstance("xtraweb.JobPlace")
        Set objJobPlace.Owner = mobjOwner
        For Each fldData In rs.Fields
            objJobPlace.DataValues(fldData.Name) = fldData.Value
        Next
        Add objJobPlace
        rs.MoveNext
    Wend
    Set rs = Nothing
    Set objJobPlace = Nothing
    Set ObjDAL = Nothing
    Exit Sub

err_Refresh:
    WriteLog "JobPlaces_Refresh", Error$, EF_LogEventType_Error
End Sub

Public Sub Purge()
On Error GoTo err_Purge
    ' ##BD purges the Jobplaces collection
    
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim fldData As ADODB.Field
    Dim rs As ADODB.Recordset
    Dim lOwnerID As Long
    Dim strTable As String
    Dim objJobPlace As JobPlace
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " CV has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    lOwnerID = mobjOwner.DataValues("JobApplicationID").Value
    strSQL = "DELETE FROM [jobplace] WHERE [Jobapplicationid] = " & lOwnerID
        
    ObjDAL.ExecuteCRUD mobjOwner.XtraConString, strSQL
     While (mCol.Count > 0)
        Set mCol(1).Owner = Nothing
        mCol.Remove (1)
     Wend
     
    Set ObjDAL = Nothing
    Exit Sub

err_Purge:
    WriteLog "JobPlaces_Purge", Error$, EF_LogEventType_Error
End Sub


Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub
