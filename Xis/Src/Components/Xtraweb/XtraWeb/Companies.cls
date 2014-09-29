VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Companies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'MTS
Implements ObjectControl
Private objContext As ObjectContext
Private mStrXtraDataShapeCon As String
Private mCol As Collection

Public Function Add(ByVal objNewMember As Company, Optional strKey As String) As Company
    
    If Len(strKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, strKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Company
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

Public Property Get XtraDataShapeConString() As String
    XtraDataShapeConString = mStrXtraDataShapeCon
End Property

Public Property Let XtraDataShapeConString(strConString As String)
    mStrXtraDataShapeCon = strConString
End Property

Public Sub Refresh()
'On Error GoTo err_Refresh
    Dim ObjDAL          As IdataAccess
    Dim rsComp          As ADODB.Recordset
    Dim rsLoc           As ADODB.Recordset
    Dim fldData         As ADODB.Field
    Dim strSQL          As String
    Dim objCompany      As Company
    Dim objLocation     As Location
    Dim strCompID       As String
    Dim strLocID        As String
    
    If Trim(Len(mStrXtraDataShapeCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No Datashape connection string!"
    End If

    Set ObjDAL = New ClsDataAccess
    
    strSQL = "SHAPE {SELECT * FROM SELSKAP}  AS Selskap APPEND ({SELECT * FROM AVDELINGSKONTOR order by web_sortering}  AS AvdKontor RELATE 'SelskapID' TO 'SelskapID') AS AvdKontor"
    Set rsComp = New ADODB.Recordset
    
    Set rsComp = ObjDAL.FetchRC(mStrXtraDataShapeCon, strSQL)
        
    While Not rsComp.EOF
        Set objCompany = objContext.CreateInstance("xtraweb.Company")
        
        For Each fldData In rsComp.Fields
            objCompany.DataValues(fldData.Name) = fldData.Value
        Next
        strCompID = "ID" & objCompany.DataValues("SelskapID").Value
        
        
        Set rsLoc = rsComp("Avdkontor").Value
        
        While Not rsLoc.EOF
            Set objLocation = objContext.CreateInstance("xtraweb.Location")
            
            For Each fldData In rsLoc.Fields
                objLocation.DataValues(fldData.Name) = fldData.Value
            Next
            strLocID = "ID" & objLocation.DataValues("ID").Value
            objCompany.Locations.Add objLocation, strLocID
            
            rsLoc.MoveNext
        
        Wend
        Add objCompany, strCompID
        rsComp.MoveNext
    Wend
    Set rsLoc = Nothing
    Set rsComp = Nothing
    Set ObjDAL = Nothing
    Exit Sub
err_Refresh:
    WriteLog "JobGroups_Refresh", Error$, EF_LogEventType_Error
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set objContext = Nothing
End Sub


