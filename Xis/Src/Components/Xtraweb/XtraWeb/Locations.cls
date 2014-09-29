VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Locations"
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

Public Property Get XtraDataShapeConString() As String
    XtraDataShapeConString = mStrXtraDataShapeCon
End Property

Public Property Let XtraDataShapeConString(strConString As String)
    mStrXtraDataShapeCon = strConString
End Property

Public Function Add(ByVal objNewMember As Location, Optional strKey As String) As Location
    
    If Len(strKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, strKey
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Location
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
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
         
    Dim ObjDAL          As IdataAccess
    Dim rsLoc           As ADODB.Recordset
    Dim rsComp          As ADODB.Recordset
    Dim fldData         As ADODB.Field
    Dim strSQL          As String
    Dim objComp         As Company
    Dim objLocation     As Location
    Dim strCompID       As String
    Dim strLocID        As String
    
    If Trim(Len(mStrXtraDataShapeCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No Datashape connection string!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    strSQL = "SHAPE {SELECT * FROM LOKASJON}  AS Lokasjon APPEND ({SELECT t1.*, t2.navn selskap FROM AVDELINGSKONTOR t1, SELSKAP t2 where t1.selskapid = t2.selskapid}  AS Avdelingskontor RELATE 'LokasjonID' TO 'LokasjonID') AS Avdelingskontor"
    
    Set rsLoc = New ADODB.Recordset
    
    Set rsLoc = ObjDAL.FetchRC(mStrXtraDataShapeCon, strSQL)
    
    While Not rsLoc.EOF
        Set objLocation = objContext.CreateInstance("xtraweb.Location")
        objLocation.InitCompanies
        
        For Each fldData In rsLoc.Fields
            objLocation.DataValues(fldData.Name) = fldData.Value
        Next
        strLocID = "ID" & objLocation.DataValues("lokasjonID").Value
        
        
        Set rsComp = rsLoc("Avdelingskontor").Value
        
        While Not rsComp.EOF
            Set objComp = objContext.CreateInstance("xtraweb.Company")
            
            For Each fldData In rsComp.Fields
                objComp.DataValues(fldData.Name) = fldData.Value
            Next
            strCompID = "ID" & objComp.DataValues("ID").Value
            objLocation.Companies.Add objComp, strCompID
            rsComp.MoveNext
        Wend
        Add objLocation, strLocID
        rsLoc.MoveNext
    Wend
    Set rsLoc = Nothing
    Set rsComp = Nothing
    Set ObjDAL = Nothing
    Exit Sub
 
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
End Sub
