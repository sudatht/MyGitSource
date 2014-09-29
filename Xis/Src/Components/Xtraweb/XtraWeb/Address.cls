VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mobjOwner As Object
Private mstrTable As String

Public Property Set Owner(ByVal vData As Object)
    '##BD Sets or gets the Owner (Suspect or Consultant) for the Address object.
    
    Set mobjOwner = vData
    
    'Check for type of Owner. Then sets the appropriate address table.
    If TypeOf mobjOwner Is Consultant Then
        mstrTable = "adresse"
    'ElseIf TypeOf mobjOwner Is Customer Then
    '    mstrTable = "adresse"
    ElseIf TypeOf mobjOwner Is Suspect Then
        mstrTable = "v_suspect_adresse"
    End If
    
    'Set mobjDataValues.Owner = SafeRef(Me)
End Property

Public Property Get Owner() As Object
    Set Owner = mobjOwner
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    '##BD This property gets or sets the %DataValues% collection
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
Attribute DataValues.VB_UserMemId = 0
    Set DataValues = mobjDataValues
End Property

Public Function Save() As Boolean
'##BD Saves the Address object to the 'Addresse' or 'v_suspect_address' _
      table in the database dependent on the %Address_Owner% property. Returns true _
      if the operation is successfull.
      
'on error GoTo err_Save

    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim strColName As String
    Dim lAdrID As Long
    
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " Owner has no connection string (Xtra)!"
    End If

    Set ObjDAL = New ClsDataAccess
    
    Set rs = New ADODB.Recordset
    
    ' Check if owner is Suspect or Consultant. Then sets the column name _
      for Owners ID.
    If TypeOf mobjOwner Is Suspect Then
        strColName = "suspectid"
    ElseIf TypeOf mobjOwner Is Consultant Then
        strColName = "VikarID"
    ElseIf TypeOf mobjOwner Is Customer Then
        strColName = "FirmaID"
    End If
    
    If ExistsInDataValues(mobjDataValues, "adrid") Then
        lAdrID = mobjDataValues("AdrID").Value
    End If
    
    strSQL = "Select * from " & mstrTable & " where AdrID = " & lAdrID
    
    
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    
    If rs.RecordCount = 0 Then
        'new
        rs.AddNew
        ' Set the right value to the foreign key column
        rs("AdresseRelID") = mobjOwner.DataValues(strColName).Value
    End If
    
    'Loop through the DataValues collection and assign each _
     value to the right column in the recordset
    For Each objValue In mobjDataValues
        ' Check for null values because theese won't go into the _
          recordset
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "adrid" Then
                If LCase(objValue.ValueName) <> "adresserelid" Then
                    rs(objValue.ValueName) = objValue.Value
                End If
            End If
        End If
    Next
    
    If ObjDAL.BatchUpdate(mobjOwner.XtraConString, rs) Then
        Save = True
    End If
    rs.Close
    Set rs = Nothing
    Set ObjDAL = Nothing
    Exit Function
    
err_Save:
    Save = False
    WriteLog "Address_Save", Error$, EF_LogEventType_Error
End Function

Public Function Delete() As Boolean
'##BD Deletes the Address object from the 'Addresse' or 'v_suspect_address' _
      table in the database dependent on the %Address_Owner% property. Returns true _
      if the operation is successfull.

On Error GoTo err_Delete
    Dim ObjDAL As IdataAccess
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lAdrID As Long
        
    If Trim(Len(mobjOwner.XtraConString)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " Owner has no connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess
    Set rs = New ADODB.Recordset
    
    ' Get the unique AdrID value
    If ExistsInDataValues(mobjDataValues, "adrid") Then
        lAdrID = mobjDataValues("AdrID").Value
    End If
        
    'Get the specific record from the database
    strSQL = "Select * from " & mstrTable & " where AdrID = " & lAdrID
    Set rs = ObjDAL.FetchRC(mobjOwner.XtraConString, strSQL)
    
    ' Check if it really exist
    If Not rs.EOF Then
        rs.Delete
        If ObjDAL.BatchUpdate(mobjOwner.XtraConString, rs) Then
            'All went well
            Delete = True
        End If
        rs.Close
    End If
    Set rs = Nothing
    Set ObjDAL = Nothing
    Exit Function
    
err_Delete:
    Delete = False
    WriteLog "Address_Delete", Error$, EF_LogEventType_Error
End Function

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    'Set mobjDataValues.Owner = Nothing
    Set mobjDataValues = Nothing
    Set mobjOwner = Nothing
    Set objContext = Nothing
End Sub
