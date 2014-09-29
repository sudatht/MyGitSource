VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjDataValues As DataValues
Private mrsCust As ADODB.Recordset
Private mobjAddresses As Addresses
Private mStrXtraCon As String
Private mStrXtraDataShapeCon As String
Private mStrImpCon As String
'The internal Xtra user to impersonate
Private Const CLWebExternalID = "1"
Private Const CStrType = "KUNDE"

Public Property Set Addresses(ByVal vData As Addresses)
    Set mobjAddresses = vData
End Property

Public Property Get Addresses() As Addresses
    Set Addresses = mobjAddresses
    mobjAddresses.Refresh
End Property

Public Property Get XtraConString() As String
    XtraConString = mStrXtraCon
End Property

Public Property Let XtraConString(strConString As String)
    mStrXtraCon = strConString
End Property

Public Property Get XtraDataShapeConString() As String
    XtraDataShapeConString = mStrXtraDataShapeCon
End Property

Public Property Let XtraDataShapeConString(strConString As String)
    mStrXtraDataShapeCon = strConString
End Property

Public Property Get IMPConString() As String
    IMPConString = mStrImpCon
End Property

Public Property Let IMPConString(strConString As String)
    mStrImpCon = strConString
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
    Set mobjAddresses = objContext.CreateInstance("xtraweb.Addresses")
    Set mobjAddresses.Owner = SafeRef(Me)
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjDataValues = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
    Set objContext = Nothing
End Sub

Public Sub CleanUp()
    Set mobjDataValues = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
End Sub

Public Function GetWebRights() As Webrights
On Error GoTo Err_handler
    Dim objRights As Webrights

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No customer has been choosen (datavalues is empty) !"
    Else
        If mobjDataValues("kontaktid").Value < 1 Then
           Err.Raise 702 + vbObjectError, "", " Invalid customer id !"
        End If
    End If
    Set GetWebRights = objContext.CreateInstance("xtraweb.webrights")
    With mobjDataValues
        GetWebRights.DataValues.Add "KontaktID", CStr(.Item("kontaktID").Value)
        GetWebRights.DataValues.Add "TYPE", "KUNDE"
        GetWebRights.DataValues.Add "XtraCon", mStrXtraCon
    End With
    Exit Function
    
Err_handler:
    Set GetWebRights = Nothing
    WriteLog "customer_GetWebRights", Error$, EF_LogEventType_Error
End Function

Public Function GetCustomer(ByVal LContactID As Long) As Boolean
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim RsCustomer As ADODB.Recordset
    Dim fldData As ADODB.Field
    
    On Error GoTo Err_handler
    
    If Trim(Len(mStrXtraCon)) = 0 Then
        GetCustomer = False
        Exit Function
    End If
    
    If LContactID < 1 Then
        Err.Raise 701 + vbObjectError, "", " Invalid customerID! (Less than 1)"
    End If
    
    Set ObjDAL = New ClsDataAccess
    
    strSQL = "Select Distinct F.FirmaID, F.Firma, F.Telefon, F.Fax, F.EPost, F.Hjemmeside, F.StatusID, KontaktPerson=(K.Fornavn + ' '+ K.etternavn), K.kontaktid " & _
        "from Firma F, Kontakt K Where K.kontaktid = " & LContactID & " and K.firmaid = F.firmaid"
    Set RsCustomer = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    If Not RsCustomer.EOF Then
        For Each fldData In RsCustomer.Fields
            mobjDataValues.Add fldData.Name, fldData.Value
        Next
        GetCustomer = True
    Else
        GetCustomer = False
    End If
    'Close & clean up
    Set ObjDAL = Nothing
    Set RsCustomer = Nothing
    Exit Function

Err_handler:
    GetCustomer = False
    WriteLog "customer_GetCustomer", Error$ & strSQL, EF_LogEventType_Error
End Function

Public Function GetGreetings() As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim LCustomerID As Long

    On Error GoTo Err_handler
    
    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "customer_GetGreetings", " No connection string (Xtra) !"
    End If
        
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "customer_GetGreetings", " No customer has been choosen (Lvikarid) !"
    End If
    
    LCustomerID = mobjDataValues("firmaid").Value
    
    If LCustomerID < 1 Then
       Err.Raise 702 + vbObjectError, "customer_GetGreetings", " Invalid customer id !"
    End If
    
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "UserID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LCustomerID
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = New Parameter
    ObjParam.Name = "HilsenFor"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adVarChar
    ObjParam.Value = CStrType
    ObjParam.Size = Len(CStrType)
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = New Parameter
    ObjParam.Name = "Bursdag"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = Null
    ObjCmd.Parameters.Append ObjParam

    Set ObjParam = Nothing
    
    'Fetch recordset
    Set GetGreetings = ObjDAL.FetchRCSP(mStrXtraCon, "Getgreetings", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

Err_handler:
    Set GetGreetings = Nothing
    WriteLog "customer_GetGreetings", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetCourseConfirmations(ByVal dtfradato As Date, ByVal dtTilDato As Date) As ADODB.Recordset
    Dim strSQL As String
    Dim RsVikar As ADODB.Recordset
    Dim rsOppdragVikar As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim BExecOK As Boolean
    Dim LCustomerID As Long
    Dim LContactID As Long

On Error GoTo Err_handler
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra) !"
    End If

    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No customer has been choosen (Lvikarid) !"
    End If
    
    LCustomerID = mobjDataValues("firmaid").Value
    
    If LCustomerID < 1 Then
       Err.Raise 702 + vbObjectError, "", " Invalid customer id !"
    End If

    LContactID = mobjDataValues("kontaktid").Value
    
    If LContactID < 1 Then
       Err.Raise 703 + vbObjectError, "", " Invalid contact id !"
    End If

' Check if input parameters are valid
    If Not IsDate(dtfradato) Then
        Err.Raise 704 + vbObjectError, "", " Method Requires fromdate!"
    End If
    
    If Not IsDate(dtTilDato) Then
        Err.Raise 705 + vbObjectError, "", " Method Requires fromdate!"
    End If

    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command
    
    If Not ObjDAL.ExecuteCRUD(mStrXtraCon, "Delete from BEKREFTELSE_KURS where BrukerID =" & CLWebExternalID & " and FirmaID = " & LCustomerID) Then
        Err.Raise 706 + vbObjectError, "", "Could not delete temporary values from bekreftelse_kurs!"
    End If
    
'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "firmaid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LCustomerID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjParam = New Parameter
    ObjParam.Name = "kontaktID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LContactID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "fradato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtfradato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "tildato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
' Get the customers job information
    Set rsOppdragVikar = ObjDAL.FetchRCSP(mStrXtraCon, "GetCustomerCourseConfirmations", ObjCmd.Parameters)
    If Not IsObject(rsOppdragVikar) Then
        Err.Raise 707 + vbObjectError, "", "Error in GetCustomerCourseConfirmations!"
    End If

' Read all accepted & confirmed jobs for customer and create a post for
'each day in oppdrag..
    While Not rsOppdragVikar.EOF
    
'Prepare command object anew (For each loop)
         Set ObjCmd = Nothing
         Set ObjCmd = New Command
    
'Add input parameters to command object
        Set ObjParam = New Parameter
        ObjParam.Name = "aOppdragVikarID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = rsOppdragVikar("OppdragVikarID")
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                
        Set ObjParam = New Parameter
        ObjParam.Name = "aOppdragID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = rsOppdragVikar("OppdragID")
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
               
        Set ObjParam = New Parameter
        ObjParam.Name = "aVikarID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = rsOppdragVikar("VikarID")
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        Set ObjParam = New Parameter
        ObjParam.Name = "aFirmaID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = LCustomerID
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        Set ObjParam = New Parameter
        ObjParam.Name = "afradato"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adDate
        ObjParam.Value = rsOppdragVikar("Fradato")
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjParam = New Parameter
        ObjParam.Name = "atildato"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adDate
        ObjParam.Value = rsOppdragVikar("TilDato")
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjParam = New Parameter
        ObjParam.Name = "abrukerID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = CLWebExternalID
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
' Stored procedure "Lag_Bekreftelse_kurs" creates a post for each
' day between oppdrag fradato and oppdrag tildato.
        If Not ObjDAL.ExecuteCRUDSP(mStrXtraCon, "Lag_Bekreftelse_kurs", ObjCmd.Parameters) Then
            Err.Raise 708 + vbObjectError, "", "Could not insert values!"
        End If
        
' Get next OPPDRAG_VIKAR
        rsOppdragVikar.MoveNext
    Wend
    
' Close and release recordset
    Set rsOppdragVikar = Nothing
    Set ObjCmd = Nothing
    Set ObjCmd = New Command

    Set ObjParam = New Parameter
    ObjParam.Name = "LCustomerID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LCustomerID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
        
    Set ObjParam = New Parameter
    ObjParam.Name = "LContactID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LContactID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "dtFraDato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtfradato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "dtTilDato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    
' Get the customers job information
    Set rsOppdragVikar = ObjDAL.FetchRCSP(mStrXtraCon, "GetCustomerCourses", ObjCmd.Parameters)
    If Not IsObject(rsOppdragVikar) Then
        Err.Raise 709 + vbObjectError, "", "Error in stored procedure GetCustomerCourses!"
    End If
        
    Set GetCourseConfirmations = rsOppdragVikar
    Set ObjCmd = Nothing
    Set ObjDAL = Nothing
    Exit Function

Err_handler:
    Set GetCourseConfirmations = Nothing
    WriteLog "customer_GetCourseConfirmations", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetTimeSheetStatus(ByVal INoOfWeeks As Integer) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim dtfradato As Date
    Dim dtTilDato As Date
    Dim bParError As Boolean
    Dim bReturnError As Boolean
    Dim LfirmaID As Long
    Dim LContactID As Long
    
    On Error GoTo Err_handler
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "customer_GetTimeSheetStatus", " No connection string (Xtra) !"
    End If
    
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "customer_GetTimeSheetStatus", " No customer has been choosen (LfirmaId) !"
    End If
    
    LfirmaID = mobjDataValues("firmaid").Value
    
    If LfirmaID < 1 Then
       Err.Raise 702 + vbObjectError, "customer_GetTimeSheetStatus", " Invalid customer id !"
    End If
    
    LContactID = mobjDataValues("kontaktid").Value
    
    If LContactID < 1 Then
       Err.Raise 703 + vbObjectError, "customer_GetTimeSheetStatus", " Invalid contact id !"
    End If
    
    bParError = False
    bReturnError = False
    'Legger datoen for siste dag i forrige uke in i dttildato
    dtTilDato = Format(DateAdd("d", -Weekday(Now(), vbMonday), Now()), "mm/dd/yyyy", vbMonday)
    'Kalkulerer startdatoen og legger denne inn i dtfradato
    dtfradato = Format(DateAdd("ww", -INoOfWeeks, DateAdd("d", -6, dtTilDato)), "mm/dd/yyyy", vbMonday)
       
    Set ObjCmd = New ADODB.Command
    
    bParError = True
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "Firmaid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LfirmaID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "KontaktID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LContactID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "fradato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtfradato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "tildato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    bParError = False
    Set ObjDAL = New ClsDataAccess
    
    'Fetch recordset
    bReturnError = True
    Set GetTimeSheetStatus = ObjDAL.FetchRCSP(mStrXtraCon, "GetcustomersNonApprovedtimesheets", ObjCmd.Parameters)
    bReturnError = False
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

Err_handler:
    Set GetTimeSheetStatus = Nothing
    WriteLog "customer_GetTimeSheetStatus", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetTaskConfirmation(ByVal lOppdragVikarID As Long) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim LfirmaID As Long
    Dim rsOppdragVikar As ADODB.Recordset
    Dim LContactID As Long

On Error GoTo Err_handler
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "customer_GetTaskConfirmation", " No connection string (Xtra) !"
    End If
    
    If IsMissing(mobjDataValues) Then
       Err.Raise 701 + vbObjectError, "customer_GetTaskConfirmation", " No customer has been choosen (LfirmaId) !"
    End If
    
    LfirmaID = mobjDataValues("firmaid").Value
        
    If LfirmaID < 1 Then
       Err.Raise 702 + vbObjectError, "customer_GetTaskConfirmation", " Invalid customer id !"
    End If

    LContactID = mobjDataValues("kontaktid").Value
    
    If LContactID < 1 Then
       Err.Raise 703 + vbObjectError, "customer_GetTaskConfirmation", " Invalid contact id !"
    End If

    If lOppdragVikarID < 1 Then
       Err.Raise 704 + vbObjectError, "customer_GetTaskConfirmation", " Method Requires a valid OppdragVikarID!"
    End If
          
    Set ObjCmd = New ADODB.Command

    'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "firmaId"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LfirmaID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "KontaktID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LContactID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "OppdragVikarID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = lOppdragVikarID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
       
    Set ObjDAL = New ClsDataAccess
    
    ' Get the customers task information
    Set GetTaskConfirmation = ObjDAL.FetchRCSP(mStrXtraCon, "GetCustomerTaskConfirmation", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
       
Err_handler:
    Set GetTaskConfirmation = Nothing
    WriteLog "customer_GetTaskConfirmation", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function
   
Public Function GetAllTaskConfirmations(ByVal dtfradato As Date, ByVal dtTilDato As Date) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim LfirmaID As Long
    Dim LContactID As Long

On Error GoTo Err_handler

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "customer_GetAllTaskConfirmations", " No connection string (Xtra) !"
    End If

    If IsMissing(mobjDataValues) Then
       Err.Raise 701 + vbObjectError, "customer_GetAllTaskConfirmations", " No customer has been choosen (LfirmaId) !"
    End If
    LfirmaID = mobjDataValues("firmaid").Value
    
    If LfirmaID < 1 Then
       Err.Raise 702 + vbObjectError, "customer_GetAllTaskConfirmations", " Invalid customer id !"
    End If

    LContactID = mobjDataValues("kontaktid").Value
    
    If LContactID < 1 Then
       Err.Raise 703 + vbObjectError, "customer_GetAllTaskConfirmations", " Invalid contact id !"
    End If
        
    ' Check if input parameters are valid
    If Not IsDate(dtfradato) Then
        Err.Raise 704 + vbObjectError, "customer_GetAllTaskConfirmations", " Method Requires fromdate!"
    End If
    
    If Not IsDate(dtTilDato) Then
        Err.Raise 705 + vbObjectError, "customer_GetAllTaskConfirmations", " Method Requires todate!"
    End If

    Set ObjCmd = New ADODB.Command
    
'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "firmaid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LfirmaID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "kontaktID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LContactID
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "fradato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtfradato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    Set ObjParam = New Parameter
    ObjParam.Name = "tildato"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adDate
    ObjParam.Value = dtTilDato
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
           
    Set ObjDAL = New ClsDataAccess
    
    ' Get the customers task information
    Set GetAllTaskConfirmations = ObjDAL.FetchRCSP(mStrXtraCon, "GetAllcustomerTaskConfirmations", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
       
Err_handler:
    Set GetAllTaskConfirmations = Nothing
    WriteLog "customer_GetAllTaskConfirmations", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function UpdateWebUser(ByVal LImpUserID As Long, ByVal StrNewUserName As String, StrNewPassword As String) As Boolean
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim ObjOutputCmd As ADODB.Command
    
   
    If ((Trim(mStrImpCon) <> "") And (Trim(StrNewUserName) <> "") And (Trim(StrNewPassword) <> "") And (LImpUserID > 0)) Then
        
        Set ObjDAL = New ClsDataAccess
        Set ObjCmd = New ADODB.Command
        
        'Add parameters to command object
            Set ObjParam = New Parameter
            ObjParam.Name = "Userid"
            ObjParam.Direction = adParamInput
            ObjParam.Value = LImpUserID
            ObjParam.Type = adInteger
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "strType"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = CStrType
            ObjParam.Size = Len(CStrType)
            ObjCmd.Parameters.Append ObjParam

            Set ObjParam = New Parameter
            ObjParam.Name = "userName"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = StrNewUserName
            ObjParam.Size = Len(StrNewUserName)
            ObjCmd.Parameters.Append ObjParam
            
            Set ObjParam = New Parameter
            ObjParam.Name = "Password"
            ObjParam.Direction = adParamInput
            ObjParam.Type = adVarChar
            ObjParam.Value = StrNewPassword
            ObjParam.Size = Len(StrNewPassword)
            ObjCmd.Parameters.Append ObjParam
                                        
        ' Update Web information
        If ObjDAL.ExecuteCRUDSP(mStrImpCon, "UpdateUsernamePassword", ObjCmd.Parameters) Then
            UpdateWebUser = True
        Else
            UpdateWebUser = False
        End If
        Set ObjDAL = Nothing
        Set ObjCmd = Nothing
    Else
        UpdateWebUser = False
    End If
End Function

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

