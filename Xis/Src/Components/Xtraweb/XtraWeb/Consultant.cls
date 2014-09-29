VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Consultant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mobjCV As CV
Private mobjChangedCV As CV
Private mobjEmployee As Employee
Private mobjDataValues As DataValues
Private mobjAddresses As Addresses
Private MobjInvoiceMonth As InvoiceCustomerMonth
Private mStrXtraCon As String
Private mStrXtraDataShapeCon As String
Private mStrImpCon As String
'The internal Xtra user to impersonate
Private Const CLWebExternalID = "1"
Private Const CStrType = "ANSATT"

Public Property Set InvoiceMonth(ByVal vObjInvoiceMonth As InvoiceCustomerMonth)
    Set MobjInvoiceMonth = vObjInvoiceMonth
End Property

Public Property Get InvoiceMonth() As InvoiceCustomerMonth
    Set InvoiceMonth = MobjInvoiceMonth
End Property

Public Property Set Addresses(ByVal vData As Addresses)
    Set mobjAddresses = vData
End Property

Public Property Get Addresses() As Addresses
    Set Addresses = mobjAddresses
    mobjAddresses.Refresh
End Property

Public Property Set DataValues(ByVal vData As DataValues)
    Set mobjDataValues = vData
End Property

Public Property Get DataValues() As DataValues
    Set DataValues = mobjDataValues
End Property

Public Property Get Employee() As Employee
    If mobjEmployee Is Nothing Then
        Set mobjEmployee = objContext.CreateInstance("xtraweb.Employee")
    End If
    Set Employee = mobjEmployee
End Property

Public Property Set Employee(vData As Employee)
    Set mobjEmployee = vData
End Property

Public Property Get CV() As CV
    If mobjCV Is Nothing Then
        Set mobjCV = objContext.CreateInstance("xtraweb.CV")
        Set mobjCV.Owner = SafeRef(Me)
    End If
    mobjCV.XtraConString = mStrXtraCon
    mobjCV.Refresh
    Set CV = mobjCV
End Property

Public Property Set CV(vData As CV)
    Set mobjCV = vData
End Property

Public Property Get ChangedCV() As CV
    If mobjChangedCV Is Nothing Then
        Set mobjChangedCV = objContext.CreateInstance("xtraweb.CV")
        Set mobjChangedCV.Owner = SafeRef(Me)
    End If
    mobjChangedCV.XtraConString = mStrXtraCon
    mobjChangedCV.RefreshChanged
    Set ChangedCV = mobjChangedCV
End Property

Public Property Set ChangedCV(vData As CV)
    Set mobjChangedCV = vData
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

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mobjEmployee = objContext.CreateInstance("xtraweb.Employee")
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    Set mobjAddresses = objContext.CreateInstance("xtraweb.Addresses")
    Set mobjAddresses.Owner = SafeRef(Me)
End Sub

Public Sub CleanUp()
    If Not mobjCV Is Nothing Then
        Set mobjCV.Owner = Nothing
        Set mobjCV = Nothing
    End If
    If Not mobjChangedCV Is Nothing Then
        Set mobjChangedCV.Owner = Nothing
        Set mobjChangedCV = Nothing
    End If
    Set mobjEmployee = Nothing
    Set mobjDataValues = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
    Set MobjInvoiceMonth = Nothing
    Set objContext = Nothing
End Sub

Private Sub ObjectControl_Deactivate()
    Set mobjCV.Owner = Nothing
    Set mobjCV = Nothing
    Set mobjEmployee = Nothing
    Set mobjDataValues = Nothing
    Set mobjAddresses.Owner = Nothing
    Set mobjAddresses = Nothing
    Set MobjInvoiceMonth = Nothing
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Public Function GetWebRights() As Webrights
On Error GoTo Err_handler
    Dim objRights As Webrights
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If

    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No consultant has been choosen (datavalues is empty) !"
    Else
        If mobjDataValues("vikarid").Value < 1 Then
           Err.Raise 702 + vbObjectError, "", " Invalid consultant id !"
        End If
    End If
    Set GetWebRights = objContext.CreateInstance("xtraweb.webrights")
    With mobjDataValues
        GetWebRights.DataValues.Add "VikarID", CStr(.Item("Vikarid").Value)
        GetWebRights.DataValues.Add "TYPE", "VIKAR"
        GetWebRights.DataValues.Add "XtraCon", mStrXtraCon
    End With
    Exit Function
    
Err_handler:
    Set GetWebRights = Nothing
    WriteLog "consultant_GetWebRights", Error$, EF_LogEventType_Error
End Function

Public Function GetConsultant(ByVal lConsultantID As Long) As Boolean
    Dim ObjDAL As IdataAccess
    Dim strSQL As String
    Dim rsCons As ADODB.Recordset
    Dim fldData As ADODB.Field
           
    On Error GoTo Err_handler
           
    If Trim(Len(mStrXtraCon)) = 0 Then
        GetConsultant = False
        Exit Function
    End If
    
    If lConsultantID < 1 Then
        Err.Raise 701 + vbObjectError, "", " Invalid lConsultantID! (Less than 1)"
    End If
    
    Set ObjDAL = New ClsDataAccess
   
    strSQL = "Select * from vikar where vikarid = " & lConsultantID
    Set rsCons = ObjDAL.FetchRC(mStrXtraCon, strSQL)
    Set ObjDAL = Nothing
    If Not rsCons.EOF Then
        For Each fldData In rsCons.Fields
            DataValues.Add fldData.Name, fldData.Value
        Next
        GetConsultant = True
    Else
        GetConsultant = False
    End If
    'Close & clean up
    Set ObjDAL = Nothing
    Set rsCons = Nothing
    Exit Function
    
Err_handler:
    GetConsultant = False
    WriteLog "customer_GetConsultant", Error$ & strSQL, EF_LogEventType_Error
End Function

Public Function GetCourseConfirmations(ByVal dtfradato As Date, ByVal dtTilDato As Date) As ADODB.Recordset
    Dim strSQL As String
    Dim RsVikar As ADODB.Recordset
    Dim rsOppdragVikar As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim BExecOK As Boolean
    Dim Lvikarid As Long

On Error GoTo Err_handler
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No consultant has been choosen (Lvikarid) !"
    End If
    Lvikarid = mobjDataValues("vikarid").Value
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "", " Invalid consultant id !"
    End If

    ' Check if input parameters are valid
    If Not IsDate(dtfradato) Then
        Err.Raise 703 + vbObjectError, "", " Method Requires fromdate!"
    End If
    If Not IsDate(dtTilDato) Then
        Err.Raise 704 + vbObjectError, "", " Method Requires todate!"
    End If
    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command
    
    If Not ObjDAL.ExecuteCRUD(mStrXtraCon, "Delete from BEKREFTELSE_KURS where BrukerID = " & CLWebExternalID & " and VikarID = " & Lvikarid) Then
        Err.Raise 705 + vbObjectError, "", "Could not delete temporary values from bekreftelse_kurs!"
    End If
'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    
    ' Get the Consultants job information
    Set rsOppdragVikar = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantCourseConfirmations", ObjCmd.Parameters)
    If rsOppdragVikar Is Nothing Then
        Err.Raise 706 + vbObjectError, "", "Error in GetConsultantCourseConfirmations!"
    End If
    ' Read all accepted & confirmed jobs for consultant and create a post for
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
        ObjParam.Value = rsOppdragVikar("OppdragVikarID").Value
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
    
        Set ObjParam = New Parameter
        ObjParam.Name = "aOppdragID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = rsOppdragVikar("OppdragID").Value
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjParam = New Parameter
        ObjParam.Name = "aVikarID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = Lvikarid
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        Set ObjParam = New Parameter
        ObjParam.Name = "aFirmaID"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = rsOppdragVikar("FirmaID").Value
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing

        Set ObjParam = New Parameter
        ObjParam.Name = "afradato"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adDate
        ObjParam.Value = rsOppdragVikar("Fradato").Value
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
            Err.Raise 707 + vbObjectError, "", "Could not insert values!"
        End If
        
        ' Get next OPPDRAG_VIKAR
        rsOppdragVikar.MoveNext
    Wend
    
    ' Close and release recordset
    Set rsOppdragVikar = Nothing
    'Prepare command object yet anew
    Set ObjCmd = Nothing
    Set ObjCmd = New Command

    ' Get all details
    Set ObjParam = New Parameter
    ObjParam.Name = "Lvikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
        
    Set rsOppdragVikar = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantCourses", ObjCmd.Parameters)

    If rsOppdragVikar Is Nothing Then
        Err.Raise 708 + vbObjectError, "", "Could not retrieve confirmations.."
    End If
    Set GetCourseConfirmations = rsOppdragVikar
    Set ObjDAL = Nothing
    Exit Function

Err_handler:
    Set GetCourseConfirmations = Nothing
    WriteLog "consultant_GetCourseConfirmations", Error$, EF_LogEventType_Error
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


Public Function GetGreetings() As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long
    
    On Error GoTo Err_handler
    
    Set ObjCmd = New ADODB.Command
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "consultant_GetGreetings", " No connection string (Xtra) !"
    End If
    
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "consultant_GetGreetings", " No consultant has been choosen (Lvikarid) !"
    End If
    
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "consultant_GetGreetings", " Invalid consultant id !"
    End If

    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "UserID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    
    Set ObjDAL = New ClsDataAccess
    'Fetch recordset
    Set GetGreetings = ObjDAL.FetchRCSP(mStrXtraCon, "Getgreetings", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

Err_handler:
    Set GetGreetings = Nothing
    WriteLog "consultant_GetGreetings", Error$, EF_LogEventType_Error
End Function

Public Function GetContact() As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long
    On Error GoTo Err_handler
    
    Set ObjDAL = New ClsDataAccess
    Set ObjCmd = New ADODB.Command
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "consultant_GetContact", " No connection string (Xtra) !"
    End If
        
    If mobjDataValues.Count < 1 Then
       Err.Raise 701 + vbObjectError, "consultant_GetContact", " No consultant has been choosen (Lvikarid) !"
    End If
    
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "consultant_GetContact", " Invalid consultant id !"
    End If

    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "Vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
    ObjCmd.Parameters.Append ObjParam
    Set ObjParam = Nothing
    
    'Fetch recordset
    Set GetContact = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantContact", ObjCmd.Parameters)
    If GetContact.EOF Then
        Set GetContact = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantAvdkontor", ObjCmd.Parameters)
    End If
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

Err_handler:
    Set GetContact = Nothing
    WriteLog "consultant_GetContact", Error$, EF_LogEventType_Error
End Function

Public Function GetTimeSheetStatus(ByVal INoOfWeeks As Integer) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim dtfradato As Date
    Dim dtTilDato As Date
    Dim bParError As Boolean
    Dim bReturnError As Boolean
    Dim Lvikarid As Long
    
    On Error GoTo Err_handler
    
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 701 + vbObjectError, "", " No connection string (Xtra) !"
    End If
    
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No consultant has been choosen (Lvikarid) !"
    End If
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "", " Invalid consultant id !"
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
    ObjParam.Name = "UserID"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    Set GetTimeSheetStatus = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantsNonApprovedtimesheets", ObjCmd.Parameters)
    bReturnError = False
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function

Err_handler:
    Set GetTimeSheetStatus = Nothing
    WriteLog "consultant_GetTimeSheetStatus", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function

Public Function GetTaskConfirmation(ByVal lOppdragVikarID As Long) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long
    Dim rsOppdragVikar As ADODB.Recordset

On Error GoTo Err_handler
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra) !"
    End If
    
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No consultant has been choosen (Lvikarid) !"
    End If
    
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "", " Invalid consultant id !"
    End If
   
    If lOppdragVikarID < 1 Then
       Err.Raise 703 + vbObjectError, "", " Method Requires a valid OppdragVikarID!"
    End If


    Set ObjCmd = New ADODB.Command

'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    
    ' Get the Consultants task information
    Set GetTaskConfirmation = ObjDAL.FetchRCSP(mStrXtraCon, "GetConsultantTaskConfirmation", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
       
Err_handler:
    Set GetTaskConfirmation = Nothing
    WriteLog "consultant_GetTaskConfirmation", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function
   
Public Function GetAllTaskConfirmations(ByVal dtfradato As Date, ByVal dtTilDato As Date) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long

On Error GoTo Err_handler

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "consultant_GetAllTaskConfirmations", " No connection string (Xtra) !"
    End If
    
    If IsMissing(mobjDataValues) Then
       Err.Raise 701 + vbObjectError, "consultant_GetAllTaskConfirmations", " No consultant has been choosen (Lvikarid) !"
    End If
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "consultant_GetAllTaskConfirmations", " Invalid consultant id !"
    End If

    'Get and verify input pars.
    If Not IsDate(dtfradato) Then
        Err.Raise 703 + vbObjectError, "consultant_GetAllTaskConfirmations", " Invalid fromdate!"
    End If
    
    If Not IsDate(dtTilDato) Then
        Err.Raise 704 + vbObjectError, "consultant_GetAllTaskConfirmations", " Invalid todate!"
    End If

    Set ObjCmd = New ADODB.Command
    
'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    
    ' Get the Consultants task information
    Set GetAllTaskConfirmations = ObjDAL.FetchRCSP(mStrXtraCon, "GetAllConsultantTaskConfirmations", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
       
Err_handler:
    Set GetAllTaskConfirmations = Nothing
    WriteLog "consultant_GetAllTaskConfirmations ", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function


Public Function Save() As Boolean
On Error GoTo err_Save
    '##BD Lagrer objektet. Lagrer både endringer og nye objekter.
    Dim fldData As ADODB.Field
    Dim ObjDAL As IdataAccess
    Dim objValue As XtraWeb.DataValue
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lConsultantID As Long

    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra)!"
    End If
    
    Set ObjDAL = New ClsDataAccess

    Set rs = New ADODB.Recordset

    'check if new
    If ExistsInDataValues(mobjDataValues, "vikarid") Then
        'exists
        lConsultantID = mobjDataValues("vikarid").Value
    End If

    strSQL = "Select * from vikar where vikarid = " & lConsultantID
    Set rs = ObjDAL.FetchRC(mStrXtraCon, strSQL)

    If rs.RecordCount = 0 Then
        'new
        rs.AddNew
    End If
           
    For Each objValue In mobjDataValues
        If Not IsNull(objValue.Value) Then
            If LCase(objValue.ValueName) <> "vikarid" Then
                rs(objValue.ValueName) = objValue.Value
            End If
        End If
    Next objValue
    
    rs("regdato") = Date
    ObjDAL.BatchUpdate mStrXtraCon, rs

    ' Refresh all the values from the database
    Set mobjDataValues = objContext.CreateInstance("xtraweb.DataValues")
    For Each fldData In rs.Fields
        mobjDataValues(fldData.Name) = fldData.Value
    Next
    Save = True
    Set rs = Nothing
    rs.Close
    Set ObjDAL = Nothing
    
    Exit Function
err_Save:
    WriteLog "Consultant_Save", Error$, EF_LogEventType_Error
    Save = False
End Function

Public Function GetEcTaskConfirmation(ByVal lOppdragVikarID As Long) As ADODB.Recordset
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim Lvikarid As Long
    Dim rsOppdragVikar As ADODB.Recordset

On Error GoTo Err_handler
    If Trim(Len(mStrXtraCon)) = 0 Then
       Err.Raise 700 + vbObjectError, "", " No connection string (Xtra) !"
    End If
    
    If mobjDataValues.Count = 0 Then
       Err.Raise 701 + vbObjectError, "", " No consultant has been choosen (Lvikarid) !"
    End If
    
    Lvikarid = mobjDataValues("vikarid").Value
    
    If Lvikarid < 1 Then
       Err.Raise 702 + vbObjectError, "", " Invalid consultant id !"
    End If
   
    If lOppdragVikarID < 1 Then
       Err.Raise 703 + vbObjectError, "", " Method Requires a valid OppdragVikarID!"
    End If


    Set ObjCmd = New ADODB.Command

'Add input parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "vikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
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
    
    ' Get the Consultants task information
    Set GetEcTaskConfirmation = ObjDAL.FetchRCSP(mStrXtraCon, "spECGetConsultantTaskConfirmation", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
    Exit Function
       
Err_handler:
    Set GetEcTaskConfirmation = Nothing
    WriteLog "consultant_GetTaskConfirmation", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Function
