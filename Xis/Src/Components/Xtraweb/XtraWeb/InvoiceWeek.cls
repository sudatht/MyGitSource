VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "InvoiceWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'MTS
Implements ObjectControl
Private objContext As ObjectContext

Private mCol As Collection
Private mobjInvoice As Invoice
Private mObjOvertimes As InvoiceOvertimes
Private mObjWeekSums As InvoiceWeekSums

Public Property Get WeekNo() As Integer
    If mCol.Count > 0 Then
        WeekNo = DatePart("ww", mCol.Item(1).DataValues("dagdato").Value, vbMonday, vbFirstFourDays)
    Else
        WeekNo = 0
    End If
End Property

Public Property Set WeekSums(ByVal vData As InvoiceWeekSums)
    Set mObjWeekSums = vData
End Property

Public Property Get WeekSums() As InvoiceWeekSums
    Set WeekSums = mObjWeekSums
End Property

Public Property Set Overtimes(ByVal vData As InvoiceOvertimes)
    Set mObjOvertimes = vData
End Property

Public Property Get Overtimes() As InvoiceOvertimes
    Set Overtimes = mObjOvertimes
End Property

Public Property Get Owner() As Invoice
    Set Owner = mobjInvoice
End Property

Public Property Set Owner(ByRef objInvoice As Invoice)
    Set mobjInvoice = objInvoice
End Property

Public Function Add(ByVal ObjDay As InvoiceDay) As InvoiceDay
    mCol.Add ObjDay
    Set Add = ObjDay
    Set ObjDay = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As InvoiceDay
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Refresh()
    'summer nytt fakturagrunnlag for uka
    If TypeOf mobjInvoice.Owner Is Consultant Then
        Call GetConsultantWeekOvertime
        Call GetConsultantWeeksum
    End If
End Sub

Private Sub ObjectControl_Activate()
    Set objContext = GetObjectContext()
    Set mCol = New Collection
    Set mObjOvertimes = objContext.CreateInstance("xtraweb.InvoiceOvertimes")
    Set mObjWeekSums = objContext.CreateInstance("xtraweb.InvoiceWeekSums")
End Sub
Private Sub ObjectControl_Deactivate()
    Set mCol = Nothing
    Set mobjInvoice = Nothing
    Set mObjOvertimes = Nothing
    Set mObjWeekSums = Nothing
End Sub

Private Sub GetConsultantWeekOvertime()
    
    Dim a As Date
    Dim b As Date
    Dim c As Integer
    Dim ObjDAL As IdataAccess
    Dim rsOvertid As ADODB.Recordset
    Dim overtidBelop As Double
    Dim Periode As String
    Dim dblkorrBelop As Double
    Dim StrSplittkorr As String
    Dim sisteUkeMnd As Integer
    Dim strSQL As String
    Dim dtDay As Date
    Dim LYearWeekNo As Long
    Dim Lvikarid As Long
    Dim IdagNr As Integer
    Dim IUkedagNr As Integer
    Dim IukenrMnd As Integer
    Dim ukeSumUtbetalt As Double
    Dim ICurrentWeekNo As Integer
    Dim ObjOvertime As InvoiceOvertime

On Error GoTo Err_handler

    If mCol.Count = 0 Then
        Exit Sub
    End If

    If mobjInvoice.Owner Is Nothing Then
        Exit Sub
    End If

    dtDay = mCol.Item(1).DataValues("dagdato").Value
    LYearWeekNo = DatePart("yyyy", dtDay, vbMonday, vbFirstFourDays) & leftpad(CStr(DatePart("ww", dtDay, vbMonday, vbFirstFourDays)), 2, "0")
       
    ICurrentWeekNo = DatePart("ww", dtDay, vbMonday, vbFirstFourDays)
        
    If TypeOf mobjInvoice.Owner Is Consultant Then
        Lvikarid = mobjInvoice.Owner.DataValues("vikarid").Value
    ElseIf TypeOf mobjInvoice.Owner Is Customer Then
        Lvikarid = Me.Item(1).DataValues("vikarid")
    End If

    'Gets day in month
    IdagNr = DatePart("d", dtDay, vbMonday)

    'Gets day in week
    IUkedagNr = DatePart("w", dtDay, vbMonday)

    'Gets weeknr of the first week of previous month
    IukenrMnd = DatePart("ww", DateAdd("d", -IdagNr + 1, dtDay), vbMonday, vbFirstFourDays)
    StrSplittkorr = ""

    'If day in month is in the first week of the month and
    'it is the weeknr of the first week of month.
    If IdagNr < 8 And IukenrMnd = ICurrentWeekNo Then
        StrSplittkorr = "and not notat like '1' "
    End If

    'Gets the first day of month
    a = DateAdd("d", -IdagNr + 1, dtDay)

    'Gets the first day of the next month
    b = DateAdd("m", 1, a)

    'Gets the last day of the last week
    c = DatePart("d", DateAdd("d", -8, b), vbMonday, vbFirstFourDays)

    'Gets the weeknr of the last week of the month
    sisteUkeMnd = DatePart("ww", DateAdd("d", -1, b), vbMonday, vbFirstFourDays)
    IukenrMnd = DatePart("ww", dtDay, vbMonday, vbFirstFourDays)

    If IdagNr > c And IukenrMnd = sisteUkeMnd Then
        StrSplittkorr = "and not notat like '2' "
    End If

    strSQL = "SELECT overtidType = (H.LoennRate*100)-100, OppdragID, Antall, Sats, Belop, loennperiode = isnull(loennperiode,''), overfort_loenn_status, notat " & _
        " FROM VIKAR_UKELISTE AS V " & _
        " INNER JOIN H_loennsArt AS H ON V.loennsartnr = H.loennsartnr " & _
        " WHERE vikarid = " & Lvikarid & _
        " AND H.LoennRate > 1.0 " & _
        " AND Ukenr = " & LYearWeekNo & " " & _
        StrSplittkorr & _
        " ORDER BY loennperiode, H.LoennRate, overfort_loenn_status "

    Set ObjDAL = New ClsDataAccess
    'Henter Overtidsgrunnlag for denne uka
    Set rsOvertid = ObjDAL.FetchRC(mobjInvoice.StrConnection, strSQL)
    Set ObjDAL = Nothing
    'hvis det er overtid registrert
    If Not rsOvertid.EOF Then
        Do Until rsOvertid.EOF
            overtidBelop = rsOvertid("Belop")
            Periode = CStr(rsOvertid("loennperiode"))

            'Ingen periode er registrert
            If Trim(Periode) = "" Then
                Periode = "ikke lønnet"
                dblkorrBelop = overtidBelop / (rsOvertid("overtidType") * 100) * ((rsOvertid("overtidType") * 100) - 100)
            End If

            Set ObjOvertime = objContext.CreateInstance("xtraweb.InvoiceOvertime")
            ObjOvertime.DataValues.Add rsOvertid("overtidType").Name, rsOvertid("overtidType").Value
            ObjOvertime.DataValues.Add rsOvertid("oppdragID").Name, rsOvertid("oppdragID").Value
            ObjOvertime.DataValues.Add rsOvertid("loennperiode").Name, Periode
            ObjOvertime.DataValues.Add rsOvertid("Antall").Name, rsOvertid("Antall").Value
            ObjOvertime.DataValues.Add "overtidBelop", FormatNumber(overtidBelop, 2)
            ObjOvertime.DataValues.Add "sats", FormatNumber(rsOvertid("sats").Value, 2)
            ObjOvertime.DataValues.Add "korrBelop", FormatNumber(dblkorrBelop, 2)
            mObjOvertimes.Add ObjOvertime

        rsOvertid.MoveNext
        Loop
    End If
    Set rsOvertid = Nothing
    Set ObjDAL = Nothing
    Exit Sub

Err_handler:
    WriteLog "invoiceWeek_getConsultantWeekOvertime", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
End Sub

Private Sub GetConsultantWeeksum()
    'Temp var used to calculate if there is a splitt week
    Dim dta                     As Date
    Dim dtb                     As Date
    'Data access object used to execute SQL, stored procedures and so forth.
    Dim ObjDAL                  As IdataAccess

    Dim ISplittkorr             As Integer
    'Date to retrieve the weekly sums for.
    Dim dtDay                   As Date
    'Holds the current period in the format YYYYMM (199912)
    Dim LYearWeekNo             As Long
    'The consultant id to retreive the weekly sums for.
    Dim Lvikarid                As Long
    'Holds the day in month 1 - 31.
    Dim IdagNr                  As Integer
    'Holds weekday 1 - 7.
    Dim IUkedagNr               As Integer
    'Holds the weeknumber (1 - 52) of the previous month
    Dim IukenrMnd               As Integer
    Dim DblukeSumUtbetalt       As Double
    Dim DblIkkeUtbetBelop       As Double
    Dim DblAntallTimerTot       As Double
    Dim DblukeSumIkkeUtbetalt   As Double
    Dim DblPrevSats             As Double
    Dim dblAntall               As Double
    Dim DblSats                 As Double
    Dim dblBelop                As Double
    Dim prvStrLoennperiode      As String
    Dim strLoennperiode         As String
    Dim ic                      As Integer
    Dim sisteUkeMnd             As Integer
    Dim ICurrentWeekNo          As Integer
    Dim Lloennstatus            As Long
    Dim prvLLoennstatus         As Long
    Dim rsUkeloenn              As ADODB.Recordset
    Dim ObjCmd                  As ADODB.Command
    Dim ObjParam                As ADODB.Parameter
    Dim ObjWeekSum              As InvoiceWeeksum
    Dim ObjOvertime             As InvoiceOvertime
    Dim BlnFirstRun             As Boolean

    On Error GoTo Err_handler
    
    If mCol.Count = 0 Then
        'raise err
        Exit Sub
    End If

    If mobjInvoice.Owner Is Nothing Then
        'raise err
        Exit Sub
    End If

    dtDay = mCol.Item(1).DataValues("dagdato").Value
    LYearWeekNo = DatePart("yyyy", dtDay, vbMonday, vbFirstFourDays) & leftpad(DatePart("ww", dtDay, vbMonday, vbFirstFourDays), 2, "0")
    ICurrentWeekNo = DatePart("ww", dtDay, vbMonday, vbFirstFourDays)
    Lvikarid = mobjInvoice.Owner.DataValues("vikarid").Value

    'Gets day in month
    IdagNr = DatePart("d", dtDay, vbMonday)

    'Gets day in week
    IUkedagNr = DatePart("w", dtDay, vbMonday)

    'Gets weeknr of the first week of previous month
    IukenrMnd = DatePart("ww", DateAdd("d", -IdagNr + 1, dtDay), vbMonday, vbFirstFourDays)
    sisteUkeMnd = DatePart("ww", DateAdd("d", -1, dtb), 2, 2)

    ISplittkorr = 0

    'If day in month is in the first week of the month and
    'it is the weeknr of the first week of month.
    If IdagNr < 8 And IukenrMnd = ICurrentWeekNo Then
        ISplittkorr = 1
    End If

    'Brilliant naming
    'Gets the first day of month
    dta = DateAdd("d", -IdagNr + 1, dtDay)

    'Gets the first day of the next month
    dtb = DateAdd("m", 1, dta)

    'Gets the last day of the last week
    ic = DatePart("d", DateAdd("d", -8, dtb), vbMonday, vbFirstFourDays)

    'Gets the weeknr of the last week of the month
    sisteUkeMnd = DatePart("ww", DateAdd("d", -1, dtb), vbMonday, vbFirstFourDays)
    IukenrMnd = DatePart("ww", dtDay, vbMonday, vbFirstFourDays)
    If IdagNr > ic And IukenrMnd = sisteUkeMnd Then
        ISplittkorr = 2
    End If

    Set ObjCmd = New ADODB.Command
    'Add parameters to command object
    Set ObjParam = New Parameter
    ObjParam.Name = "intvikarid"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = Lvikarid
    ObjCmd.Parameters.Append ObjParam
       
    Set ObjParam = New Parameter
    ObjParam.Name = "intYearWeekNo"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = LYearWeekNo
    ObjCmd.Parameters.Append ObjParam
        
    Set ObjParam = New Parameter
    ObjParam.Name = "intSplit"
    ObjParam.Direction = adParamInput
    ObjParam.Type = adInteger
    ObjParam.Value = ISplittkorr
    ObjCmd.Parameters.Append ObjParam
    
    Set ObjParam = Nothing
    
    'Fetch recordset
    Set ObjDAL = New ClsDataAccess
    Set rsUkeloenn = ObjDAL.FetchRCSP(mobjInvoice.StrConnection, "GetConsultantSumFormWeek", ObjCmd.Parameters)
    'Clean up
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing

    If Not rsUkeloenn.EOF Then
        BlnFirstRun = True
        Do Until rsUkeloenn.EOF
            If IsNull(rsUkeloenn("loennperiode")) Then
                strLoennperiode = ""
            Else
                strLoennperiode = CStr(rsUkeloenn("loennperiode"))
            End If

            dblAntall = rsUkeloenn("Antall")
            DblSats = rsUkeloenn("Sats")
            dblBelop = rsUkeloenn("Belop")
            Lloennstatus = rsUkeloenn("overfort_loenn_status")

            If Trim(strLoennperiode) = "" Then
                If Lloennstatus < 3 Then
                    strLoennperiode = "ikke utbetalt"
                    DblIkkeUtbetBelop = dblBelop
                    dblBelop = 0
                End If
            End If

            'Etterslengs adding..
            If (((DblPrevSats <> DblSats) Or (prvLLoennstatus <> Lloennstatus)) And (Not BlnFirstRun)) Then
                Set ObjWeekSum = objContext.CreateInstance("xtraweb.InvoiceWeeksum")
                With ObjWeekSum.DataValues
                    .Add "periode", prvStrLoennperiode
                    .Add "antall", DblAntallTimerTot
                    .Add "sats", DblPrevSats
                    .Add "belop", DblukeSumUtbetalt
                    .Add "belopIkkeUtbet", DblukeSumIkkeUtbetalt
                    .Add "Loennstatus", prvLLoennstatus
                End With
                mObjWeekSums.Add ObjWeekSum
                DblAntallTimerTot = 0
                DblukeSumUtbetalt = 0
                DblukeSumIkkeUtbetalt = 0
            End If
            DblAntallTimerTot = DblAntallTimerTot + dblAntall
            DblukeSumUtbetalt = DblukeSumUtbetalt + dblBelop
            DblukeSumIkkeUtbetalt = DblukeSumIkkeUtbetalt + DblIkkeUtbetBelop

            BlnFirstRun = False
            DblPrevSats = DblSats
            prvLLoennstatus = Lloennstatus
            prvStrLoennperiode = strLoennperiode
            rsUkeloenn.MoveNext
        Loop
        Set ObjWeekSum = objContext.CreateInstance("xtraweb.InvoiceWeeksum")
        With ObjWeekSum.DataValues
            .Add "periode", prvStrLoennperiode
            .Add "antall", DblAntallTimerTot
            .Add "sats", DblPrevSats
            .Add "belop", DblukeSumUtbetalt
            .Add "belopIkkeUtbet", DblukeSumIkkeUtbetalt
            .Add "Loennstatus", prvLLoennstatus
        End With
        mObjWeekSums.Add ObjWeekSum
    End If
    Set rsUkeloenn = Nothing
    Set ObjDAL = Nothing
    Exit Sub

Err_handler:
    WriteLog "invoiceWeek_getConsultantweekSum", Error$, EF_LogEventType_Error
    Set ObjDAL = Nothing
    Set ObjCmd = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

