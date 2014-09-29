VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Invoicebasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstrconnection As String

Public Function GetInvoice(ByVal strConnection As String, ByRef strPeriod As String, ByRef InvoiceOwner As Object) As InvoiceMonth
    Dim IYear As Integer
    Dim IMonth As Integer
    Dim ObjDAL As IdataAccess
    Dim ObjCmd As ADODB.Command
    Dim ObjParam As ADODB.Parameter
    Dim RcInvoiceData As Recordset
    Dim fldCol As Field
    Dim ObjDay As InvoiceDay
    Dim OBjWeek As InvoiceWeek
    Dim Objmonth As InvoiceMonth
    Dim IPrevWeek As Integer
    Dim ICurrentWeek As Integer
    
    Dim ukeSumUbet As Double
    Dim gmlUkesumUbet As Double
    Dim ukeSumUtbetalt As Double
    Dim sumDag As Double
    Dim tempSumDag As Double
    Dim WeekNosumUbet As Double
    Dim korrBelop As Double
    Dim sumTotIkkeUtbet As Double
       
    Dim dtCurrentDay As Date
    Dim Dtaar As Date
    Dim Iaar As Integer

    Dim Luke As Long
    Dim Iukedag As Integer
    Dim IUkedagNr As Integer
    Dim IgmlUkedagNr As Integer
    Dim IDagnr As Integer
    Dim IukenrMnd As Integer
    Dim strSplittmerke As String
    Dim StrStatus As String
    Dim IaarMinus As Long
    Dim iTmpWeekNo As Integer
    Dim LCurrentPeriod As Long
    Dim BLStart As Boolean
    Dim BlUkeTest  As Boolean
    Dim PersonId As Long
    Dim dtPrevDay As Date
    
    'Validate the in-parameters
    On Error GoTo Err_handler
    
    If Len(Trim(strConnection)) = 0 Then
        Err.Raise vbObjectError + 709, "invoicebasis", "Missing strConnection parameter."
    Else
        mstrconnection = strConnection
    End If
    
    If Len(Trim(strPeriod)) = 0 Then
        Err.Raise vbObjectError + 710, "invoicebasis", "Missing Period parameter."
    End If

    If Len(Trim(strPeriod)) <> 6 Then
        Err.Raise vbObjectError + 711, "invoicebasis", "strPeriod has a invalid parameter value."
    End If
    
    If Not IsObject(InvoiceOwner) Then
        Err.Raise vbObjectError + 712, "invoicebasis", "Missing Invoiceowner parameter."
    End If
    
    If ((Not TypeOf InvoiceOwner Is Consultant) And (Not TypeOf InvoiceOwner Is Customer)) Then
        Err.Raise vbObjectError + 713, "invoicebasis", "Invoiceowner parameter is of wrong type."
    End If
    
    'Check & transfer the parameters to variables..
    IYear = Val(Left(strPeriod, 4))
    If IYear > Year(Now()) Then
        Err.Raise vbObjectError + 714, "invoicebasis", "strPeriod consists of a invalid year."
    End If
    
    IMonth = Val(Right(strPeriod, 2))
    If IMonth > 12 Then
        Err.Raise vbObjectError + 714, "invoicebasis", "strPeriod consists of a invalid month."
    End If
    
    
    'Retreive the invoice data for consultant
    Set ObjCmd = New ADODB.Command

    If TypeOf InvoiceOwner Is Consultant Then
        
        PersonId = InvoiceOwner.DataValues("vikarid").Value
        
        'Add input parameters to command object
        Set ObjParam = New Parameter
        ObjParam.Name = "vikarid"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = PersonId
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
        
        Set ObjParam = New Parameter
        ObjParam.Name = "aar"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = IYear
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
           
        Set ObjParam = New Parameter
        ObjParam.Name = "maaned"
        ObjParam.Direction = adParamInput
        ObjParam.Type = adInteger
        ObjParam.Value = IMonth
        ObjCmd.Parameters.Append ObjParam
        Set ObjParam = Nothing
                      
        Set ObjDAL = CreateObject("DataAccess.ClsDataAccess")
        
        ' Get the Consultants invoice basis
        Set RcInvoiceData = ObjDAL.FetchRCSP(mstrconnection, "GetConsultantInvoiceBasis", ObjCmd.Parameters)
        If Not IsObject(RcInvoiceData) Then
            Err.Raise "704" + vbObjectError, "Error in or during execution of stored procedure GetConsultantInvoiceBasis!"
        End If
         
        'Initialize vars
        BLStart = True
        ukeSumUbet = 0
        gmlUkesumUbet = 0
        ukeSumUtbetalt = 0
         
        'Build the various collections and classes..
        Set Objmonth = New InvoiceMonth
        Set OBjWeek = New InvoiceWeek
        While Not RcInvoiceData.EOF
            Set ObjDay = New InvoiceDay
            'For transer all values for each day..
            For Each fldCol In RcInvoiceData.Fields
                ObjDay.DataValues.Add fldCol.Name, fldCol.Value
            Next
            
            'Converts the "uke" field from "20001" to "200001" format
            Luke = CLng(RcInvoiceData("uke"))
            If Luke < 100000 Then
                Luke = CLng(Left(Luke, 4) & "0" & Right(Luke, 1))
            End If
            
            Iukedag = RcInvoiceData("Ukedag")
            IUkedagNr = RcInvoiceData("Ukedag")
            
            If ((RcInvoiceData("splittuke") <> "") And (Not IsNull(RcInvoiceData("splittuke")))) Then
                strSplittmerke = " - " & RcInvoiceData("splittuke")
            End If
            
            StrStatus = RcInvoiceData("status")
            sumDag = (RcInvoiceData("antTimer") * RcInvoiceData("Timelonn"))
            tempSumDag = sumDag
            
            Select Case StrStatus
                Case "utbetalt"
                    sumDag = 0
                Case Else
                    ukeSumUbet = ukeSumUbet + sumDag
            End Select
            
            'for å kompansere for feil ukenr i kalender...
            '*********************************************
            Dtaar = RcInvoiceData("dato")
            
            If (DatePart("d", Dtaar) < 8) And (DatePart("m", Dtaar) = 1) And (Not DatePart("ww", Dtaar, vbMonday, vbFirstFourDays) = 1) Then
                IaarMinus = 1
            Else
                IaarMinus = 0
            End If
            
            Iaar = DatePart("yyyy", Dtaar, vbMonday, vbFirstFourDays) - IaarMinus
            iTmpWeekNo = DatePart("ww", Dtaar, vbMonday, vbFirstFourDays)
            
            If iTmpWeekNo < 10 Then
                LCurrentPeriod = CLng(Iaar & "0" & iTmpWeekNo)
            Else
                LCurrentPeriod = Iaar & iTmpWeekNo
            End If
                      
            ICurrentWeek = Format(CDate(ObjDay.DataValues("dato").Value), "ww", vbMonday)
            'If first day..
            If IPrevWeek = 0 Then
                IPrevWeek = ICurrentWeek
            End If
            'Check to see if new week..
            If ICurrentWeek <> IPrevWeek Then
                If BLStart = False Then
                    BlUkeTest = False
                    korrBelop = 0
                                                                
                    Call GetConsultantOvertime(PersonId, IPrevWeek, ukeSumUtbetalt, dtPrevDay)
                    
                    If RcInvoiceData("status") <> "utbetalt" Then
                        gmlUkesumUbet = ukeSumUbet - sumDag + korrBelop
                        sumTotIkkeUtbet = sumTotIkkeUtbet + gmlUkesumUbet
                    Else
                        gmlUkesumUbet = 0
                    End If
                    OBjWeek.SumIkkeUtebetalt = FormatNumber(gmlUkesumUbet, 2)
                    OBjWeek.WeekNo = IPrevWeek
                    'ukeSumUbet = sumDag
                End If
                Objmonth.Add OBjWeek
                Set OBjWeek = New InvoiceWeek
            End If
            ObjDay.SumUtbetalt = FormatNumber(tempSumDag, 2)

            dtPrevDay = RcInvoiceData("dato")
            
'            igmlUke = Luke
            BlUkeTest = True
            BLStart = False
    
            IPrevWeek = ICurrentWeek
            OBjWeek.Add ObjDay

            RcInvoiceData.MoveNext
        Wend
        
        korrBelop = 0
        
        Call GetConsultantOvertime(PersonId, IPrevWeek, ukeSumUtbetalt, dtPrevDay)
                       
        gmlUkesumUbet = ukeSumUbet + korrBelop

        sumTotIkkeUtbet = sumTotIkkeUtbet + gmlUkesumUbet
                
        Objmonth.SumtotaltUtbetalt = FormatNumber(isumTotal, 2)
        Objmonth.SumTotaltIkkeUtebetalt = FormatNumber(sumTotIkkeUtbet, 2)
            
        'Clean up temp objects
        RcInvoiceData.Close
        Set RcInvoiceData = Nothing
        Set GetInvoice = Objmonth
        Set Objmonth = Nothing
        'return the collection..
    Else
    
    End If
    
    Exit Function
Err_handler:
    Set GetInvoice = Nothing
    WriteLog "invoicebasis_getinvoice", Error$, EF_LogEventType_Error
End Function



Private Sub GetConsultantOvertime(ByVal vikarid As Long, ByVal WeekNo As Integer, ByRef ukeSumUtbetalt As Double, ByVal Dag As Integer)
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim ObjDAL As IdataAccess
    Dim rsOvertid As ADODB.Recordset
    Dim overtidBelop As Double
    Dim periode As Integer
    Dim StrUkedel As String
    Dim WeekNosumUbet As Double
    Dim korrBelop As Double
    Dim StrSplittkorr As String
    Dim sisteUkeMnd As Integer
    Dim strSQL As String

    'Gets day in month
    IDagnr = DatePart("d", Dag, vbMonday)
    
    'Gets day in week
    IUkedagNr = DatePart("w", Dag, vbMonday)
    
    'Gets weeknr of the first week of month
    IukenrMnd = DatePart("ww", DateAdd("d", -IDagnr + 1, Dag), vbMonday, vbFirstFourDays)
    
    StrSplittkorr = ""

    'If day in month is in the first week of the month and
    'it is the weeknr of the first week of month.
    If IDagnr < 8 And IukenrMnd = DatePart("ww", Dag, 2, 2) Then
        StrSplittkorr = " and not notat like '1' "
    End If

    'Brilliant naming
    'Gets the first day of month
    a = DateAdd("d", -IDagnr + 1, Dag)

    'Gets the first day of the next month
    b = DateAdd("m", 1, a)

    'Gets the last day of the last week
    c = DatePart("d", DateAdd("d", -8, b), vbMonday, vbFirstFourDays)
    
    'Gets the weeknr of the last week of the month
    sisteUkeMnd = DatePart("ww", DateAdd("d", -1, b), vbMonday, vbFirstFourDays)
    'Gets the weeknr of the current date
    IukenrMnd = DatePart("ww", Dag, vbMonday, vbFirstFourDays)
     
    
    If IDagnr > c And IukenrMnd = sisteUkeMnd Then
        StrSplittkorr = " and not notat like '2' "
    End If
                                
    strSQL = "Select overtidType= " & _
        " CASE Loennsartnr " & _
        " WHEN 160 THEN 50 " & _
        " WHEN 161 THEN 50 " & _
        " WHEN 162 THEN 50 " & _
        " WHEN 163 THEN 100 " & _
        " WHEN 164 THEN 100 " & _
        " WHEN 165 THEN 100 " & _
        " END, " & _
        " OppdragID, Antall, Sats, Belop, loennperiode, overfort_loenn_status, notat " & _
        " from VIKAR_UKELISTE " & _
        " where vikarid =" & vikarid & _
        " and Loennsartnr >=160 " & _
        " and Loennsartnr <=165 " & _
        " and Ukenr = " & WeekNo & _
        StrSplittkorr & _
        " order by loennperiode, loennsartnr, overfort_loenn_status "

    Set ObjDAL = CreateObject("DataAccess.ClsDataAccess")
    Set rsOvertid = ObjDAL.FetchRC(mstrconnection, strSQL)

    If Not rsOvertid.EOF Then 'hvis det er overtid registrert
            
        Do Until rsOvertid.EOF

            overtidBelop = rsOvertid("Belop")
            periode = "" & rsOvertid("loennperiode")
            StrUkedel = ""
            
            If Not Trim(periode) = "" Then
                ukeSumUtbetalt = ukeSumUtbetalt + overtidBelop
            End If
            
            If Trim(periode) = "" Then
                WeekNosumUbet = WeekNosumUbet + overtidBelop
                periode = "ikke lønnet"
                If rsOvertid("overtidType") = 50 Then
                    korrBelop = korrBelop + overtidBelop / 3
                End If
                If rsOvertid("overtidType") = 100 Then
                    korrBelop = korrBelop + overtidBelop / 2
                End If
            End If
            
            If Not rsOvertid("notat") = " " Then
                StrUkedel = "- Ukedel " & rsOvertid("notat")
            End If
                
            'Str = "Overtid " & rsOvertid("overtidType") & StrUkedel
            'Str2 = rsOvertid("oppdragID") & "Periode:" & periode & rsOvertid("Antall")
            'str3 = FormatNumber(overtidBelop, 2)


        rsOvertid.MoveNext
        Loop
    End If
    rsOvertid.Close
    Set rsOvertid = Nothing

End Sub


Sub GetConsultantWeeksum(vikarid, gmlUke, nyPeriode, ukeSumUtbetalt, Dag)
'ByVal vikarid As Long, ByVal WeekNo As Integer, ByRef ukeSumUtbetalt As Double, ByVal Dag As Integer)
              
    'Gets day in month
    IDagnr = DatePart("d", Dag, vbMonday)
    
    'Gets day in week
    IUkedagNr = DatePart("w", Dag, vbMonday)
    
    'Gets weeknr of the first week of month
    IukenrMnd = DatePart("ww", DateAdd("d", -IDagnr + 1, Dag), vbMonday, vbFirstFourDays)
    
    StrSplittkorr = ""

    'If day in month is in the first week of the month and
    'it is the weeknr of the first week of month.
    If IDagnr < 8 And IukenrMnd = DatePart("ww", Dag, 2, 2) Then
        StrSplittkorr = " and not notat like '1' "
    End If

    'Brilliant naming
    'Gets the first day of month
    a = DateAdd("d", -IDagnr + 1, Dag)

    'Gets the first day of the next month
    b = DateAdd("m", 1, a)

    'Gets the last day of the last week
    c = DatePart("d", DateAdd("d", -8, b), vbMonday, vbFirstFourDays)
    
    'Gets the weeknr of the last week of the month
    sisteUkeMnd = DatePart("ww", DateAdd("d", -1, b), vbMonday, vbFirstFourDays)
    'Gets the weeknr of the current date
    IukenrMnd = DatePart("ww", Dag, vbMonday, vbFirstFourDays)
         
    If IDagnr > c And IukenrMnd = sisteUkeMnd Then
        StrSplittkorr = " and not notat like '2' "
    End If


    strSQL = "Select  OppdragID, Antall, Sats, Belop,Loennsartnr, loennperiode, notat, overfort_loenn_status " & _
    " from VIKAR_UKELISTE " & _
    " where vikarid =" & vikarNR & _
    " and Loennsartnr in(26,32,33,34,38,150,151,152,153,154,155,156,157,158,159)" & _
    " and Ukenr = " & gmlUke & _
    Splittkorr & _
    " order by Loennsartnr, loennperiode "

    Set ObjDAL = CreateObject("DataAccess.ClsDataAccess")
    Set rsUkeloenn = ObjDAL.FetchRC(mstrconnection, strSQL)

        If Not rsUkeloenn.EOF Then

            oppdragID = 0
            loennsart = 0
            antall = 0
            sats = 0
            belop = 0
            gmlLart = 0
            gmlOpdrID = 0
            gmlSats = 0
        
            Do Until rsUkeloenn.EOF
                loennperiode = "" & rsUkeloenn("loennperiode")
                loennsart = rsUkeloenn("Loennsartnr")
                oppdragID = rsUkeloenn("oppdragID")
                antall = rsUkeloenn("Antall")
                sats = rsUkeloenn("Sats")
                belop = rsUkeloenn("Belop")
                ukedel = ""
                loennstatus = rsUkeloenn("overfort_loenn_status")
                
                If rsUkeloenn("notat") <> " " Then
                    ukedel = "-" & rsUkeloenn("notat")
                End If
                    
                If Not Trim(loennperiode) = "" Then
                    ukeSumUtbetalt = ukeSumUtbetalt + rsUkeloenn("Belop")
                End If
                
                If loennperiode = "" Then
                    If loennstatus > 2 Then
                        loennperiode = "ukjent"
                    Else
                        loennperiode = "ikke utbetalt"
                    End If
                End If
                    
                If belop <> 0 Then
                    Call skrivUtUkesum(loennsart, gmlUke, loennperiode, oppdragID, antall, sats, belop, ukedel)
                End If
                    
                gmlLart = loennsart
                gmlOpdrID = oppdragID
                gmlSats = sats
            
            rsUkeloenn.MoveNext
            Loop
        End If
        rsUkeloenn.Close
        Set rsUkeloenn = Nothing
        sumTotal = sumTotal + ukeSumUtbetalt
        Call skrivUtUkesumTotal(gmlUke, gmlOpdrID, ukeSumUtbetalt, loennperiode)
        ukeSumUtbetalt = 0
End Sub
