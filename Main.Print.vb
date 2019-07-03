Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain

    Private Sub ProcessReportDataTable(ByVal OutputType As Integer)
        If Not gbGroupSort Then
            Call ProcessDataTableByNonGroup(OutputType)
        Else
            Call ProcessDataTableByGroup(OutputType)
        End If
    End Sub

    Private Function OpenOutput(ByVal OutputType As Integer) As Boolean
        OpenOutput = False

        OpenOutput = OpenCsv()
    End Function

    Private Function OpenCsv() As Boolean
        OpenCsv = False

        giFileNumber = FreeFile()
        Try
            FileOpen(giFileNumber, gsTempOutputName, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)
        Catch ex As Exception
            Call LogToFile("  Error: OpenCsv - Unable to open file (" & gsTempOutputName & ") - " & ex.Message)
            Exit Function
        End Try

        If gbCsvHeader Then
            Call CsvHeader()
        End If

        OpenCsv = True

    End Function

    Private Sub CsvHeader()
        Dim sRecord As String = ""

        ' Header
        'sRecord = "AccCode" & gsCsvDelimiter & "Dest" & gsCsvDelimiter & "ActStat" & gsCsvDelimiter & "LastAct" & _
        '  gsCsvDelimiter & "B/S" & gsCsvDelimiter & "SecCode" & gsCsvDelimiter & "OrdPrc" & gsCsvDelimiter & "PrcInst" & _
        '  gsCsvDelimiter & "Lifetime" & gsCsvDelimiter & "OrdVol" & gsCsvDelimiter & "DoneVolTot" & gsCsvDelimiter & _
        '  "RemVol" & gsCsvDelimiter & "AvgPrc" & gsCsvDelimiter &  "AccType" & gsCsvDelimiter & "ExecInstr" & _
        '  gsCsvDelimiter & "PostTradeStatus" & gsCsvDelimiter & "OrdNo" & gsCsvDelimiter & "PriCliOrd"


        'sRecord = "OrdNo" & gsCsvDelimiter
        sRecord = sRecord & "cusipCode" & gsCsvDelimiter
        sRecord = sRecord & "exchange" & gsCsvDelimiter
        sRecord = sRecord & "exchangeRate" & gsCsvDelimiter
        sRecord = sRecord & "executingBroker" & gsCsvDelimiter
        sRecord = sRecord & "grossPrice" & gsCsvDelimiter
        sRecord = sRecord & "isinCode" & gsCsvDelimiter
        sRecord = sRecord & "masterAccount" & gsCsvDelimiter
        sRecord = sRecord & "netPrice" & gsCsvDelimiter
        sRecord = sRecord & "openClose" & gsCsvDelimiter
        sRecord = sRecord & "orignalPrice" & gsCsvDelimiter
        sRecord = sRecord & "quantity" & gsCsvDelimiter
        'sRecord = sRecord & "tradedQty" & gsCsvDelimiter
        sRecord = sRecord & "sedolCode" & gsCsvDelimiter
        sRecord = sRecord & "settlementCcy" & gsCsvDelimiter
        sRecord = sRecord & "settlementDate" & gsCsvDelimiter
        sRecord = sRecord & "side" & gsCsvDelimiter
        sRecord = sRecord & "symbol" & gsCsvDelimiter
        sRecord = sRecord & "tradeCcy" & gsCsvDelimiter
        sRecord = sRecord & "tradeDate" & gsCsvDelimiter
        sRecord = sRecord & "tradeID" & gsCsvDelimiter
        sRecord = sRecord & "description" & gsCsvDelimiter
        'sRecord = sRecord & "securityType" & gsCsvDelimiter
        sRecord = sRecord & "type" & gsCsvDelimiter
        'sRecord = sRecord & "accCode" & gsCsvDelimiter
        sRecord = sRecord & "account" & gsCsvDelimiter
        'sRecord = sRecord & "accType" & gsCsvDelimiter
        sRecord = sRecord & "type" & gsCsvDelimiter
        'sRecord = sRecord & "ExecInstr"
        sRecord = sRecord & "DFANotes1" & gsCsvDelimiter
        sRecord = sRecord & "DFANotes2" & gsCsvDelimiter
        sRecord = sRecord & "DFANotes3" & gsCsvDelimiter
        sRecord = sRecord & "SecExtra" & gsCsvDelimiter
        sRecord = sRecord & "Source" & gsCsvDelimiter
        'CurrencyDenomination
        sRecord = sRecord & "CurrencyDenomination"

        PrintLine(giFileNumber, sRecord)
    End Sub

    Private Sub ProcessDataTableByNonGroup(ByVal OutputType As Integer)
        Dim dvReportOrdersView As DataView
        Dim lIndex As Long
        Dim drReportOrderRow As Data.DataRowView

        Try
            dvReportOrdersView = New DataView(dtReportOrdersTable)
            dvReportOrdersView.Sort = gsSortOrder
        Catch ex As Exception
            Call LogToFile("  Error: ProcessDataTableByNonGroup (OutputType" & CStr(OutputType) &
              ") - Unable to create DataView - " & ex.Message)
            Exit Sub
        End Try

        If OutputType = 0 Then
            ReDim DataArray(MAX_ARRAY_ROWS - 1, MAX_ARRAY_COLUMNS - 1)
        End If

        For lIndex = 0 To dvReportOrdersView.Count - 1
            Try
                drReportOrderRow = dvReportOrdersView.Item(lIndex)
                Call PrintRecord(drReportOrderRow, OutputType)
            Catch ex As Exception
                Call LogToFile("  Error: ProcessDataTableByNonGroup (OutputType" & CStr(OutputType) &
                  ") - Unable to read DataView row (" & CStr(lIndex) & ") - " & ex.Message)
            End Try
        Next
    End Sub

    Private Sub PrintRecord(ByVal ReportOrderRow As Data.DataRowView, ByVal OutputType As Integer)

        Call PrintCsvRecord(ReportOrderRow("AccCode"), ReportOrderRow("Dest"), ReportOrderRow("ActStat"),
                            ReportOrderRow("LastAct"), ReportOrderRow("BuySell"), ReportOrderRow("SecCode"), ReportOrderRow("OrdPrc"), ReportOrderRow("TrdPrc"),
                            ReportOrderRow("PrcInst"), ReportOrderRow("Lifetime"), ReportOrderRow("OrdVol"), ReportOrderRow("TrdVol"), ReportOrderRow("DoneVolTot"),
                            ReportOrderRow("RemVol"), ReportOrderRow("AvgPrc"), ReportOrderRow("AccType"), ReportOrderRow("ExecInstr"),
                            ReportOrderRow("PostTradeStatus"), ReportOrderRow("OrdNo"), ReportOrderRow("TradeNo"), ReportOrderRow("PriCliOrd"), ReportOrderRow("EXBR"),
                            ReportOrderRow("SettlementCurrency"), ReportOrderRow("FXrate"), ReportOrderRow("Organization"),
                            ReportOrderRow("TradeTime"), ReportOrderRow("SettlementTime"), ReportOrderRow("SecurityType"),
                            ReportOrderRow("Currency"), ReportOrderRow("Exchange"), ReportOrderRow("Cusip"), ReportOrderRow("ISIN"),
                            ReportOrderRow("SEDOL"), ReportOrderRow("Description"), ReportOrderRow("DFANote1"), ReportOrderRow("DFANote2"),
                            ReportOrderRow("DFANote3"), ReportOrderRow("OpenClose"), ReportOrderRow("UkIrishStampDutyReserveTaxMarker"), ReportOrderRow("PtmLevyIndicator"), ReportOrderRow("Source"), ReportOrderRow("CurrencyDenomination"))

    End Sub

    Private Sub PrintCsvRecord(ByVal AccCode As String, ByVal Dest As String, ByVal ActStat As String,
      ByVal LastAct As String, ByVal BuySell As String, ByVal SecCode As String, ByVal OrdPrc As Decimal, ByVal TrdPrc As Decimal,
      ByVal PrcInst As String, ByVal Lifetime As String, ByVal OrdVol As Long, ByVal TrdVol As Long, ByVal DoneVolTot As Long,
      ByVal RemVol As Long, ByVal AvgPrc As Decimal, ByVal AccType As String, ByVal ExecInstr As String,
      ByVal PostTradeStatus As String, ByVal OrdNo As Long, ByVal TradeNo As Long, ByVal PriCliOrd As String, ByVal EXBR As String,
                               ByVal SettlementCurrency As String, ByVal FXrate As String, ByVal Organization As String,
                               ByVal TradeTime As String, ByVal SettlementTime As String, ByVal SecurityType As String,
                               ByVal Currency As String, ByVal Exchange As String, ByVal Cusip As String, ByVal ISIN As String,
                               ByVal SEDOL As String, ByVal Description As String, ByVal DFANote1 As String, ByVal DFANote2 As String,
                               ByVal DFANote3 As String, ByVal OpenClose As String, ByVal UkIrishStampDutyReserveTaxMarker As String, ByVal PtmLevyIndicator As String, ByVal Source As String, ByVal CurrencyDenomination As String)
        Dim sRecord As String
        Dim SecExtra As String
        Dim sUkIrishStampDutyReserveTaxMarker As String
        Dim sPtmLevyIndicator As String

        'sRecord = CStr(OrdNo) & gsCsvDelimiter
        'Cusip
        sRecord = sRecord & Cusip & gsCsvDelimiter

        sRecord = sRecord & Exchange & gsCsvDelimiter
        sRecord = sRecord & FXrate & gsCsvDelimiter
        sRecord = sRecord & EXBR & gsCsvDelimiter

        'RLU003
        'sRecord = sRecord & CStr(AvgPrc) & gsCsvDelimiter
        sRecord = sRecord & CStr(TrdPrc) & gsCsvDelimiter

        'ISIN
        sRecord = sRecord & ISIN & gsCsvDelimiter

        sRecord = sRecord & """" & Organization & """" & gsCsvDelimiter

        'RLU003
        'sRecord = sRecord & CStr(AvgPrc) & gsCsvDelimiter
        sRecord = sRecord & CStr(TrdPrc) & gsCsvDelimiter

        'OpenClose
        sRecord = sRecord & OpenClose & gsCsvDelimiter

        'RLU003
        'sRecord = sRecord & CStr(AvgPrc) & gsCsvDelimiter
        sRecord = sRecord & CStr(TrdPrc) & gsCsvDelimiter


        'RLU003
        'sRecord = sRecord & CStr(OrdVol) & gsCsvDelimiter
        If gbNetting Then
            sRecord = sRecord & CStr(TrdVol) & gsCsvDelimiter
        Else
            sRecord = sRecord & CStr(OrdVol) & gsCsvDelimiter
        End If




        'sRecord = sRecord & CStr(DoneVolTot) & gsCsvDelimiter
        'SEDOL
        sRecord = sRecord & SEDOL & gsCsvDelimiter

        sRecord = sRecord & SettlementCurrency & gsCsvDelimiter
        sRecord = sRecord & SettlementTime & gsCsvDelimiter
        sRecord = sRecord & BuySell & gsCsvDelimiter
        sRecord = sRecord & SecCode & gsCsvDelimiter
        'Trade currency
        sRecord = sRecord & Currency & gsCsvDelimiter

        sRecord = sRecord & TradeTime & gsCsvDelimiter

        'RLU003 MultiTrades
        'sRecord = sRecord & CStr(OrdNo) & "-" & CStr(TradeNo) & gsCsvDelimiter

        sRecord = sRecord & CStr(OrdNo) & gsCsvDelimiter

        'Description
        sRecord = sRecord & """" & Description & """" & gsCsvDelimiter

        sRecord = sRecord & SecurityType & gsCsvDelimiter
        sRecord = sRecord & AccCode & gsCsvDelimiter
        sRecord = sRecord & AccType & gsCsvDelimiter

        sRecord = sRecord & DFANote1 & gsCsvDelimiter
        sRecord = sRecord & DFANote2 & gsCsvDelimiter
        sRecord = sRecord & DFANote3 & gsCsvDelimiter
        'sRecord = sRecord & """" & ExecInstr & """" & gsCsvDelimiter
        'RL002
        'sUkIrishStampDutyReserveTaxMarker = "UkIrishStampDutyReserveTaxMarker" & "(" & UkIrishStampDutyReserveTaxMarker & ")"

        'sPtmLevyIndicator = "PtmLevyIndicator" & "(" & PtmLevyIndicator & ")"


        sUkIrishStampDutyReserveTaxMarker = UkIrishStampDutyReserveTaxMarker
        sPtmLevyIndicator = PtmLevyIndicator

        SecExtra = sUkIrishStampDutyReserveTaxMarker & "," & sPtmLevyIndicator

        If SecExtra.StartsWith(",") Or SecExtra.EndsWith(",") Then
            SecExtra = SecExtra.Replace(",", "")
        Else
            'RLU003
            SecExtra = "STAMPPTMEXEMPT"
        End If

        ''RLU003
        'If sUkIrishStampDutyReserveTaxMarker <> "" And sPtmLevyIndicator <> "" Then
        '    SecExtra = "STAMPPTMEXEMPT"
        'End If


        SecExtra = """" & SecExtra & """"

        'SecExtra = """" & sUkIrishStampDutyReserveTaxMarker & "," & sPtmLevyIndicator & """"

        sRecord = sRecord & SecExtra & gsCsvDelimiter
        CurrencyDenomination = """" & CurrencyDenomination & """"
        sRecord = sRecord & Source & gsCsvDelimiter
        sRecord = sRecord & CurrencyDenomination

        PrintLine(giFileNumber, sRecord)
    End Sub

    Private Sub ProcessDataTableByGroup(ByVal OutputType As Integer)
        If OutputType = 0 Then
            ReDim DataArray(MAX_ARRAY_ROWS - 1, MAX_ARRAY_COLUMNS - 1)
        End If

        Call ProcessGroup(OutputType, "RemVol = 0", "Fills")
        Call ProcessGroup(OutputType, "DoneVolTot = 0", "Open Orders")
        Call ProcessGroup(OutputType, "RemVol <> 0 AND DoneVolTot <> 0", "Partial Fills")
    End Sub

    Private Sub ProcessGroup(ByVal OutputType As Integer, ByVal Filter As String, ByVal Header As String)
        Dim dvReportOrdersView As DataView
        Dim lIndex As Long
        Dim drReportOrderRow As Data.DataRowView

        Try
            dvReportOrdersView = New DataView(dtReportOrdersTable)
            dvReportOrdersView.RowFilter = Filter
            dvReportOrdersView.Sort = gsSortOrder
        Catch ex As Exception
            Call LogToFile("  Error: ProcessGroup (OutputType" & CStr(OutputType) & " - Filter (" & Filter & _
              ") - Unable to create DataView - " & ex.Message)
            Exit Sub
        End Try

        If dvReportOrdersView.Count > 0 Then
            If OutputType = 0 Then

            Else
                PrintLine(giFileNumber, "")
                PrintLine(giFileNumber, Header)
            End If
        End If

        For lIndex = 0 To dvReportOrdersView.Count - 1
            Try
                drReportOrderRow = dvReportOrdersView.Item(lIndex)
                Call PrintRecord(drReportOrderRow, OutputType)
            Catch ex As Exception
                Call LogToFile("  Error: ProcessGroup (OutputType" & CStr(OutputType) & " - Filter (" & Filter & _
                  ") - Unable to read DataView row (" & CStr(lIndex) & ") - " & ex.Message)
            End Try
        Next
    End Sub

    Private Sub CloseOutput(ByVal OutputType As Integer)
        FileClose(giFileNumber)

    End Sub
End Class