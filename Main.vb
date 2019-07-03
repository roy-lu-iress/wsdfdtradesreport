
'-----------------------------------------------------------------------------------------------
' Program               :  WSOrdersReport
'
' Description           :  This program reports orders from IOS databases through IRESS
'                          Webservices and writes them in an Excel and/or csv format.
'
' Operating System      :  32-bit Windows with .Net support
'
' Author(s)             : 
'
' Change Activity       :
'
' Author  Date        Version   Tag       Description
' ------  ----        -------   ---       -----------
' RLU     19-FEB-19   1.2.1     RL011     Fix bust order.
' RLU     15-JAN-19   1.1.09    RL010     Add previous day start time
' KCHE    21-Dec-18   1.1.08    KC005     Fix for settlement dates during holidays.
'                               KC006     Make settlement and trade date format configurable and cultureinvariant.
' KCHE    19-Dec-18   1.1.07    KC004     Report exchange timestamp of trades.
' KCHE    18-Dec-18   1.1.07    KC001     Retain history of previous processed trades.
'                               KC002     Add previous day end time
'                               KC003     Add T+3 settlement for exchanges.
'Roy      12/12/18    0.0.16    RL009     Currency and Price MultiPlier setting for exchanges.
'Roy      12/12/18    0.0.15    RL008     Fixed of null Currency Denomination
'Roy      12/05/18    0.0.14    RL007     Add in SecInfoEx.CurrencyDenomination and extra report column
'Roy      12/05/18    0.0.13    RL006     Fixed date issue
'Roy      12/05/18    0.0.13    RL005     Remove the currency multiplication for minor currencies (amount is currently being divided by 100).
'Roy      12/05/18    0.0.13    RL004     Include double quotes to the security description.
'Roy      07/24/18    0.0.12    RL003     Remove all DFD filter
'Roy      07/19/18    0.0.3     RL002     New column UkIrishStampDutyReserveTaxMarker PtmLevyIndicator
'Roy      07/05/18    0.0.2     RL001     Type Should Be Client or Market
'Roy      07/05/18    0.0.2     RL001     Short sells should be reported as SSL


' Add Service Reference: Under Project - Add Service Reference
' - In the Address box, type http://corpowa:82/v4/wsdl.aspx?un=webservices&cp=iress&pw=Iressete3&svc=IOSPlus&svr=ALGOTEST&SH=1
'   and click on Go button.
' - Set the Namespace to "IOSPlus" and hit the OK button.
'
' - Change buffer sizes in app.config: 
'   maxBufferSize="2147483647"
'   maxReceivedMessageSize="2147483647"
'   maxNameTableCharCount="2147483647"

' Add Reference: Under Tab COM - Microsoft Excel 12.0 Object Library
'   To run on server with Excel 2003 - Add 2003 Interop.Excel.dll in 
'   Project -> WSOrdersReport Properties -> References -> Add -> Browse

Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization

Public Class frmMain

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Initialize()
        Call ReadIniFile()
        Call MainProcess()
        Call ApplicationClose()
    End Sub

    Private Sub Initialize()
        gsAppPath = Environment.CurrentDirectory
        gdDate = Today

        gsIniFile = gsAppPath & "\" & APP_NAME & ".ini"
        gsLogFile = gsAppPath & "\" & APP_NAME & ".log"
        gsBackupLog = gsAppPath & "\" & APP_NAME & "log.bak"

        Call BackupLog()
        Call LogToFile("Start of process")

        Call ParseCommandLine()
        If gsInstanceName = "" Then
            Call LogToFile("  Info: Initialize - Instance name not specified (with -r)")
            gbReadConfigFromINI = True
        Else
            Call ReadRegistry()
        End If
    End Sub

    Private Sub ReadIniFile()
        Call ReadIniSource()
        Call ReadIniOutput()
        Call ReadIniOptions()
        Call ReadIniReportAccounts()
    End Sub

    Private Sub MainProcess()
        Call CreateDataTables()
        Call ReadSource()
        Call CreateReport()
    End Sub

  Private Sub ReadSource()
    Dim iNum As Integer

    For iNum = 0 To (MAX_SOURCES - 1)
      If gaSourcesList(iNum).DatabaseType = "WS" Then
        ' Read WS
        giIressSession = iNum

        ' KC001
        If gsReportingType = "DELTA" Then
          Call ReadProcessedTrades(giIressSession)
        End If

        CreateIRESSSession(giIressSession)
        If CreateIOSPlusSession(giIressSession) Then

          If CreateIOSPlusService(giIressSession) Then

            Call ReadWSIOSOrders(giIressSession)
                        'Call FilterNewOrders()

            'Call GetWebServiceIOSAuditTrailByUser(gdStartTime.AddMonths(-4), Now)

            ' KC001
            ' Call GetTradeByOrderNo()
            Call GetTradeByOrderNo(giIressSession)


            'Call GetDFDTimeStemp()
            'RLU003
            'If gbDFD Then
            '    Call GetDFDTimeStemp()
            'End If

            Call ProcessWSIOSOrdersDataTable(giIressSession)

            If gbNetting Then
              Call ReportNetting()
              dtReportOrdersTable = dtNettingReportOrdersTable
            End If
            Call EndIOSPlusService(giIressSession)
          End If

          Call EndIOSPlusSession(giIressSession)
        End If
        EndIRESSSession()
      Else
        Exit For
      End If
    Next

    ' KC001
    If gsReportingType = "DELTA" Then
      Call WriteINI(gsIniFile, APP_NAME, "TradeDate", Format(gdDate, "MM/dd/yyyy"))
    End If
  End Sub

  ' KC001
  Private Sub ReadProcessedTrades(ByVal Session As Integer)
    Dim dTradeDate As Date

    dTradeDate = ReadIniDate(gsIniFile, APP_NAME, "TradeDate", gdDate.AddDays(-1))
    gaSourcesList(Session).TradesFile = ReadIniFullPath(gsIniFile, APP_NAME, "TradesFile" & CStr(Session), _
      gdDate, False, gsAppPath & "\" & APP_NAME & "-Trades-" & CStr(Session) & ".irs")

    If dTradeDate = gdDate Then
      Call ReadTradesFile(gaSourcesList(Session).TradesFile)
    Else
      Call DeleteFile(gaSourcesList(Session).TradesFile)
    End If
  End Sub

  ' KC001
  Private Sub ReadTradesFile(ByVal FileName As String)
    Dim iFileNum As Integer
    Dim sInputLine As String

    htProcessedTradesTable.Clear()

    If File.Exists(FileName) Then
      iFileNum = FreeFile()

      Try
        FileOpen(iFileNum, FileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        While Not EOF(iFileNum)
          sInputLine = Trim(LineInput(iFileNum))

          Try
            htProcessedTradesTable.Add(sInputLine, vbNull)
          Catch ex As Exception
            Call LogToFile("  Error: ReadTradesFile (" & FileName & ") - Unable to add to table (" & _
              sInputLine & ") - " & ex.Message)
          End Try
        End While

        FileClose(iFileNum)
      Catch ex As Exception
        Call LogToFile("  Error: ReadTradesFile - Unable to open file (" & FileName & ") - " & ex.Message)
      End Try
    End If
  End Sub

    Private Sub ReportNetting()
        Dim iIndex As Integer
        Dim dvreport As DataView
        Dim drReportRow As Data.DataRowView
        Dim drReportRow2 As Data.DataRowView
        Dim RtOrd As Long = 0
        Dim lastRootOrder As Long = 0
        Dim lastAccType As String = ""
        Dim AccType As String = ""

        Dim davgTrdPrc As Decimal = 0
        Dim lastTotalTrdVol As Long = 0
        Dim lastTotalTrdVal As Decimal = 0
        Dim TrdPrc As Decimal = 0
        Dim TrdVol As Long = 0
        Dim TrdVal As Decimal = 0

        Try
            dvreport = New DataView(dtReportOrdersTable)
            dvreport.Sort = "AccType,RootOrdNo"
            dvreport.RowFilter = "AccType = 'Market'"
            If dvreport.Count > 0 Then
                For iIndex = 0 To dvreport.Count - 1
                    drReportRow = dvreport.Item(iIndex)

                    Try
                        Call AddNettingReportOrder(drReportRow("AccCode"), drReportRow("Dest"), drReportRow("ActStat"),
                           drReportRow("LastAct"), drReportRow("BuySell"), drReportRow("SecCode"), drReportRow("OrdPrc"), drReportRow("TrdPrc"),
                           drReportRow("PrcInst"), drReportRow("Lifetime"), drReportRow("OrdVol"), drReportRow("TrdVol"), drReportRow("DoneVolTot"),
                           drReportRow("RemVol"), drReportRow("AvgPrc"), drReportRow("AccType"), drReportRow("ExecInstr"),
                           drReportRow("PostTradeStatus"), drReportRow("OrdNo"), drReportRow("RootOrdNo"), drReportRow("TradeNo"), drReportRow("PriCliOrd"), drReportRow("EXBR"), drReportRow("SettlementCurrency"),
                           drReportRow("FXrate"), drReportRow("Organization"), drReportRow("TradeTime"), drReportRow("SettlementTime"), drReportRow("SecurityType"), drReportRow("Currency"),
                           drReportRow("Exchange"), drReportRow("CUSIP"), drReportRow("ISIN"), drReportRow("SEDOL"), drReportRow("Description"),
                           drReportRow("DFANote1"), drReportRow("DFANote2"), drReportRow("DFANote3"), drReportRow("OpenClose"), drReportRow("UkIrishStampDutyReserveTaxMarker"), drReportRow("PtmLevyIndicator"), drReportRow("Source"), drReportRow("CurrencyDenomination"))

                    Catch ex As Exception
                        Call LogToFile("  Error: ReportNetting (Source" & CStr(0) & ") - AddNettingReportOrder - Market- " & ex.Message)
                    End Try

                    Continue For

                Next
            End If

            dvreport.RowFilter = "AccType = 'Client'"
            If dvreport.Count > 0 Then
                

                For iIndex = 0 To dvreport.Count - 1
                    drReportRow = dvreport.Item(iIndex)

                    RtOrd = drReportRow("RootOrdNo")
                    AccType = drReportRow("AccType")

                    'RLU011
                    If lastTotalTrdVol <> 0 And (lastRootOrder <> RtOrd Or AccType <> lastAccType) Then
                        'AddNettingReportOrder 

                        Try

                            drReportRow2 = dvreport.Item(iIndex - 1)




                            Call AddNettingReportOrder(drReportRow2("AccCode"), drReportRow2("Dest"), drReportRow2("ActStat"),
                                                   drReportRow2("LastAct"), drReportRow2("BuySell"), drReportRow2("SecCode"), drReportRow2("OrdPrc"), davgTrdPrc,
                                                   drReportRow2("PrcInst"), drReportRow2("Lifetime"), drReportRow2("OrdVol"), lastTotalTrdVol, drReportRow2("DoneVolTot"),
                                                   drReportRow2("RemVol"), drReportRow2("AvgPrc"), drReportRow2("AccType"), drReportRow2("ExecInstr"),
                                                   drReportRow2("PostTradeStatus"), drReportRow2("OrdNo"), drReportRow2("RootOrdNo"), drReportRow2("TradeNo"), drReportRow2("PriCliOrd"), drReportRow2("EXBR"), drReportRow2("SettlementCurrency"),
                                                   drReportRow2("FXrate"), drReportRow2("Organization"), drReportRow2("TradeTime"), drReportRow2("SettlementTime"), drReportRow2("SecurityType"), drReportRow2("Currency"),
                                                   drReportRow2("Exchange"), drReportRow2("CUSIP"), drReportRow2("ISIN"), drReportRow2("SEDOL"), drReportRow2("Description"),
                                                   drReportRow2("DFANote1"), drReportRow2("DFANote2"), drReportRow2("DFANote3"), drReportRow2("OpenClose"), drReportRow2("UkIrishStampDutyReserveTaxMarker"), drReportRow2("PtmLevyIndicator"), drReportRow2("Source"), drReportRow2("CurrencyDenomination"))


                        Catch ex As Exception
            Call LogToFile("  Error: ReportNetting (Source" & CStr(0) & ") - AddNettingReportOrder - " & ex.Message)
        End Try
                        lastTotalTrdVol = 0
                        lastTotalTrdVal = 0

                    End If

                    TrdVol = drReportRow("TrdVol")
                    TrdPrc = drReportRow("TrdPrc")
                    TrdVal = TrdVol * TrdPrc

                    lastTotalTrdVol = lastTotalTrdVol + TrdVol
                    lastTotalTrdVal = lastTotalTrdVal + TrdVal
                    If lastTotalTrdVol = 0 Then
                        lastTotalTrdVal = 0
                        lastTotalTrdVol = 1
                    End If

                    davgTrdPrc = lastTotalTrdVal / lastTotalTrdVol
                    lastRootOrder = RtOrd
                    lastAccType = AccType
                Next

                If lastTotalTrdVol > 0 Then
                    Call AddNettingReportOrder(drReportRow("AccCode"), drReportRow("Dest"), drReportRow("ActStat"),
                                                       drReportRow("LastAct"), drReportRow("BuySell"), drReportRow("SecCode"), drReportRow("OrdPrc"), davgTrdPrc,
                                                       drReportRow("PrcInst"), drReportRow("Lifetime"), drReportRow("OrdVol"), lastTotalTrdVol, drReportRow("DoneVolTot"),
                                                       drReportRow("RemVol"), drReportRow("AvgPrc"), drReportRow("AccType"), drReportRow("ExecInstr"),
                                                       drReportRow("PostTradeStatus"), drReportRow("OrdNo"), drReportRow("RootOrdNo"), drReportRow("TradeNo"), drReportRow("PriCliOrd"), drReportRow("EXBR"), drReportRow("SettlementCurrency"),
                                                       drReportRow("FXrate"), drReportRow("Organization"), drReportRow("TradeTime"), drReportRow("SettlementTime"), drReportRow("SecurityType"), drReportRow("Currency"),
                                                       drReportRow("Exchange"), drReportRow("CUSIP"), drReportRow("ISIN"), drReportRow("SEDOL"), drReportRow("Description"),
                                                       drReportRow("DFANote1"), drReportRow("DFANote2"), drReportRow("DFANote3"), drReportRow("OpenClose"), drReportRow("UkIrishStampDutyReserveTaxMarker"), drReportRow("PtmLevyIndicator"), drReportRow("Source"), drReportRow("CurrencyDenomination"))

                End If

            End If
 
        Catch ex As Exception
            Call LogToFile("  Error: ReportNetting (Source" & CStr(0) & ") - Unable to create DataView - " & ex.Message)
        End Try
    End Sub

    Private Sub GetDFDTimeStemp()
        Dim iIndex As Integer
        Dim dvWSIOSOrdersView As DataView
        Dim drWSIOSOrderRow As Data.DataRowView
        Dim sDestination As String = ""
        Try
            dvWSIOSOrdersView = New DataView(dtWSIOSOrdersTable)
            dvWSIOSOrdersView.Sort = "OrderNumber"

            'Call LogToFile("  info: GetWebServiceIOSAuditTrailByOrderNo - dvWSIOSOrdersView-" & dvWSIOSOrdersView.Count)

            For iIndex = 0 To dvWSIOSOrdersView.Count - 1
                drWSIOSOrderRow = dvWSIOSOrdersView.Item(iIndex)

                'RLU003
                sDestination = drWSIOSOrderRow("Destination")

                'If sDestination <> "DESK" Then
                '    Continue For
                'End If

                'RL004
                'If gsDefaultReportExceptionDestinations.Contains(sDestination) Then
                '    Continue For
                'End If

                'Call LogToFile(drWSIOSOrderRow("OrderNumber") & ":" & drWSIOSOrderRow("CreateDateTime") & ":" & dtWSTradesTable.Rows.Count)

                Call GetWebServiceIOSAuditTrailByOrderNo(drWSIOSOrderRow("OrderNumber"), drWSIOSOrderRow("RootOrderNumber"), gdStartTime, gdEndTime)
                'Call GetWebServiceIOSAuditTrailByOrderNo(drWSIOSOrderRow("OrderNumber"), drWSIOSOrderRow("RootOrderNumber"), Nothing, gdEndTime)
            Next
            'Call LogToFile("  info: GetWebServiceIOSAuditTrailByOrderNo - gDicDeskTradeNo " & gDicDeskTradeNo.Count)

        Catch ex As Exception
            Call LogToFile("  Error: GetDFDTimeStemp (Source" & CStr(0) & ") - Unable to create DataView - " & ex.Message)
            Exit Sub
        End Try
    End Sub



    Private Sub ReadWSIOSOrders(ByVal Session As Integer)
        'Dim iNum As Integer

        'For iNum = 0 To (MAX_REPORT_ACCOUNTS - 1)
        '        If gaReportAccountsList(iNum).AccCode <> "" Then
        '            Call GetWebServiceIOSOrdersByUser(gaReportAccountsList(iNum).AccCode, _
        '              gaReportAccountsList(iNum).Destinations, gaReportAccountsList(iNum).ExceptionDestinations)
        '        Else
        '            Exit For
        '        End If
        '    Next

        'Call GetWebServiceIOSOrdersByUser(gaReportAccountsList(Session).Destinations, gaReportAccountsList(Session).ExceptionDestinations)

        Call GetWebServiceIOSOrdersSearchByUser()
        'Call GetWebServiceIOSTradesByOrderNo(0)

        'Call LogToFile("  info: GetWebServiceIOSOrdersByUser- " & dtWSIOSOrdersTable.Rows.Count)
    End Sub

    Private Sub CreateReport()

        If Not IO.File.Exists(gsTempOutputName) Then
            Dim file As FileStream = IO.File.Create(gsTempOutputName)
            file.Close()
        End If






        If OpenOutput(0) Then
            Call ProcessReportDataTable(0)
            Call CloseOutput(0)
        End If

        If gbTempToFinal And File.Exists(gsTempOutputName) Then
            If File.Exists(gsOutputName) Then
                gsOutputName = gsOutputName.Replace(Format(gdDate, "yyyyMMdd"), Format(gdDate, "yyyyMMdd") & "-" & Format(TimeOfDay, "hhmmss"))
            End If
            File.Copy(gsTempOutputName, gsOutputName)
        End If
        If gbRemoveTemp And gbTempToFinal Then
            File.Delete(gsTempOutputName)
        End If
    End Sub

    Private Sub ApplicationClose()
        Call LogToFile("End of process")
        Me.Close()
    End Sub

End Class
