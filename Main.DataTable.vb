Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain
#Region "Create"
    Private Sub CreateDataTables()
        Call CreateWSIOSOrdersDataTable()
        Call CreateWSTradesDataTable()
        Call CreateReportOrdersDataTable()
        Call CreateNettingReportOrdersDataTable()
    End Sub

    Private Sub CreateWSIOSOrdersDataTable()
        dtWSIOSOrdersTable = New Data.DataTable("WSIOSOrders")

        Dim colOrderNumber As New Data.DataColumn               ' 0
        colOrderNumber.DataType = System.Type.GetType("System.Int64")
        colOrderNumber.ColumnName = "OrderNumber"
        dtWSIOSOrdersTable.Columns.Add(colOrderNumber)
        Dim colRootOrderNumber As New Data.DataColumn               ' 0
        colRootOrderNumber.DataType = System.Type.GetType("System.Int64")
        colRootOrderNumber.ColumnName = "RootOrderNumber"
        dtWSIOSOrdersTable.Columns.Add(colRootOrderNumber)
        Dim colAccountCode As New Data.DataColumn               ' 1
        colAccountCode.DataType = System.Type.GetType("System.String")
        colAccountCode.ColumnName = "AccountCode"
        dtWSIOSOrdersTable.Columns.Add(colAccountCode)
        Dim colSecurityCode As New Data.DataColumn              ' 2
        colSecurityCode.DataType = System.Type.GetType("System.String")
        colSecurityCode.ColumnName = "SecurityCode"
        dtWSIOSOrdersTable.Columns.Add(colSecurityCode)
        Dim colExchange As New Data.DataColumn                  ' 3
        colExchange.DataType = System.Type.GetType("System.String")
        colExchange.ColumnName = "Exchange"
        dtWSIOSOrdersTable.Columns.Add(colExchange)
        Dim colDestination As New Data.DataColumn               ' 4
        colDestination.DataType = System.Type.GetType("System.String")
        colDestination.ColumnName = "Destination"
        dtWSIOSOrdersTable.Columns.Add(colDestination)
        Dim colBuyOrSell As New Data.DataColumn                 ' 5
        colBuyOrSell.DataType = System.Type.GetType("System.String")
        colBuyOrSell.ColumnName = "BuyOrSell"
        dtWSIOSOrdersTable.Columns.Add(colBuyOrSell)
        Dim colPricingInstructions As New Data.DataColumn       ' 6
        colPricingInstructions.DataType = System.Type.GetType("System.String")
        colPricingInstructions.ColumnName = "PricingInstructions"
        dtWSIOSOrdersTable.Columns.Add(colPricingInstructions)
        Dim colLastAction As New Data.DataColumn                ' 7
        colLastAction.DataType = System.Type.GetType("System.String")
        colLastAction.ColumnName = "LastAction"
        dtWSIOSOrdersTable.Columns.Add(colLastAction)
        Dim colActionStatus As New Data.DataColumn              ' 8
        colActionStatus.DataType = System.Type.GetType("System.String")
        colActionStatus.ColumnName = "ActionStatus"
        dtWSIOSOrdersTable.Columns.Add(colActionStatus)
        Dim colOrderVolume As New Data.DataColumn               ' 9
        colOrderVolume.DataType = System.Type.GetType("System.Int32")
        colOrderVolume.ColumnName = "OrderVolume"
        dtWSIOSOrdersTable.Columns.Add(colOrderVolume)
        Dim colOrderPrice As New Data.DataColumn                ' 10
        colOrderPrice.DataType = System.Type.GetType("System.Decimal")
        colOrderPrice.ColumnName = "OrderPrice"
        dtWSIOSOrdersTable.Columns.Add(colOrderPrice)
        Dim colRemainingVolume As New Data.DataColumn           ' 11
        colRemainingVolume.DataType = System.Type.GetType("System.Int32")
        colRemainingVolume.ColumnName = "RemainingVolume"
        dtWSIOSOrdersTable.Columns.Add(colRemainingVolume)
        Dim colDoneVolumeTotal As New Data.DataColumn           ' 12
        colDoneVolumeTotal.DataType = System.Type.GetType("System.Int32")
        colDoneVolumeTotal.ColumnName = "DoneVolumeTotal"
        dtWSIOSOrdersTable.Columns.Add(colDoneVolumeTotal)
        Dim colAveragePrice As New Data.DataColumn              ' 13
        colAveragePrice.DataType = System.Type.GetType("System.Decimal")
        colAveragePrice.ColumnName = "AveragePrice"
        dtWSIOSOrdersTable.Columns.Add(colAveragePrice)
        Dim colLifetime As New Data.DataColumn                  ' 14
        colLifetime.DataType = System.Type.GetType("System.String")
        colLifetime.ColumnName = "Lifetime"
        dtWSIOSOrdersTable.Columns.Add(colLifetime)
        Dim colExecutionInstructions As New Data.DataColumn     ' 15
        colExecutionInstructions.DataType = System.Type.GetType("System.String")
        colExecutionInstructions.ColumnName = "ExecutionInstructions"
        dtWSIOSOrdersTable.Columns.Add(colExecutionInstructions)
        Dim colCurrency As New Data.DataColumn                  ' 16
        colCurrency.DataType = System.Type.GetType("System.String")
        colCurrency.ColumnName = "Currency"
        dtWSIOSOrdersTable.Columns.Add(colCurrency)
        Dim colPrimaryClientOrderId As New Data.DataColumn      ' 17
        colPrimaryClientOrderId.DataType = System.Type.GetType("System.String")
        colPrimaryClientOrderId.ColumnName = "PrimaryClientOrderId"
        dtWSIOSOrdersTable.Columns.Add(colPrimaryClientOrderId)

        Dim colPostTradeStatusNumber As New Data.DataColumn     ' 18
        colPostTradeStatusNumber.DataType = System.Type.GetType("System.Int32")
        colPostTradeStatusNumber.ColumnName = "PostTradeStatusNumber"
        dtWSIOSOrdersTable.Columns.Add(colPostTradeStatusNumber)

        Dim cDFXrate As New Data.DataColumn     ' 19
        cDFXrate.DataType = System.Type.GetType("System.Decimal")
        cDFXrate.ColumnName = "FXrate"
        dtWSIOSOrdersTable.Columns.Add(cDFXrate)

        Dim sOrganization As New Data.DataColumn     ' 20
        sOrganization.DataType = System.Type.GetType("System.String")
        sOrganization.ColumnName = "Organization"
        dtWSIOSOrdersTable.Columns.Add(sOrganization)

        Dim CreateTime As New Data.DataColumn     ' 21
        CreateTime.DataType = System.Type.GetType("System.DateTime")
        CreateTime.ColumnName = "CreateDateTime"
        dtWSIOSOrdersTable.Columns.Add(CreateTime)

        Dim TradeTime As New Data.DataColumn     ' 21
        TradeTime.DataType = System.Type.GetType("System.DateTime")
        TradeTime.ColumnName = "UpdateDateTime"
        dtWSIOSOrdersTable.Columns.Add(TradeTime)

        Dim iSecurityType As New Data.DataColumn     ' 22
        iSecurityType.DataType = System.Type.GetType("System.Int32")
        iSecurityType.ColumnName = "SecurityType"
        dtWSIOSOrdersTable.Columns.Add(iSecurityType)

        Dim iOrderFlags As New Data.DataColumn     ' 22
        iOrderFlags.DataType = System.Type.GetType("System.Int64")
        iOrderFlags.ColumnName = "OrderFlags"
        dtWSIOSOrdersTable.Columns.Add(iOrderFlags)

    End Sub

    Private Sub CreateReportOrdersDataTable()
        dtReportOrdersTable = New Data.DataTable("ReportOrders")

        Dim colAccCode As New Data.DataColumn                   ' 0
        colAccCode.DataType = System.Type.GetType("System.String")
        colAccCode.ColumnName = "AccCode"
        dtReportOrdersTable.Columns.Add(colAccCode)
        Dim colDest As New Data.DataColumn                      ' 1
        colDest.DataType = System.Type.GetType("System.String")
        colDest.ColumnName = "Dest"
        dtReportOrdersTable.Columns.Add(colDest)
        Dim colActStat As New Data.DataColumn                   ' 2
        colActStat.DataType = System.Type.GetType("System.String")
        colActStat.ColumnName = "ActStat"
        dtReportOrdersTable.Columns.Add(colActStat)
        Dim colLastAct As New Data.DataColumn                   ' 3
        colLastAct.DataType = System.Type.GetType("System.String")
        colLastAct.ColumnName = "LastAct"
        dtReportOrdersTable.Columns.Add(colLastAct)
        Dim colBuySell As New Data.DataColumn                   ' 4
        colBuySell.DataType = System.Type.GetType("System.String")
        colBuySell.ColumnName = "BuySell"
        dtReportOrdersTable.Columns.Add(colBuySell)
        Dim colSecCode As New Data.DataColumn                   ' 5
        colSecCode.DataType = System.Type.GetType("System.String")
        colSecCode.ColumnName = "SecCode"
        dtReportOrdersTable.Columns.Add(colSecCode)

        Dim colOrdPrc As New Data.DataColumn                    ' 6
        colOrdPrc.DataType = System.Type.GetType("System.Decimal")
        colOrdPrc.ColumnName = "OrdPrc"
        dtReportOrdersTable.Columns.Add(colOrdPrc)

        Dim colTrdPrc As New Data.DataColumn                    ' 6
        colTrdPrc.DataType = System.Type.GetType("System.Decimal")
        colTrdPrc.ColumnName = "TrdPrc"
        dtReportOrdersTable.Columns.Add(colTrdPrc)

        Dim colPrcInst As New Data.DataColumn                   ' 7
        colPrcInst.DataType = System.Type.GetType("System.String")
        colPrcInst.ColumnName = "PrcInst"
        dtReportOrdersTable.Columns.Add(colPrcInst)
        Dim colLifetime As New Data.DataColumn                  ' 8
        colLifetime.DataType = System.Type.GetType("System.String")
        colLifetime.ColumnName = "Lifetime"
        dtReportOrdersTable.Columns.Add(colLifetime)

        Dim colOrdVol As New Data.DataColumn                    ' 9
        colOrdVol.DataType = System.Type.GetType("System.Int32")
        colOrdVol.ColumnName = "OrdVol"
        dtReportOrdersTable.Columns.Add(colOrdVol)

        Dim colTrdVol As New Data.DataColumn                    ' 9
        colTrdVol.DataType = System.Type.GetType("System.Int32")
        colTrdVol.ColumnName = "TrdVol"
        dtReportOrdersTable.Columns.Add(colTrdVol)

        Dim colDoneVolTot As New Data.DataColumn                ' 10
        colDoneVolTot.DataType = System.Type.GetType("System.Int32")
        colDoneVolTot.ColumnName = "DoneVolTot"
        dtReportOrdersTable.Columns.Add(colDoneVolTot)
        Dim colRemVol As New Data.DataColumn                    ' 11
        colRemVol.DataType = System.Type.GetType("System.Int32")
        colRemVol.ColumnName = "RemVol"
        dtReportOrdersTable.Columns.Add(colRemVol)
        Dim colAvgPrc As New Data.DataColumn                    ' 12
        colAvgPrc.DataType = System.Type.GetType("System.Decimal")
        colAvgPrc.ColumnName = "AvgPrc"
        dtReportOrdersTable.Columns.Add(colAvgPrc)
        Dim colAccType As New Data.DataColumn                   ' 13
        colAccType.DataType = System.Type.GetType("System.String")
        colAccType.ColumnName = "AccType"
        dtReportOrdersTable.Columns.Add(colAccType)
        Dim colExecInstr As New Data.DataColumn                 ' 14
        colExecInstr.DataType = System.Type.GetType("System.String")
        colExecInstr.ColumnName = "ExecInstr"
        dtReportOrdersTable.Columns.Add(colExecInstr)
        Dim colPostTradeStatus As New Data.DataColumn           ' 15
        colPostTradeStatus.DataType = System.Type.GetType("System.String")
        colPostTradeStatus.ColumnName = "PostTradeStatus"
        dtReportOrdersTable.Columns.Add(colPostTradeStatus)

        Dim colOrdNo As New Data.DataColumn                     ' 16
        colOrdNo.DataType = System.Type.GetType("System.Int64")
        colOrdNo.ColumnName = "OrdNo"
        dtReportOrdersTable.Columns.Add(colOrdNo)

        Dim colRtOrdNo As New Data.DataColumn                     ' 16
        colRtOrdNo.DataType = System.Type.GetType("System.Int64")
        colRtOrdNo.ColumnName = "RootOrdNo"
        dtReportOrdersTable.Columns.Add(colRtOrdNo)

        Dim colTradeNo As New Data.DataColumn                     ' 16
        colTradeNo.DataType = System.Type.GetType("System.Int64")
        colTradeNo.ColumnName = "TradeNo"
        dtReportOrdersTable.Columns.Add(colTradeNo)

        Dim colPriCliOrd As New Data.DataColumn                 ' 17
        colPriCliOrd.DataType = System.Type.GetType("System.String")
        colPriCliOrd.ColumnName = "PriCliOrd"
        dtReportOrdersTable.Columns.Add(colPriCliOrd)

        Dim sEXBR As New Data.DataColumn     ' 20
        sEXBR.DataType = System.Type.GetType("System.String")
        sEXBR.ColumnName = "EXBR"
        dtReportOrdersTable.Columns.Add(sEXBR)

        Dim sSettlementCurrency As New Data.DataColumn     ' 20
        sSettlementCurrency.DataType = System.Type.GetType("System.String")
        sSettlementCurrency.ColumnName = "SettlementCurrency"
        dtReportOrdersTable.Columns.Add(sSettlementCurrency)

        Dim cDFXrate As New Data.DataColumn     ' 19
        cDFXrate.DataType = System.Type.GetType("System.Decimal")
        cDFXrate.ColumnName = "FXrate"
        dtReportOrdersTable.Columns.Add(cDFXrate)

        Dim sOrganization As New Data.DataColumn     ' 20
        sOrganization.DataType = System.Type.GetType("System.String")
        sOrganization.ColumnName = "Organization"
        dtReportOrdersTable.Columns.Add(sOrganization)

        Dim TradeTime As New Data.DataColumn     ' 21
        TradeTime.DataType = System.Type.GetType("System.String")
        TradeTime.ColumnName = "TradeTime"
        dtReportOrdersTable.Columns.Add(TradeTime)
        Dim SettlementTime As New Data.DataColumn     ' 21
        SettlementTime.DataType = System.Type.GetType("System.String")
        SettlementTime.ColumnName = "SettlementTime"
        dtReportOrdersTable.Columns.Add(SettlementTime)
        Dim iSecurityType As New Data.DataColumn     ' 22
        iSecurityType.DataType = System.Type.GetType("System.String")
        iSecurityType.ColumnName = "SecurityType"
        dtReportOrdersTable.Columns.Add(iSecurityType)
        'Currency
        Dim Currency As New Data.DataColumn     ' 22
        Currency.DataType = System.Type.GetType("System.String")
        Currency.ColumnName = "Currency"
        dtReportOrdersTable.Columns.Add(Currency)
        'Exchange
        Dim Exchange As New Data.DataColumn     ' 22
        Exchange.DataType = System.Type.GetType("System.String")
        Exchange.ColumnName = "Exchange"
        dtReportOrdersTable.Columns.Add(Exchange)


        ' CUSIP, ISIN, SEDOL, Description
        Dim cISIN As New Data.DataColumn     ' 22
        cISIN.ColumnName = "ISIN"
        dtReportOrdersTable.Columns.Add(cISIN)
        Dim cCUSIP As New Data.DataColumn                  ' 3
        cCUSIP.DataType = System.Type.GetType("System.String")
        cCUSIP.ColumnName = "CUSIP"
        dtReportOrdersTable.Columns.Add(cCUSIP)
        Dim cSEDOL As New Data.DataColumn                 ' 4
        cSEDOL.DataType = System.Type.GetType("System.String")
        cSEDOL.ColumnName = "SEDOL"
        dtReportOrdersTable.Columns.Add(cSEDOL)
        Dim cDescription As New Data.DataColumn                 ' 5
        cDescription.DataType = System.Type.GetType("System.String")
        cDescription.ColumnName = "Description"
        dtReportOrdersTable.Columns.Add(cDescription)

        Dim cDFANote1 As New Data.DataColumn                  ' 3
        cDFANote1.DataType = System.Type.GetType("System.String")
        cDFANote1.ColumnName = "DFANote1"
        dtReportOrdersTable.Columns.Add(cDFANote1)

        Dim DFANote2 As New Data.DataColumn                 ' 4
        DFANote2.DataType = System.Type.GetType("System.String")
        DFANote2.ColumnName = "DFANote2"
        dtReportOrdersTable.Columns.Add(DFANote2)

        Dim cDFANote3 As New Data.DataColumn                 ' 5
        cDFANote3.DataType = System.Type.GetType("System.String")
        cDFANote3.ColumnName = "DFANote3"
        dtReportOrdersTable.Columns.Add(cDFANote3)

        Dim cOpenClose As New Data.DataColumn                 ' 5
        cOpenClose.DataType = System.Type.GetType("System.String")
        cOpenClose.ColumnName = "OpenClose"
        dtReportOrdersTable.Columns.Add(cOpenClose)

        Dim cUkIrishStampDutyReserveTaxMarker As New Data.DataColumn                 ' 5
        cUkIrishStampDutyReserveTaxMarker.DataType = System.Type.GetType("System.String")
        cUkIrishStampDutyReserveTaxMarker.ColumnName = "UkIrishStampDutyReserveTaxMarker"
        dtReportOrdersTable.Columns.Add(cUkIrishStampDutyReserveTaxMarker)

        Dim cPtmLevyIndicator As New Data.DataColumn                 ' 5
        cPtmLevyIndicator.DataType = System.Type.GetType("System.String")
        cPtmLevyIndicator.ColumnName = "PtmLevyIndicator"
        dtReportOrdersTable.Columns.Add(cPtmLevyIndicator)

        Dim cSource As New Data.DataColumn                 ' 5
        cSource.DataType = System.Type.GetType("System.String")
        cSource.ColumnName = "Source"
        dtReportOrdersTable.Columns.Add(cSource)

        'CurrencyDenomination
        Dim cCurrencyDenomination As New Data.DataColumn                 ' 5
        cCurrencyDenomination.DataType = System.Type.GetType("System.String")
        cCurrencyDenomination.ColumnName = "CurrencyDenomination"
        dtReportOrdersTable.Columns.Add(cCurrencyDenomination)
    End Sub

    Private Sub CreateNettingReportOrdersDataTable()
        dtNettingReportOrdersTable = New Data.DataTable("NettingReportOrders")

        Dim colAccCode As New Data.DataColumn                   ' 0
        colAccCode.DataType = System.Type.GetType("System.String")
        colAccCode.ColumnName = "AccCode"
        dtNettingReportOrdersTable.Columns.Add(colAccCode)
        Dim colDest As New Data.DataColumn                      ' 1
        colDest.DataType = System.Type.GetType("System.String")
        colDest.ColumnName = "Dest"
        dtNettingReportOrdersTable.Columns.Add(colDest)
        Dim colActStat As New Data.DataColumn                   ' 2
        colActStat.DataType = System.Type.GetType("System.String")
        colActStat.ColumnName = "ActStat"
        dtNettingReportOrdersTable.Columns.Add(colActStat)
        Dim colLastAct As New Data.DataColumn                   ' 3
        colLastAct.DataType = System.Type.GetType("System.String")
        colLastAct.ColumnName = "LastAct"
        dtNettingReportOrdersTable.Columns.Add(colLastAct)
        Dim colBuySell As New Data.DataColumn                   ' 4
        colBuySell.DataType = System.Type.GetType("System.String")
        colBuySell.ColumnName = "BuySell"
        dtNettingReportOrdersTable.Columns.Add(colBuySell)
        Dim colSecCode As New Data.DataColumn                   ' 5
        colSecCode.DataType = System.Type.GetType("System.String")
        colSecCode.ColumnName = "SecCode"
        dtNettingReportOrdersTable.Columns.Add(colSecCode)

        Dim colOrdPrc As New Data.DataColumn                    ' 6
        colOrdPrc.DataType = System.Type.GetType("System.Decimal")
        colOrdPrc.ColumnName = "OrdPrc"
        dtNettingReportOrdersTable.Columns.Add(colOrdPrc)

        Dim colTrdPrc As New Data.DataColumn                    ' 6
        colTrdPrc.DataType = System.Type.GetType("System.Decimal")
        colTrdPrc.ColumnName = "TrdPrc"
        dtNettingReportOrdersTable.Columns.Add(colTrdPrc)

        Dim colPrcInst As New Data.DataColumn                   ' 7
        colPrcInst.DataType = System.Type.GetType("System.String")
        colPrcInst.ColumnName = "PrcInst"
        dtNettingReportOrdersTable.Columns.Add(colPrcInst)
        Dim colLifetime As New Data.DataColumn                  ' 8
        colLifetime.DataType = System.Type.GetType("System.String")
        colLifetime.ColumnName = "Lifetime"
        dtNettingReportOrdersTable.Columns.Add(colLifetime)

        Dim colOrdVol As New Data.DataColumn                    ' 9
        colOrdVol.DataType = System.Type.GetType("System.Int32")
        colOrdVol.ColumnName = "OrdVol"
        dtNettingReportOrdersTable.Columns.Add(colOrdVol)

        Dim colTrdVol As New Data.DataColumn                    ' 9
        colTrdVol.DataType = System.Type.GetType("System.Int32")
        colTrdVol.ColumnName = "TrdVol"
        dtNettingReportOrdersTable.Columns.Add(colTrdVol)

        Dim colDoneVolTot As New Data.DataColumn                ' 10
        colDoneVolTot.DataType = System.Type.GetType("System.Int32")
        colDoneVolTot.ColumnName = "DoneVolTot"
        dtNettingReportOrdersTable.Columns.Add(colDoneVolTot)
        Dim colRemVol As New Data.DataColumn                    ' 11
        colRemVol.DataType = System.Type.GetType("System.Int32")
        colRemVol.ColumnName = "RemVol"
        dtNettingReportOrdersTable.Columns.Add(colRemVol)
        Dim colAvgPrc As New Data.DataColumn                    ' 12
        colAvgPrc.DataType = System.Type.GetType("System.Decimal")
        colAvgPrc.ColumnName = "AvgPrc"
        dtNettingReportOrdersTable.Columns.Add(colAvgPrc)
        Dim colAccType As New Data.DataColumn                   ' 13
        colAccType.DataType = System.Type.GetType("System.String")
        colAccType.ColumnName = "AccType"
        dtNettingReportOrdersTable.Columns.Add(colAccType)
        Dim colExecInstr As New Data.DataColumn                 ' 14
        colExecInstr.DataType = System.Type.GetType("System.String")
        colExecInstr.ColumnName = "ExecInstr"
        dtNettingReportOrdersTable.Columns.Add(colExecInstr)
        Dim colPostTradeStatus As New Data.DataColumn           ' 15
        colPostTradeStatus.DataType = System.Type.GetType("System.String")
        colPostTradeStatus.ColumnName = "PostTradeStatus"
        dtNettingReportOrdersTable.Columns.Add(colPostTradeStatus)

        Dim colOrdNo As New Data.DataColumn                     ' 16
        colOrdNo.DataType = System.Type.GetType("System.Int64")
        colOrdNo.ColumnName = "OrdNo"
        dtNettingReportOrdersTable.Columns.Add(colOrdNo)

        Dim colRtOrdNo As New Data.DataColumn                     ' 16
        colRtOrdNo.DataType = System.Type.GetType("System.Int64")
        colRtOrdNo.ColumnName = "RootOrdNo"
        dtNettingReportOrdersTable.Columns.Add(colRtOrdNo)

        Dim colTradeNo As New Data.DataColumn                     ' 16
        colTradeNo.DataType = System.Type.GetType("System.Int64")
        colTradeNo.ColumnName = "TradeNo"
        dtNettingReportOrdersTable.Columns.Add(colTradeNo)

        Dim colPriCliOrd As New Data.DataColumn                 ' 17
        colPriCliOrd.DataType = System.Type.GetType("System.String")
        colPriCliOrd.ColumnName = "PriCliOrd"
        dtNettingReportOrdersTable.Columns.Add(colPriCliOrd)

        Dim sEXBR As New Data.DataColumn     ' 20
        sEXBR.DataType = System.Type.GetType("System.String")
        sEXBR.ColumnName = "EXBR"
        dtNettingReportOrdersTable.Columns.Add(sEXBR)

        Dim sSettlementCurrency As New Data.DataColumn     ' 20
        sSettlementCurrency.DataType = System.Type.GetType("System.String")
        sSettlementCurrency.ColumnName = "SettlementCurrency"
        dtNettingReportOrdersTable.Columns.Add(sSettlementCurrency)

        Dim cDFXrate As New Data.DataColumn     ' 19
        cDFXrate.DataType = System.Type.GetType("System.Decimal")
        cDFXrate.ColumnName = "FXrate"
        dtNettingReportOrdersTable.Columns.Add(cDFXrate)

        Dim sOrganization As New Data.DataColumn     ' 20
        sOrganization.DataType = System.Type.GetType("System.String")
        sOrganization.ColumnName = "Organization"
        dtNettingReportOrdersTable.Columns.Add(sOrganization)

        Dim TradeTime As New Data.DataColumn     ' 21
        TradeTime.DataType = System.Type.GetType("System.String")
        TradeTime.ColumnName = "TradeTime"
        dtNettingReportOrdersTable.Columns.Add(TradeTime)
        Dim SettlementTime As New Data.DataColumn     ' 21
        SettlementTime.DataType = System.Type.GetType("System.String")
        SettlementTime.ColumnName = "SettlementTime"
        dtNettingReportOrdersTable.Columns.Add(SettlementTime)
        Dim iSecurityType As New Data.DataColumn     ' 22
        iSecurityType.DataType = System.Type.GetType("System.String")
        iSecurityType.ColumnName = "SecurityType"
        dtNettingReportOrdersTable.Columns.Add(iSecurityType)
        'Currency
        Dim Currency As New Data.DataColumn     ' 22
        Currency.DataType = System.Type.GetType("System.String")
        Currency.ColumnName = "Currency"
        dtNettingReportOrdersTable.Columns.Add(Currency)
        'Exchange
        Dim Exchange As New Data.DataColumn     ' 22
        Exchange.DataType = System.Type.GetType("System.String")
        Exchange.ColumnName = "Exchange"
        dtNettingReportOrdersTable.Columns.Add(Exchange)


        ' CUSIP, ISIN, SEDOL, Description
        Dim cISIN As New Data.DataColumn     ' 22
        cISIN.ColumnName = "ISIN"
        dtNettingReportOrdersTable.Columns.Add(cISIN)
        Dim cCUSIP As New Data.DataColumn                  ' 3
        cCUSIP.DataType = System.Type.GetType("System.String")
        cCUSIP.ColumnName = "CUSIP"
        dtNettingReportOrdersTable.Columns.Add(cCUSIP)

        Dim cSEDOL As New Data.DataColumn                 ' 4
        cSEDOL.DataType = System.Type.GetType("System.String")
        cSEDOL.ColumnName = "SEDOL"
        dtNettingReportOrdersTable.Columns.Add(cSEDOL)

        Dim cDescription As New Data.DataColumn                 ' 5
        cDescription.DataType = System.Type.GetType("System.String")
        cDescription.ColumnName = "Description"
        dtNettingReportOrdersTable.Columns.Add(cDescription)

        Dim cDFANote1 As New Data.DataColumn                  ' 3
        cDFANote1.DataType = System.Type.GetType("System.String")
        cDFANote1.ColumnName = "DFANote1"
        dtNettingReportOrdersTable.Columns.Add(cDFANote1)

        Dim DFANote2 As New Data.DataColumn                 ' 4
        DFANote2.DataType = System.Type.GetType("System.String")
        DFANote2.ColumnName = "DFANote2"
        dtNettingReportOrdersTable.Columns.Add(DFANote2)

        Dim cDFANote3 As New Data.DataColumn                 ' 5
        cDFANote3.DataType = System.Type.GetType("System.String")
        cDFANote3.ColumnName = "DFANote3"
        dtNettingReportOrdersTable.Columns.Add(cDFANote3)

        Dim cOpenClose As New Data.DataColumn                 ' 5
        cOpenClose.DataType = System.Type.GetType("System.String")
        cOpenClose.ColumnName = "OpenClose"
        dtNettingReportOrdersTable.Columns.Add(cOpenClose)

        Dim cUkIrishStampDutyReserveTaxMarker As New Data.DataColumn                 ' 5
        cUkIrishStampDutyReserveTaxMarker.DataType = System.Type.GetType("System.String")
        cUkIrishStampDutyReserveTaxMarker.ColumnName = "UkIrishStampDutyReserveTaxMarker"
        dtNettingReportOrdersTable.Columns.Add(cUkIrishStampDutyReserveTaxMarker)

        Dim cPtmLevyIndicator As New Data.DataColumn                 ' 5
        cPtmLevyIndicator.DataType = System.Type.GetType("System.String")
        cPtmLevyIndicator.ColumnName = "PtmLevyIndicator"
        dtNettingReportOrdersTable.Columns.Add(cPtmLevyIndicator)

        Dim cSource As New Data.DataColumn                 ' 5
        cSource.DataType = System.Type.GetType("System.String")
        cSource.ColumnName = "Source"
        dtNettingReportOrdersTable.Columns.Add(cSource)

        Dim cCurrencyDenomination As New Data.DataColumn                 ' 5
        cCurrencyDenomination.DataType = System.Type.GetType("System.String")
        cCurrencyDenomination.ColumnName = "CurrencyDenomination"
        dtNettingReportOrdersTable.Columns.Add(cCurrencyDenomination)
    End Sub

    Private Sub CreateWSTradesDataTable()
        dtWSTradesTable = New Data.DataTable("WSTrades")

        Dim colTrdNo As New Data.DataColumn                     ' 0
        colTrdNo.DataType = System.Type.GetType("System.Int64")
        colTrdNo.ColumnName = "TradeNumber"
        dtWSTradesTable.Columns.Add(colTrdNo)

        Dim colOrdNo As New Data.DataColumn                     ' 1
        colOrdNo.DataType = System.Type.GetType("System.Int64")
        colOrdNo.ColumnName = "OrderNumber"
        dtWSTradesTable.Columns.Add(colOrdNo)

        Dim colPrtOrdNo As New Data.DataColumn                     ' 1
        colPrtOrdNo.DataType = System.Type.GetType("System.Int64")
        colPrtOrdNo.ColumnName = "RootOrderNumber"
        dtWSTradesTable.Columns.Add(colPrtOrdNo)

        Dim colExchange As New Data.DataColumn                  ' 1
        colExchange.DataType = System.Type.GetType("System.String")
        colExchange.ColumnName = "Exchange"
        dtWSTradesTable.Columns.Add(colExchange)

        Dim cDestination As New Data.DataColumn                  ' 1
        cDestination.DataType = System.Type.GetType("System.String")
        cDestination.ColumnName = "Destination"
        dtWSTradesTable.Columns.Add(cDestination)


        Dim cTradeVolume As New Data.DataColumn                     ' 1
        cTradeVolume.DataType = System.Type.GetType("System.Int64")
        cTradeVolume.ColumnName = "TradeVolume"
        dtWSTradesTable.Columns.Add(cTradeVolume)

        Dim cTradePrice As New Data.DataColumn                     ' 1
        cTradePrice.DataType = System.Type.GetType("System.Decimal")
        cTradePrice.ColumnName = "TradePrice"
        dtWSTradesTable.Columns.Add(cTradePrice)

        Dim cFXrate As New Data.DataColumn                     ' 1
        cFXrate.DataType = System.Type.GetType("System.Decimal")
        cFXrate.ColumnName = "FXrate"
        dtWSTradesTable.Columns.Add(cFXrate)

        Dim cTradeDateTime As New Data.DataColumn                     ' 1
        cTradeDateTime.DataType = System.Type.GetType("System.DateTime")
        cTradeDateTime.ColumnName = "TradeDateTime"
        dtWSTradesTable.Columns.Add(cTradeDateTime)

        Dim cOpposingBrokerNumber As New Data.DataColumn                     ' 1
        cOpposingBrokerNumber.DataType = System.Type.GetType("System.Int64")
        cOpposingBrokerNumber.ColumnName = "OpposingBrokerNumber"
        dtWSTradesTable.Columns.Add(cOpposingBrokerNumber)

        Dim cTradeMarkers As New Data.DataColumn                     ' 1
        cTradeMarkers.DataType = System.Type.GetType("System.String")
        cTradeMarkers.ColumnName = "TradeMarkers"
        dtWSTradesTable.Columns.Add(cTradeMarkers)

        Dim cSourcePrice As New Data.DataColumn                     ' 1
        cSourcePrice.DataType = System.Type.GetType("System.Decimal")
        cSourcePrice.ColumnName = "SourcePrice"
        dtWSTradesTable.Columns.Add(cSourcePrice)

        Dim cAccountCode As New Data.DataColumn                     ' 1
        cAccountCode.DataType = System.Type.GetType("System.String")
        cAccountCode.ColumnName = "AccountCode"
        dtWSTradesTable.Columns.Add(cAccountCode)

        Dim colSecurityCode As New Data.DataColumn              ' 0
        colSecurityCode.DataType = System.Type.GetType("System.String")
        colSecurityCode.ColumnName = "SecurityCode"
        dtWSTradesTable.Columns.Add(colSecurityCode)
    End Sub

#End Region

    Private Sub AddWSIOSOrder(ByVal OrderNumber As Long, ByVal RootOrderNumber As Long, ByVal AccountCode As String, ByVal SecurityCode As String,
      ByVal Exchange As String, ByVal Destination As String, ByVal BuyOrSell As String,
      ByVal PricingInstructions As String, ByVal LastAction As String, ByVal ActionStatus As String,
      ByVal OrderVolume As Long, ByVal OrderPrice As Decimal, ByVal RemainingVolume As Long, ByVal DoneVolumeTotal As Long,
      ByVal AveragePrice As Decimal, ByVal Lifetime As String, ByVal ExecutionInstructions As String,
      ByVal Currency As String, ByVal PrimaryClientOrderId As String, ByVal PostTradeStatusNumber As Integer, ByVal DFXrate As Decimal, ByVal sOrganization As String, ByVal CreateDateTime As DateTime, ByVal TradeTime As DateTime, ByVal iSecurityType As Integer, ByVal iOrderFlags As Long)
        Dim drWSIOSOrder As DataRow
        Try
            'RLU
            'If Exchange = "LSE" Or Exchange = "ASX" Or Exchange = "JSE" Then
            '    OrderPrice = OrderPrice / 100
            '    AveragePrice = AveragePrice / 100
            'End If
            'RL009
            Dim Multipier As Decimal = 1
            If htExchangeMultipliers.ContainsKey(Exchange) Then
                Multipier = htExchangeMultipliers.Item(Exchange)
            End If
            OrderPrice = OrderPrice * Multipier
            AveragePrice = AveragePrice * Multipier

            drWSIOSOrder = dtWSIOSOrdersTable.NewRow()

            drWSIOSOrder("OrderNumber") = OrderNumber
            drWSIOSOrder("RootOrderNumber") = RootOrderNumber

            drWSIOSOrder("AccountCode") = AccountCode
            drWSIOSOrder("SecurityCode") = SecurityCode
            drWSIOSOrder("Exchange") = Exchange
            drWSIOSOrder("Destination") = Destination
            drWSIOSOrder("BuyOrSell") = BuyOrSell
            drWSIOSOrder("PricingInstructions") = PricingInstructions
            drWSIOSOrder("LastAction") = LastAction
            drWSIOSOrder("ActionStatus") = ActionStatus
            drWSIOSOrder("OrderVolume") = OrderVolume
            drWSIOSOrder("OrderPrice") = OrderPrice
            drWSIOSOrder("RemainingVolume") = RemainingVolume
            drWSIOSOrder("DoneVolumeTotal") = DoneVolumeTotal
            drWSIOSOrder("AveragePrice") = AveragePrice
            drWSIOSOrder("Lifetime") = Lifetime
            drWSIOSOrder("ExecutionInstructions") = ExecutionInstructions
            drWSIOSOrder("Currency") = Currency
            drWSIOSOrder("PrimaryClientOrderId") = PrimaryClientOrderId
            drWSIOSOrder("PostTradeStatusNumber") = PostTradeStatusNumber
            drWSIOSOrder("FXrate") = DFXrate
            drWSIOSOrder("Organization") = sOrganization
            '            CreateDateTime
            drWSIOSOrder("CreateDateTime") = CreateDateTime
            drWSIOSOrder("UpdateDateTime") = TradeTime
            drWSIOSOrder("SecurityType") = iSecurityType
            'iOrderFlags
            drWSIOSOrder("OrderFlags") = iOrderFlags

            dtWSIOSOrdersTable.Rows.Add(drWSIOSOrder)
        Catch ex As Exception
            Call LogToFile("  Error: AddWSIOSOrder - " & ex.Message)
        End Try

    End Sub

    Private Sub AddWSTrades(ByVal TradeNumber As Long, ByVal OrderNumber As Long, ByVal RootOrderNumber As Long, ByVal Exchange As String, ByVal Destination As String, ByVal TradeVolume As Long, ByVal TradePrice As Decimal,
                            ByVal DFXrate As Decimal, ByVal TradeTime As DateTime, ByVal OpposingBrokerNumber As Long, ByVal TradeMarkers As String, ByVal SourcePrice As Decimal, ByVal Currency As String,
                            ByVal AccountCode As String, ByVal SecurityCode As String)
        Dim drWSIOSTrades As DataRow
        Try

            'If Exchange = "LSE" Or Exchange = "ASX" Or Exchange = "JSE" Then
            '    TradePrice = TradePrice / 100
            '    SourcePrice = SourcePrice / 100
            'End If
            'RL008
            Dim Multipier As Decimal = 1
            If htExchangeMultipliers.ContainsKey(Exchange) Then
                Multipier = htExchangeMultipliers.Item(Exchange)
            End If
            TradePrice = TradePrice * Multipier
            SourcePrice = SourcePrice * Multipier
            drWSIOSTrades = dtWSTradesTable.NewRow()

            drWSIOSTrades("TradeNumber") = TradeNumber
            drWSIOSTrades("OrderNumber") = OrderNumber
            drWSIOSTrades("RootOrderNumber") = RootOrderNumber

            drWSIOSTrades("Exchange") = Exchange
            drWSIOSTrades("Destination") = Destination
            drWSIOSTrades("TradeVolume") = TradeVolume
            drWSIOSTrades("TradePrice") = TradePrice
            drWSIOSTrades("FXrate") = DFXrate
            drWSIOSTrades("TradeDateTime") = TradeTime
            drWSIOSTrades("OpposingBrokerNumber") = OpposingBrokerNumber
            drWSIOSTrades("TradeMarkers") = TradeMarkers
            drWSIOSTrades("SourcePrice") = SourcePrice
            drWSIOSTrades("AccountCode") = AccountCode
            drWSIOSTrades("SecurityCode") = SecurityCode

            dtWSTradesTable.Rows.Add(drWSIOSTrades)
        Catch ex As Exception
            Call LogToFile("  Error: AddWSTrades - " & ex.Message)
        End Try

    End Sub
    Private Sub AddReportOrder(ByVal AccCode As String, ByVal Dest As String, ByVal ActStat As String,
          ByVal LastAct As String, ByVal BuySell As String, ByVal SecCode As String, ByVal OrdPrc As Decimal, ByVal TrdPrc As Decimal,
          ByVal PrcInst As String, ByVal Lifetime As String, ByVal OrdVol As Long, ByVal TrdVol As Long, ByVal DoneVolTot As Long,
          ByVal RemVol As Long, ByVal AvgPrc As Decimal, ByVal AccType As String, ByVal ExecInstr As String,
          ByVal PostTradeStatus As String, ByVal OrdNo As Long, ByVal RtOrdNo As Long, ByVal TradeNo As Long, ByVal PriCliOrd As String, ByVal sEXBR As String, ByVal sSettlementCurrency As String,
          ByVal DFXrate As Decimal, ByVal sOrganization As String, ByVal TradeTime As String, ByVal SettlementTime As String, ByVal SecurityType As String, ByVal Currency As String, ByVal Exchange As String, ByVal cusip As String,
  ByVal ISIN As String, ByVal SEDOL As String, ByVal Description As String, ByVal DFANote1 As String, ByVal DFANote2 As String, ByVal DFANote3 As String, ByVal OpenClose As String, ByVal UkIrishStampDutyReserveTaxMarker As String,
  ByVal PtmLevyIndicator As String, ByVal Source As String, ByVal CurrencyDenomination As String)
        Dim drReportOrder As DataRow
        Try

            If htExchangeCurrencys.ContainsKey(Exchange) Then
                sSettlementCurrency = htExchangeCurrencys.Item(Exchange)
                Currency = htExchangeCurrencys.Item(Exchange)
                CurrencyDenomination = htExchangeCurrencys.Item(Exchange)
            End If

            drReportOrder = dtReportOrdersTable.NewRow()

            drReportOrder("AccCode") = AccCode
            drReportOrder("Dest") = Dest
            drReportOrder("ActStat") = ActStat
            drReportOrder("LastAct") = LastAct
            drReportOrder("BuySell") = BuySell
            drReportOrder("SecCode") = SecCode

            drReportOrder("OrdVol") = OrdVol
            drReportOrder("OrdPrc") = OrdPrc

            drReportOrder("TrdVol") = TrdVol
            drReportOrder("TrdPrc") = TrdPrc

            drReportOrder("PrcInst") = PrcInst
            drReportOrder("Lifetime") = Lifetime


            drReportOrder("DoneVolTot") = DoneVolTot
            drReportOrder("RemVol") = RemVol
            drReportOrder("AvgPrc") = AvgPrc
            drReportOrder("AccType") = AccType
            drReportOrder("ExecInstr") = ExecInstr
            drReportOrder("PostTradeStatus") = PostTradeStatus
            drReportOrder("RootOrdNo") = RtOrdNo
            drReportOrder("OrdNo") = OrdNo
            drReportOrder("TradeNo") = TradeNo
            drReportOrder("PriCliOrd") = PriCliOrd
            drReportOrder("EXBR") = sEXBR
            drReportOrder("SettlementCurrency") = sSettlementCurrency
            drReportOrder("FXrate") = DFXrate
            drReportOrder("Organization") = sOrganization
            drReportOrder("TradeTime") = TradeTime

            drReportOrder("OpenClose") = OpenClose

            drReportOrder("SettlementTime") = SettlementTime
            drReportOrder("SecurityType") = SecurityType
            'Currency
            drReportOrder("Currency") = Currency
            drReportOrder("Exchange") = Exchange
            drReportOrder("ISIN") = ISIN
            drReportOrder("CUSIP") = cusip
            drReportOrder("SEDOL") = SEDOL
            drReportOrder("Description") = Description

            drReportOrder("DFANote1") = DFANote1
            drReportOrder("DFANote2") = DFANote2
            drReportOrder("DFANote3") = DFANote3


            drReportOrder("UkIrishStampDutyReserveTaxMarker") = UkIrishStampDutyReserveTaxMarker
            drReportOrder("PtmLevyIndicator") = PtmLevyIndicator

            drReportOrder("Source") = Source
            drReportOrder("CurrencyDenomination") = CurrencyDenomination

            dtReportOrdersTable.Rows.Add(drReportOrder)

        Catch ex As Exception
            Call LogToFile("  Error: AddReportOrder - " & ex.Message)
        End Try
    End Sub

    Private Sub AddNettingReportOrder(ByVal AccCode As String, ByVal Dest As String, ByVal ActStat As String,
          ByVal LastAct As String, ByVal BuySell As String, ByVal SecCode As String, ByVal OrdPrc As Decimal, ByVal TrdPrc As Decimal,
          ByVal PrcInst As String, ByVal Lifetime As String, ByVal OrdVol As Long, ByVal TrdVol As Long, ByVal DoneVolTot As Long,
          ByVal RemVol As Long, ByVal AvgPrc As Decimal, ByVal AccType As String, ByVal ExecInstr As String,
          ByVal PostTradeStatus As String, ByVal OrdNo As Long, ByVal RtOrdNo As Long, ByVal TradeNo As Long, ByVal PriCliOrd As String, ByVal sEXBR As String, ByVal sSettlementCurrency As String,
          ByVal DFXrate As Decimal, ByVal sOrganization As String, ByVal TradeTime As String, ByVal SettlementTime As String, ByVal SecurityType As String, ByVal Currency As String, ByVal Exchange As String, ByVal cusip As String,
  ByVal ISIN As String, ByVal SEDOL As String, ByVal Description As String, ByVal DFANote1 As String, ByVal DFANote2 As String, ByVal DFANote3 As String, ByVal OpenClose As String, ByVal UkIrishStampDutyReserveTaxMarker As String, ByVal PtmLevyIndicator As String, ByVal Source As String, ByVal CurrencyDenomination As String)
        Dim drNettingReportOrder As DataRow
        Try


            drNettingReportOrder = dtNettingReportOrdersTable.NewRow()

            drNettingReportOrder("AccCode") = AccCode
            drNettingReportOrder("Dest") = Dest
            drNettingReportOrder("ActStat") = ActStat
            drNettingReportOrder("LastAct") = LastAct
            drNettingReportOrder("BuySell") = BuySell
            drNettingReportOrder("SecCode") = SecCode

            drNettingReportOrder("OrdVol") = OrdVol
            drNettingReportOrder("OrdPrc") = OrdPrc

            drNettingReportOrder("TrdVol") = TrdVol
            drNettingReportOrder("TrdPrc") = TrdPrc

            drNettingReportOrder("PrcInst") = PrcInst
            drNettingReportOrder("Lifetime") = Lifetime


            drNettingReportOrder("DoneVolTot") = DoneVolTot
            drNettingReportOrder("RemVol") = RemVol
            drNettingReportOrder("AvgPrc") = AvgPrc
            drNettingReportOrder("AccType") = AccType
            drNettingReportOrder("ExecInstr") = ExecInstr
            drNettingReportOrder("PostTradeStatus") = PostTradeStatus
            drNettingReportOrder("RootOrdNo") = RtOrdNo
            drNettingReportOrder("OrdNo") = OrdNo
            drNettingReportOrder("TradeNo") = TradeNo
            drNettingReportOrder("PriCliOrd") = PriCliOrd
            drNettingReportOrder("EXBR") = sEXBR
            drNettingReportOrder("SettlementCurrency") = sSettlementCurrency
            drNettingReportOrder("FXrate") = DFXrate
            drNettingReportOrder("Organization") = sOrganization
            drNettingReportOrder("TradeTime") = TradeTime

            drNettingReportOrder("OpenClose") = OpenClose

            drNettingReportOrder("SettlementTime") = SettlementTime
            drNettingReportOrder("SecurityType") = SecurityType
            'Currency
            drNettingReportOrder("Currency") = Currency
            drNettingReportOrder("Exchange") = Exchange
            drNettingReportOrder("ISIN") = ISIN
            drNettingReportOrder("CUSIP") = cusip
            drNettingReportOrder("SEDOL") = SEDOL
            drNettingReportOrder("Description") = Description

            drNettingReportOrder("DFANote1") = DFANote1
            drNettingReportOrder("DFANote2") = DFANote2
            drNettingReportOrder("DFANote3") = DFANote3


            drNettingReportOrder("UkIrishStampDutyReserveTaxMarker") = UkIrishStampDutyReserveTaxMarker
            drNettingReportOrder("PtmLevyIndicator") = PtmLevyIndicator
            drNettingReportOrder("Source") = Source
            drNettingReportOrder("CurrencyDenomination") = CurrencyDenomination

            dtNettingReportOrdersTable.Rows.Add(drNettingReportOrder)

        Catch ex As Exception
            Call LogToFile("  Error: AddNettingReportOrder - " & ex.Message)
        End Try
    End Sub

  Private Sub GetTradeByOrderNo(ByVal Session As Integer)
    Dim iIndex As Integer
    Dim dvWSIOSOrdersView As DataView
    Dim drWSIOSOrderRow As Data.DataRowView
    Dim lOrderNo, lRootOrderNo As Long
    Dim lRemainingVolume As Long
    Try
      dvWSIOSOrdersView = New DataView(dtWSIOSOrdersTable)
      dvWSIOSOrdersView.Sort = "OrderNumber"

      For iIndex = 0 To dvWSIOSOrdersView.Count - 1
        drWSIOSOrderRow = dvWSIOSOrdersView.Item(iIndex)

        lOrderNo = drWSIOSOrderRow("OrderNumber")

        lRootOrderNo = drWSIOSOrderRow("OrderNumber")

        lRemainingVolume = drWSIOSOrderRow("RemainingVolume")
        'Call LogToFile(drWSIOSOrderRow("OrderNumber") & ":" & drWSIOSOrderRow("CreateDateTime") & ":" & dtWSTradesTable.Rows.Count)
        ' KC001
        'Call GetWebServiceIOSTradesByOrderNo(lOrderNo, lRootOrderNo, lRemainingVolume, drWSIOSOrderRow("CreateDateTime"))
        Call GetWebServiceIOSTradesByOrderNo(lOrderNo, lRootOrderNo, lRemainingVolume, Session, drWSIOSOrderRow("CreateDateTime"))

      Next
      'Call LogToFile("  info: GetWebServiceIOSTradesByOrderNo - " & dtWSTradesTable.Rows.Count)

    Catch ex As Exception
      Call LogToFile("  Error: GetTradeByOrderNo (Source" & CStr(0) &
") - Unable to create DataView - " & ex.Message)
      Exit Sub
    End Try

  End Sub
    Private Sub ProcessWSIOSOrdersDataTable(ByVal Session As Integer)
        Dim dvWSIOSOrdersView As DataView
        Dim drWSIOSOrderRow As Data.DataRowView
        Dim dvWSIOSOrdersView2 As DataView
        Dim drWSIOSOrderRow2 As Data.DataRowView
        Dim dvWSIOSTradesView As DataView
        Dim drWSIOSTradeRow As Data.DataRowView
        Dim iIndex, iIndex2 As Integer
        Dim lOrderNo, lRootParentOrder, lTradeNo As Long
        Dim sOrderNo As String
        Dim sDestination As String = ""
        Dim dtTradeDateTime, dtDFDDateTime As Date
        Dim sType As String = "Market"
        Dim bHasClient As Boolean

        Try
            dvWSIOSOrdersView = New DataView(dtWSIOSOrdersTable)
            dvWSIOSOrdersView.Sort = "OrderNumber"



        Catch ex As Exception

            Call LogToFile("  Error: ProcessWSIOSOrdersDataTable (Source" & CStr(Session) &
              ") - Unable to create DataView - " & ex.Message)
            Exit Sub
        End Try


        For iIndex = 0 To dvWSIOSOrdersView.Count - 1
            Try
                drWSIOSOrderRow = dvWSIOSOrdersView.Item(iIndex)



                dvWSIOSTradesView = New DataView(dtWSTradesTable)
                dvWSIOSTradesView.Sort = "OrderNumber"
                lOrderNo = drWSIOSOrderRow("OrderNumber")
                lRootParentOrder = drWSIOSOrderRow("RootOrderNumber")
                sOrderNo = lOrderNo.ToString

                'RLU003
                'dvWSIOSTradesView.RowFilter = "OrderNumber = '" & sOrderNo & "' OR RootOrderNumber ='" & sOrderNo & "'"
                dvWSIOSTradesView.RowFilter = "OrderNumber = '" & sOrderNo & "'"

                'dvWSIOSTradesView.Sort = "TradeDateTime"
                'RL003
                dvWSIOSTradesView.Sort = "TradeDateTime DESC"

                'TradedOrder Only Condition
                If gbTradedOrderOnly And (dvWSIOSTradesView.Count > 0 Or drWSIOSOrderRow("DoneVolumeTotal") > 0) Then

                Else

                    Continue For
                End If

                'Done For Day Order only Condition

                'RLU003
                'If gbDFD Then
                '    If htDNDOrdersTable.ContainsKey(lOrderNo) Then

                '        dtDFDDateTime = htDNDOrdersTable.Item(lOrderNo)

                '        If dtDFDDateTime > gdStartTime And dtDFDDateTime < gdEndTime Then
                '            'DND mark stamp in Report Time interval
                '        Else
                '            'LogToFile("ProcessWSIOSOrdersDataTable-" & lOrderNo & "-2")

                '            'RLU003
                '            'Continue For
                '        End If
                '    Else
                'LogToFile("ProcessWSIOSOrdersDataTable-" & lOrderNo & "-3")

                '        'RLU003
                '        'Continue For
                '    End If
                'Else

                'End If


                'RLU003 Look for parent order destination

                'lRootParentOrder = drWSIOSOrderRow("RootOrderNumber")
                'dvWSIOSOrdersView2 = New DataView(dtWSIOSOrdersTable)
                'dvWSIOSOrdersView2.Sort = "OrderNumber"
                'dvWSIOSOrdersView2.RowFilter = "OrderNumber = '" & lRootParentOrder & "'"
                'If dvWSIOSOrdersView2.Count > 0 drWSIOSOrderRow2
                '    Then = dvWSIOSOrdersView2.Item(0)
                '    sDestination = drWSIOSOrderRow2("Destination")
                '    If sDestination <> "DESK" Then
                '        sDestination = drWSIOSOrderRow("Destination")
                '    End If
                'End If

#If DEBUG Then


#End If


                dtTradeDateTime = drWSIOSOrderRow("UpdateDateTime")

                If dvWSIOSTradesView.Count > 0 Then


                    ''RLU003
                    'For iIndex2 = 0 To dvWSIOSTradesView.Count - 1
                    '    drWSIOSTradeRow = dvWSIOSTradesView.Item(iIndex2)
                    '    If drWSIOSTradeRow("TradeDateTime") < dtDFDDateTime Then
                    '        dtTradeDateTime = drWSIOSTradeRow("TradeDateTime")
                    '    Else
                    '        'LogToFile("ProcessWSIOSOrdersDataTable-" & lOrderNo & "-4")

                    '        'RLU003
                    '        Exit For
                    '    End If
                    'Next

                    For iIndex2 = 0 To dvWSIOSTradesView.Count - 1
                        drWSIOSTradeRow = dvWSIOSTradesView.Item(iIndex2)
                        dtTradeDateTime = drWSIOSTradeRow("TradeDateTime")
                        lTradeNo = drWSIOSTradeRow("TradeNumber")


                        'RL003
                        'If htDNDOrdersTable.ContainsKey(lOrderNo) Or htDNDOrdersTable.ContainsKey(lRootParentOrder) Then
                        If True Then

                            'RL003
                            'If htDNDOrdersTable.ContainsKey(lOrderNo) Then
                            '    dtDFDDateTime = htDNDOrdersTable.Item(lOrderNo)
                            'Else

                            '    dtDFDDateTime = htDNDOrdersTable.Item(lRootParentOrder)
                            'End If

                            'RL003
                            'If dtTradeDateTime < dtDFDDateTime Then
                            If True Then
                                sType = "Client"
                                bHasClient = True

                                dvWSIOSOrdersView2 = New DataView(dtWSIOSOrdersTable)
                                dvWSIOSOrdersView2.RowFilter = "OrderNumber = '" & lRootParentOrder.ToString & "'"
                                For iIndex3 = 0 To dvWSIOSOrdersView2.Count - 1
                                    drWSIOSOrderRow2 = dvWSIOSOrdersView2.Item(iIndex3)
                                    sDestination = drWSIOSOrderRow2("Destination")
                                Next
                                'RLU 000
                                If drWSIOSOrderRow("SecurityCode") = "EMH" Or drWSIOSOrderRow("SecurityCode") = "CGC" Then
                                    Call LogToFile("Info: Sec:" & drWSIOSOrderRow("SecurityCode") & " OrdNo:" & drWSIOSOrderRow("OrderNumber") & " TrdNo" & drWSIOSTradeRow("TradeNumber"))
                                End If
                                Call ProcessWSIOSOrderRecord(drWSIOSOrderRow("ExecutionInstructions"), drWSIOSOrderRow("AccountCode"),
                                                                 sDestination, sType, bHasClient, drWSIOSOrderRow("Exchange"), drWSIOSOrderRow("Currency"),
                                                                 drWSIOSOrderRow("SecurityCode"), drWSIOSOrderRow("PostTradeStatusNumber"), drWSIOSOrderRow("ActionStatus"),
                                                                 drWSIOSOrderRow("LastAction"), drWSIOSOrderRow("BuyOrSell"), drWSIOSOrderRow("OrderPrice"), drWSIOSTradeRow("TradePrice"),
                                                                 drWSIOSOrderRow("PricingInstructions"), drWSIOSOrderRow("Lifetime"), drWSIOSOrderRow("OrderVolume"), drWSIOSTradeRow("TradeVolume"),
                                                                 drWSIOSOrderRow("DoneVolumeTotal"), drWSIOSOrderRow("RemainingVolume"), drWSIOSOrderRow("AveragePrice"),
                                                                 lOrderNo, drWSIOSOrderRow("RootOrderNumber"), lTradeNo, drWSIOSOrderRow("PrimaryClientOrderId"), drWSIOSOrderRow("FXrate"), drWSIOSOrderRow("Organization"),
                                                                 drWSIOSOrderRow("CreateDateTime"), drWSIOSOrderRow("UpdateDateTime"), dtTradeDateTime, drWSIOSOrderRow("SecurityType"), drWSIOSOrderRow("OrderFlags"))
                            End If
                        End If

                        sType = "Market"



                        Call ProcessWSIOSOrderRecord(drWSIOSOrderRow("ExecutionInstructions"), drWSIOSOrderRow("AccountCode"),
                              drWSIOSOrderRow("Destination"), sType, bHasClient, drWSIOSOrderRow("Exchange"), drWSIOSOrderRow("Currency"),
                              drWSIOSOrderRow("SecurityCode"), drWSIOSOrderRow("PostTradeStatusNumber"), drWSIOSOrderRow("ActionStatus"),
                              drWSIOSOrderRow("LastAction"), drWSIOSOrderRow("BuyOrSell"), drWSIOSOrderRow("OrderPrice"), drWSIOSTradeRow("TradePrice"),
                              drWSIOSOrderRow("PricingInstructions"), drWSIOSOrderRow("Lifetime"), drWSIOSOrderRow("OrderVolume"), drWSIOSTradeRow("TradeVolume"),
                              drWSIOSOrderRow("DoneVolumeTotal"), drWSIOSOrderRow("RemainingVolume"), drWSIOSOrderRow("AveragePrice"),
                              drWSIOSOrderRow("OrderNumber"), drWSIOSOrderRow("RootOrderNumber"), lTradeNo, drWSIOSOrderRow("PrimaryClientOrderId"), drWSIOSOrderRow("FXrate"), drWSIOSOrderRow("Organization"),
                            drWSIOSOrderRow("CreateDateTime"), drWSIOSOrderRow("UpdateDateTime"), dtTradeDateTime, drWSIOSOrderRow("SecurityType"), drWSIOSOrderRow("OrderFlags"))

                        'RLU003 Remove dup  trade on test server  
                        'iIndex2 = dvWSIOSTradesView.Count
                    Next


                Else
                    Continue For
                End If

                ' LogToFile("ProcessWSIOSOrdersDataTable-" & dtReportOrdersTable.Rows.Count)

            Catch ex As Exception
                Call LogToFile("  Error: ProcessWSIOSOrdersDataTable (Source" & CStr(Session) & ") - Unable to read row (" &
                  CStr(iIndex) & ") - " & ex.Message)
            End Try
        Next

        dtWSIOSOrdersTable.Clear()
    End Sub



    Private Sub ProcessWSIOSOrderRecord(ByVal ExecutionInstructions As String, ByVal AccountCode As String,
      ByVal Destination As String, ByVal Type As String, ByVal HasClient As Boolean, ByVal Exchange As String, ByVal Currency As String, ByVal SecurityCode As String,
      ByVal PostTradeStatusNumber As Long, ByVal ActionStatus As String, ByVal LastAction As String,
      ByVal BuyOrSell As String, ByVal OrderPrice As Decimal, ByVal TradePrice As Decimal, ByVal PricingInstructions As String,
      ByVal Lifetime As String, ByVal OrderVolume As Long, ByVal TradeVolume As Long, ByVal DoneVolumeTotal As Long, ByVal RemainingVolume As Long,
      ByVal AveragePrice As Decimal, ByVal OrderNumber As Long, ByVal lRootOrderNumber As Long, ByVal lTradeNumber As Long, ByVal PrimaryClientOrderId As String, ByVal FXrate As Decimal,
      ByVal sOrganization As String, ByVal CreateDateTime As DateTime, ByVal UpdateDateTime As DateTime, ByVal TradeDatetime As DateTime, ByVal iSecurityType As Integer, ByVal OrderFlags As Long)

        Dim sAccountType As String = Type
        Dim sAccountId As String
        Dim sPostTradeStatus As String
        Dim sSettlementCurrency As String
        Dim sEXBR As String
        Dim sSecurityType As String
        Dim settlementdate As Date
        Dim sTradeTime As String = ""
        Dim sSettlementdate As String = ""
        Dim CUSIP As String = ""
        Dim ISIN As String = ""
        Dim SEDOL As String = ""
        Dim Description As String = ""
        Dim DFANote1 As String = ""
        Dim DFANote2 As String = ""
        Dim DFANote3 As String = ""
        Dim Source As String = ""
        Dim OpenClose As String = ""
        Dim UkIrishStampDutyReserveTaxMarker As String = ""
        Dim PtmLevyIndicator As String = ""
        Dim dtDNDTime As Date
        Dim CurrencyDenomination As String = ""

        ''Checck If Order Exist in report
        'If htReportOrdersTable.ContainsKey(OrderNumber) Then
        '    Exit Sub
        'Else
        '    'Check DFD
        '    If gbDFD Then
        '        If (OrderFlags And 1) = 1 Then
        '        Else
        '            Exit Sub
        '        End If
        '    End If
        '    htReportOrdersTable.Add(OrderNumber, True)
        'End If


        'sAccountType = GetAccountType(ExecutionInstructions)                    'account	type

        ''RLU003
        'sAccountType = "Market"
        'If Destination = "DESK" Then
        '    sAccountType = "Client"
        'End If


        sAccountId = UCase(Trim(GetTaggedFieldFromTMXI("", "O", "Account ID", ExecutionInstructions)))

        'settlement currency 
        sSettlementCurrency = UCase(Trim(GetTaggedFieldFromTMXI("", "ITS", "Settlement Currency", ExecutionInstructions)))
        If sSettlementCurrency = "" Then
            sSettlementCurrency = Currency
        End If
        OpenClose = UCase(Trim(GetTaggedFieldFromTMXI("", "POS", "position", ExecutionInstructions)))

        'EXBR
        sEXBR = UCase(Trim(GetTaggedFieldFromTMXI("", "IOBN", "ExecBroker", ExecutionInstructions)))
        If sEXBR = "" Then
            If htReportOrdersTable.ContainsKey(OrderNumber) Then
                sEXBR = htReportOrdersTable.Item(OrderNumber)

            ElseIf htReportOrdersTable.ContainsKey(lRootOrderNumber) Then
                sEXBR = htReportOrdersTable.Item(lRootOrderNumber)
            End If


        End If


        Select Case iSecurityType
            Case 100 To 199
                sSecurityType = "Equity"
            Case 400 To 499
                sSecurityType = "Fixed Income"
            Case 501 To 520
                sSecurityType = "Equity Option"
            Case 521 To 540
                sSecurityType = "Futures Option"
            Case 600 To 699
                sSecurityType = "Future"
            Case 700 To 799
                sSecurityType = "Index"
            Case 800 To 899
                sSecurityType = "FX"
            Case 1000 To 1099
                sSecurityType = "Commodity"
            Case 1100 To 1199
                sSecurityType = "Fund"
            Case Else
                sSecurityType = "Other"
        End Select

        If gsDefaultReportOrderTypes.StartsWith("*ALL") Then
            If htExceptionOrderTypes.ContainsKey(sSecurityType.ToUpper.Trim.Replace(" ", "")) Then
                Exit Sub
            End If
        ElseIf Not htOrderTypes.ContainsKey(sSecurityType.ToUpper.Replace(" ", "")) Then
            Exit Sub
        End If

        sSettlementdate = UCase(Trim(GetTaggedFieldFromTMXI("", "ISD", "Settlement Date", ExecutionInstructions)))

        'LogToFile("ProcessWSIOSOrderRecord-2")
        'LogToFile("ProcessWSIOSOrderRecord-3" & TradeDatetime)
        'LogToFile("ProcessWSIOSOrderRecord-3" & gdStartTime)
        'LogToFile("ProcessWSIOSOrderRecord-3" & gdEndTime)

    ' KC001
    'If (TradeDatetime > gdEndTime Or TradeDatetime < gdStartTime) And Type = "Market" Then
    '    Exit Sub
    '    'ElseIf TradeDatetime > gdEndTime And TradeDatetime < gdDate Then
    '    '    Exit Sub
    'End If

        'If DoneVolumeTotal > 0 And sSettlementdate = "" Then
        '    settlementdate = New Date(UpdateDateTime.Year, UpdateDateTime.Month, UpdateDateTime.Day + 2, UpdateDateTime.Hour, UpdateDateTime.Minute, UpdateDateTime.Second)

        '    If TradeDatetime <> UpdateDateTime Then
        '        sTradeTime = TradeDatetime.ToString

        '        settlementdate = GetSettlementDate(Exchange, TradeDatetime, 2)
        '        sSettlementdate = settlementdate.ToString
        '    End If

        'End If

        'LogToFile("ProcessWSIOSOrderRecord-3")

    ' KC006
    '' KC004
    'sTradeTime = TradeDatetime.ToString("MM/dd/yyyy HH:mm:ss")
    sTradeTime = TradeDatetime.ToString(gsSettleTradeDateFormat & " HH:mm:ss", _
      System.Globalization.CultureInfo.InvariantCulture)

    'RLU003
    If sSettlementdate = "" Then
      ' KC004
      'sTradeTime = TradeDatetime.ToString
      If TradeVolume > 0 Then
        ' KC003
        'settlementdate = GetSettlementDate(Exchange, TradeDatetime, 2)
        Dim iSettlementDays As Integer
        If MatchExchange(Exchange, gsTplus3Exchanges) Then
          iSettlementDays = 3
        Else
          iSettlementDays = 2
        End If
        settlementdate = GetSettlementDate(Exchange, gdDate, iSettlementDays)

        ' KC006
        'sSettlementdate = settlementdate.ToString("MM/dd/yyyy")
        sSettlementdate = settlementdate.ToString(gsSettleTradeDateFormat, System.Globalization.CultureInfo.InvariantCulture)
      End If
    End If

        sPostTradeStatus = MapPostTradeStatus(PostTradeStatusNumber)
        'Call AddReportOrder(AccountCode, Destination, ActionStatus, LastAction, BuyOrSell, SecurityCode, OrderPrice,
        'PricingInstructions, Lifetime, OrderVolume, DoneVolumeTotal, RemainingVolume, AveragePrice, sAccountType,
        'ExecutionInstructions, sPostTradeStatus, OrderNumber, PrimaryClientOrderId, sEXBR, sSettlementCurrency, DFXrate, sOrganization, sTradeTime, sSettlementdate, sSecurityType, Currency, Exchange)

        Call ProcessWSSec(SecurityCode, Exchange, CUSIP, ISIN, SEDOL, Description, UkIrishStampDutyReserveTaxMarker, PtmLevyIndicator, CurrencyDenomination)

        If CurrencyDenomination = "" Or CurrencyDenomination Is Nothing Then
            CurrencyDenomination = Currency
        End If

        DFANote1 = UCase(Trim(GetTaggedFieldFromTMXI("", "DFANote1", "DFANotes1", ExecutionInstructions)))
        DFANote2 = UCase(Trim(GetTaggedFieldFromTMXI("", "DFANote2", "DFANotes2", ExecutionInstructions)))
        DFANote3 = UCase(Trim(GetTaggedFieldFromTMXI("", "DFANote3", "DFANotes3", ExecutionInstructions)))
        'LogToFile("ProcessWSIOSOrderRecord-4")

        Source = UCase(Trim(GetTaggedFieldFromTMXI("", "SOURCE", "Source", ExecutionInstructions)))
        Source = Source.Trim.ToUpper

        If Source <> "FIX" And Source <> "VOICE" And Source <> "CARE" Then
            Source = ""
        End If

        'RL002
        If UkIrishStampDutyReserveTaxMarker = "0" Then
            UkIrishStampDutyReserveTaxMarker = "STAMPEXEMPT"
        Else
            UkIrishStampDutyReserveTaxMarker = ""
        End If
        If PtmLevyIndicator = "3" Then
            PtmLevyIndicator = "PTMEXEMPT"
        Else
            PtmLevyIndicator = ""
        End If

        'RL004
        If HasClient Then
            If DFANote2 = "EXCLUDE" Then
                UkIrishStampDutyReserveTaxMarker = "STAMPEXEMPT"
            End If
        End If

        If MatchReportAccount(AccountCode, Destination, Exchange, Currency, sAccountType, sAccountId, SecurityCode) Then
            sPostTradeStatus = MapPostTradeStatus(PostTradeStatusNumber)

            'LogToFile("ProcessWSIOSOrderRecord-5")

            Call AddReportOrder(AccountCode, Destination, ActionStatus, LastAction, BuyOrSell, SecurityCode, OrderPrice, TradePrice,
                      PricingInstructions, Lifetime, OrderVolume, TradeVolume, DoneVolumeTotal, RemainingVolume, AveragePrice, sAccountType,
                      ExecutionInstructions, sPostTradeStatus, OrderNumber, lRootOrderNumber, lTradeNumber, PrimaryClientOrderId, sEXBR, sSettlementCurrency,
                                    FXrate, sOrganization, sTradeTime, sSettlementdate, sSecurityType, Currency, Exchange,
                                    CUSIP, ISIN, SEDOL, Description, DFANote1, DFANote2, DFANote3, OpenClose, UkIrishStampDutyReserveTaxMarker, PtmLevyIndicator, Source, CurrencyDenomination)


        End If

    End Sub

    Private Sub ProcessWSSec(ByVal SecCode As String, ByVal exchange As String, ByRef CUSIP As String, ByRef ISIN As String, ByRef SEDOL As String,
                             ByRef Description As String, ByRef UkIrishStampDutyReserveTaxMarker As String, ByRef PtmLevyIndicator As String, ByRef CurrencyDenomination As String)
        Dim sKey As String
        Dim iIndex As Integer
        Dim SecurityCode(1) As String
        Dim exchanges(1) As String

        SecurityCode(0) = SecCode
        exchanges(0) = exchange
        sKey = SecCode & "-" & exchange

        Try

            If htWSSecTable.ContainsKey(sKey) Then
                iIndex = htWSSecTable.Item(sKey)
                CUSIP = gaWSSecList(iIndex).CUSIP
                ISIN = gaWSSecList(iIndex).ISIN
                SEDOL = gaWSSecList(iIndex).SEDOL
                Description = gaWSSecList(iIndex).Description

                'RL002
                UkIrishStampDutyReserveTaxMarker = gaWSSecList(iIndex).UkIrishStampDutyReserveTaxMarker
                PtmLevyIndicator = gaWSSecList(iIndex).PtmLevyIndicator
                CurrencyDenomination = gaWSSecList(iIndex).CurrencyDenomination
                Exit Sub
            End If

        Catch ex As Exception
            Call LogToFile("  Error: ProcessWSSec - Unable to find secInfo - " & ex.Message)
            Exit Sub
        End Try

        Try
            Call GetWebServiceSecurity(SecurityCode, exchanges, CUSIP, ISIN, SEDOL, Description, UkIrishStampDutyReserveTaxMarker, PtmLevyIndicator, CurrencyDenomination)

            Try
                htWSSecTable.Add(sKey, giWsSecIndex)




                gaWSSecList(giWsSecIndex).CUSIP = CUSIP
                gaWSSecList(giWsSecIndex).ISIN = ISIN
                gaWSSecList(giWsSecIndex).SEDOL = SEDOL
                gaWSSecList(giWsSecIndex).Description = Description

                'RL002
                gaWSSecList(giWsSecIndex).UkIrishStampDutyReserveTaxMarker = UkIrishStampDutyReserveTaxMarker
                gaWSSecList(giWsSecIndex).PtmLevyIndicator = PtmLevyIndicator
                gaWSSecList(giWsSecIndex).CurrencyDenomination = CurrencyDenomination

                giWsSecIndex = giWsSecIndex + 1

            Catch ex As Exception

                Call LogToFile("  Error: ProcessWSSec - Unable to add to table (" & sKey & ") - " & ex.Message)
                Call LogToFile(giWsSecIndex)
            End Try

        Catch ex As Exception
            Call LogToFile("  Error: ProcessWSSec - Unable to read row (" & CStr(giWsSecIndex) & ") - " & ex.Message)
        End Try

    End Sub



End Class