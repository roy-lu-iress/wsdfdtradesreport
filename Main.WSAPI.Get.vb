Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain

    Private Sub GetWebServiceIOSOrdersSearchByUser()
        Dim saDestination() As String = {}
        Dim lOrderNumber As Long
        Dim lRootOrderNumber As Long
        Dim sAccountCode As String
        Dim sSecurityCode As String = ""
        Dim sExchange As String
        Dim sDestination, sEXBR As String
        Dim sBuyOrSell As String = ""
        Dim sPricingInstructions As String = ""
        Dim sLastAction As String
        Dim sActionStatus As String = ""
        Dim lOrderVolume As Long
        Dim decOrderPrice As Decimal
        Dim lRemainingVolume As Long
        Dim lDoneVolumeTotal As Long
        Dim decAveragePrice As Decimal
        Dim sLifetime As String
        Dim sExecutionInstructions As String = ""
        Dim sCurrency As String = ""
        Dim sPrimaryClientOrderId As String = ""
        Dim lPostTradeStatusNumber As Long

        Dim dFXrate As Decimal
        Dim sOrganisation As String = ""
        Dim CreateDateTime As Date
        Dim UpdateDateTime As Date
        Dim iSecurityType As Integer

        Dim iIndex As Integer

        Dim lOrderFlagsMask As Long

        Dim sTemp As String

        Try
            ' Retrieve information using OrderOrderPadGetByAccount method
            Dim pqgeRequest As IOSPlus.OrderSearchGetByUserInput = New IOSPlus.OrderSearchGetByUserInput

            ' Initialize the header, use the IOS Plus service session key we got earlier
            pqgeRequest.Header = New IOSPlus.OrderSearchGetByUserInputHeader
            pqgeRequest.Header.ServiceSessionKey = myIOSPlusServiceSessionKey

            Dim guid As New Guid

            ' Initialize the request options
            pqgeRequest.Header.Timeout = 25             ' Timeout after 25 seconds
            pqgeRequest.Header.PageSize = 100           ' Recommended maximum page size
            pqgeRequest.Header.Updates = False          ' Don't watch for updates
            pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
            pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

            ' Initialize the request parameters - set the parameter to retrieve information for
            pqgeRequest.Parameters = New IOSPlus.OrderSearchGetByUserInputParameters

            'RL010
            Dim dPreviousDayStartTime As Date
            dPreviousDayStartTime = ReadIniTime(gsIniFile, APP_NAME, "PreviousDayStartTime", "21:30")
            pqgeRequest.Parameters.DateTimeFrom = CDate(Format(gdDate.AddDays(-1), "yyyy/MM/dd") & " " & dPreviousDayStartTime)
            'Call LogToFile("  Info: PreviousDayStartTime - " & pqgeRequest.Parameters.DateTimeFrom.ToString)
            'pqgeRequest.Parameters.DateTimeFrom = gdDate

            pqgeRequest.Parameters.DateTimeTo = gdDate.AddDays(1)


            Dim pqgeResult As IOSPlus.OrderSearchGetByUserOutput

            ' Call the OrderPadGetByAccount method until we have received all pages of data
            Do
                pqgeResult = wsIOSPlus.OrderSearchGetByUser(pqgeRequest)

                For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)


                    lOrderNumber = pqgeResult.Result.DataRows(iIndex).OrderNumber 'tradeId
                    'Call LogToFile("  info: GetWebServiceIOSOrdersByUser - " & lOrderNumber)
                    lRootOrderNumber = pqgeResult.Result.DataRows(iIndex).RootParentOrderNumber
                    If lRootOrderNumber <> lOrderNumber Then
                        'Continue For
                    End If
                    sAccountCode = UCase(Trim(pqgeResult.Result.DataRows(iIndex).AccountCode))
                    sSecurityCode = UCase(Trim(pqgeResult.Result.DataRows(iIndex).SecurityCode)) 'symbol	
                    sExchange = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Exchange))
                    sDestination = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Destination))
                    sBuyOrSell = UCase(Trim(pqgeResult.Result.DataRows(iIndex).BuyOrSell))   'side()
                    sPricingInstructions = UCase(Trim(pqgeResult.Result.DataRows(iIndex).PricingInstructions))
                    sLastAction = UCase(Trim(pqgeResult.Result.DataRows(iIndex).LastAction))
                    sActionStatus = UCase(Trim(pqgeResult.Result.DataRows(iIndex).ActionStatus))
                    lOrderVolume = pqgeResult.Result.DataRows(iIndex).OrderVolume
                    decOrderPrice = pqgeResult.Result.DataRows(iIndex).OrderPrice
                    lRemainingVolume = pqgeResult.Result.DataRows(iIndex).RemainingVolume
                    lDoneVolumeTotal = pqgeResult.Result.DataRows(iIndex).DoneVolumeTotal
                    decAveragePrice = pqgeResult.Result.DataRows(iIndex).AveragePrice
                    sLifetime = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Lifetime))

                    Try
                       
                        sExecutionInstructions = Trim(pqgeResult.Result.DataRows(iIndex).ExecutionInstructions) 'settlement currency 'EXBR

                        sCurrency = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Currency)) 'tradeCcy
                        sPrimaryClientOrderId = Trim(pqgeResult.Result.DataRows(iIndex).PrimaryClientOrderID)

                        If pqgeResult.Result.DataRows(iIndex).PostTradeStatusNumber IsNot Nothing Then
                            lPostTradeStatusNumber = pqgeResult.Result.DataRows(iIndex).PostTradeStatusNumber
                        End If


                        dFXrate = pqgeResult.Result.DataRows(iIndex).AverageFXRate
                        If dFXrate = 0 Then
                            dFXrate = 1
                        End If
                        sOrganisation = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Organisation))

                        If sOrganisation.Contains("DEFAULT") Or sOrganisation = "" Then
                            sTemp = UCase(Trim(pqgeResult.Result.DataRows(iIndex).RootParentOrderCreatorUserCode))
                            sOrganisation = sTemp.Split("@").Last.ToString

                        End If

                    Catch ex As Exception
                        Call LogToFile("  Error: GetWebServiceIOSOrdersByUser 1- " & ex.Message)
                    End Try


                    Try

                        lOrderFlagsMask = pqgeResult.Result.DataRows(iIndex).OrderFlagsMask
                        'tradeDate 'settlement date

                        UpdateDateTime = pqgeResult.Result.DataRows(iIndex).UpdateDateTime

                        CreateDateTime = pqgeResult.Result.DataRows(iIndex).CreateDateTime
                        'description	

                        iSecurityType = pqgeResult.Result.DataRows(iIndex).SecurityType() 'type	
                    Catch ex As Exception
                        Call LogToFile("  Error: GetWebServiceIOSOrdersByUser 2- " & ex.Message)
                    End Try



                    If sExecutionInstructions.Contains("IOBN") Then
                        sEXBR = UCase(Trim(GetTaggedFieldFromTMXI("", "IOBN", "ExecBroker", sExecutionInstructions)))
                        If sEXBR <> "" Then
                            If Not htReportOrdersTable.ContainsKey(lRootOrderNumber) Then
                                htReportOrdersTable.Add(lRootOrderNumber, sEXBR)
                            End If
                            If Not htReportOrdersTable.ContainsKey(lOrderNumber) Then
                                htReportOrdersTable.Add(lOrderNumber, sEXBR)
                            End If
                        End If
                    End If


                    'Dim bAdd As Boolean = False

                    'If ((lOrderFlagsMask And 1) = 1) And sDestination = "DESK" Then
                    '    bAdd = True
                    'End If

                    'If ((lOrderFlagsMask And 67108864) = 67108864) Then
                    '    bAdd = True
                    'End If

                    If MatchReportAccount(sAccountCode, sDestination, sExchange, sCurrency, "", "", sSecurityCode) Then
                        Call AddWSIOSOrder(lOrderNumber, lRootOrderNumber, sAccountCode, sSecurityCode, sExchange, sDestination, sBuyOrSell,
  sPricingInstructions, sLastAction, sActionStatus, lOrderVolume, decOrderPrice, lRemainingVolume,
  lDoneVolumeTotal, decAveragePrice, sLifetime, sExecutionInstructions, sCurrency, sPrimaryClientOrderId,
  lPostTradeStatusNumber, dFXrate, sOrganisation, CreateDateTime, UpdateDateTime, iSecurityType, lOrderFlagsMask)
                    End If

                    '                  Call AddWSIOSOrder(lOrderNumber, lRootOrderNumber, sAccountCode, sSecurityCode, sExchange, sDestination, sBuyOrSell,
                    'sPricingInstructions, sLastAction, sActionStatus, lOrderVolume, decOrderPrice, lRemainingVolume,
                    'lDoneVolumeTotal, decAveragePrice, sLifetime, sExecutionInstructions, sCurrency, sPrimaryClientOrderId,
                    'lPostTradeStatusNumber, dFXrate, sOrganisation, CreateDateTime, UpdateDateTime, iSecurityType, lOrderFlagsMask)

                    'RLU003
                    'If ((lOrderFlagsMask And 1) = 1 And gbDFD) Or Not gbDFD Then
                    '    Call AddWSIOSOrder(lOrderNumber, lRootOrderNumber, sAccountCode, sSecurityCode, sExchange, sDestination, sBuyOrSell,
                    '      sPricingInstructions, sLastAction, sActionStatus, lOrderVolume, decOrderPrice, lRemainingVolume,
                    '      lDoneVolumeTotal, decAveragePrice, sLifetime, sExecutionInstructions, sCurrency, sPrimaryClientOrderId,
                    '      lPostTradeStatusNumber, DFXrate, sOrganisation, CreateDateTime, UpdateDateTime, iSecurityType, lOrderFlagsMask)
                    'End If


                Next

                Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
            Loop While pqgeResult.Result.Header.StatusCode = 1
            'Call LogToFile(dtWSIOSOrdersTable.Rows.Count)

        Catch ex As Exception
            Call LogToFile("  Error: GetWebServiceIOSOrdersByUser - " & ex.Message)
        End Try
    End Sub

  Private Sub GetWebServiceIOSTradesByOrderNo(ByVal OrderNo As Long, ByVal RootOrderNo As Long, ByVal lRemainingVolume As Long, _
      ByVal Session As Integer, Optional ByVal CreateDateTime As Date = Nothing)
    Dim saDestination() As String = {}
    Dim lOrderNumber, lTradeNumber, lParentOrderNumber As Long
    Dim sAccountCode As String
    Dim sSecurityCode As String
    Dim sExchange As String
    Dim sDestination As String
    Dim lTradeVolume As Long
    Dim decTradePrice As Decimal
    Dim decSourcePrice As Decimal
    Dim sTradeMarkers As String
    Dim sCurrency As String
    Dim lOpposingBrokerNumber As Long
    Dim DFXrate As Decimal
    Dim TradeDateTime As Date
    Dim iIndex As Integer
    Dim lorders() As Long = {OrderNo}


    Try

      ' Retrieve information using TradeGetByUserInput method
      Dim pqgeRequest As IOSPlus.TradeGetByOrderNumberInput = New IOSPlus.TradeGetByOrderNumberInput

      ' Initialize the header, use the IOS Plus service session key we got earlier
      pqgeRequest.Header = New IOSPlus.TradeGetByOrderNumberInputHeader
      pqgeRequest.Header.ServiceSessionKey = myIOSPlusServiceSessionKey

      Dim guid As New Guid

      ' Initialize the request options
      pqgeRequest.Header.Timeout = 25             ' Timeout after 25 seconds
      pqgeRequest.Header.PageSize = 100           ' Recommended maximum page size
      pqgeRequest.Header.Updates = False          ' Don't watch for updates
      pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
      pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

      ' Initialize the request parameters - set the parameter to retrieve information for
      pqgeRequest.Parameters = New IOSPlus.TradeGetByOrderNumberInputParameters
      If CreateDateTime = Nothing Then
        CreateDateTime = New DateTime
        CreateDateTime = Date.Today.AddDays(-10)
      End If

      ' KC001
      'pqgeRequest.Parameters.TradeDateTimeFrom = gdStartTime
      'pqgeRequest.Parameters.TradeDateTimeTo = gdEndTime
      pqgeRequest.Parameters.TradeDateTimeFrom = CDate(Format(gdDate, "yyyy/MM/dd") & " " & "00:00:00")
      pqgeRequest.Parameters.TradeDateTimeTo = CDate(Format(gdDate, "yyyy/MM/dd") & " " & "23:59:59")

      pqgeRequest.Parameters.OrderNumberArray = lorders

      pqgeRequest.Parameters.DateFilterType = 0   ' KC001


      Dim pqgeResult As IOSPlus.TradeGetByOrderNumberOutput

      ' Call the OrderPadGetByAccount method until we have received all pages of data
      Do
        pqgeResult = wsIOSPlus.TradeGetByOrderNumber(pqgeRequest)

        For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)
          Try
            lTradeNumber = pqgeResult.Result.DataRows(iIndex).TradeNumber
            lOrderNumber = pqgeResult.Result.DataRows(iIndex).OrderNumber
            sExchange = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Exchange))
            sDestination = UCase(Trim(pqgeResult.Result.DataRows(iIndex).Destination))

            lTradeVolume = pqgeResult.Result.DataRows(iIndex).TradeVolume
            decTradePrice = pqgeResult.Result.DataRows(iIndex).TradePrice
            DFXrate = pqgeResult.Result.DataRows(iIndex).TradeFXRate

            'RL003
            'If lRemainingVolume = 0 And lTradeVolume > 0 Then
            '    If Not htDNDOrdersTable.ContainsKey(lOrderNumber) Then
            '        htDNDOrdersTable.Add(lOrderNumber, gdEndTime)
            '    End If

            '    If Not htDNDOrdersTable.ContainsKey(RootOrderNo) Then
            '        htDNDOrdersTable.Add(RootOrderNo, gdEndTime)
            '    End If
            'End If

            ' KC004
            'TradeDateTime = pqgeResult.Result.DataRows(iIndex).TradeDateTime
            ''TradeDateTime = pqgeResult.Result.DataRows(iIndex).TradeDateTimeGMT
            'TradeDateTime.AddHours(giTradeServerTimeZone)
            TradeDateTime = pqgeResult.Result.DataRows(iIndex).ExchangeTradeDateTime

          Catch ex As Exception
            Call LogToFile("  Error: GetWebServiceIOSTradesByOrderNo 1- " & ex.Message)
          End Try

          Try

            lOpposingBrokerNumber = 0

            sTradeMarkers = Trim(pqgeResult.Result.DataRows(iIndex).TradeMarkers)

            decSourcePrice = pqgeResult.Result.DataRows(iIndex).SourcePrice

            sCurrency = UCase(Trim(pqgeResult.Result.DataRows(iIndex).SourceCurrency)) 'tradeCcy

            sAccountCode = UCase(Trim(pqgeResult.Result.DataRows(iIndex).AccountCode))
            sSecurityCode = UCase(Trim(pqgeResult.Result.DataRows(iIndex).SecurityCode)) 'symbol	


          Catch ex As Exception
            Call LogToFile("  Error: GetWebServiceIOSTradesByOrderNo 2- " & ex.Message)
          End Try
          lParentOrderNumber = 0

          ' KC001
          If Not IsTradeProcessed(lTradeNumber, Session) Then
            If DicTradeNo.ContainsKey(lTradeNumber) Then

            Else
              DicTradeNo.Add(lTradeNumber, lOrderNumber)

              Call AddWSTrades(lTradeNumber, lOrderNumber, RootOrderNo, sExchange, sDestination, lTradeVolume, decTradePrice, DFXrate,
                               TradeDateTime, lOpposingBrokerNumber, sTradeMarkers, decSourcePrice, sCurrency, sAccountCode, sSecurityCode)
            End If
          End If



        Next

        Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
      Loop While pqgeResult.Result.Header.StatusCode = 1


    Catch ex As Exception
      Call LogToFile("  Error: GetWebServiceIOSTradesByOrderNo - " & ex.Message)
    End Try


  End Sub


    Dim DicTradeNo As New Dictionary(Of Long, Long)
  Dim gDicDeskTradeNo As New Hashtable

  ' KC001
  Private Function IsTradeProcessed(TradeNumber As Long, Session As Integer) As Boolean
    Dim iFileNumber As Integer

    IsTradeProcessed = False

    If gsReportingType = "DELTA" Then
      If htProcessedTradesTable.ContainsKey(CStr(TradeNumber)) Then
        IsTradeProcessed = True
      Else
        ' Write to hash
        Try
          htProcessedTradesTable.Add(CStr(TradeNumber), vbNull)
        Catch ex As Exception
          Call LogToFile("  Error: IsTradeProcessed (Source" & CStr(Session) & ") - Unable to add to table (" & _
            CStr(TradeNumber) & ") - " & ex.Message)
        End Try

        ' Write to trades file
        iFileNumber = FreeFile()
        Try
          FileOpen(iFileNumber, gaSourcesList(Session).TradesFile, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)

          PrintLine(iFileNumber, CStr(TradeNumber))

          FileClose(iFileNumber)
        Catch ex As Exception
          Call LogToFile("  Error: IsTradeProcessed (Source" & CStr(Session) & ") - Unable to open file (" & _
            gaSourcesList(Session).TradesFile & ") - " & ex.Message)
        End Try
      End If
    End If
  End Function

  Private Sub GetWebServiceIOSAuditTrailByOrderNo(ByVal OrderNo As Long, ByVal RootParentOrderNo As Long, Optional ByVal StartDateTime As Date = Nothing, Optional ByVal EndDateTime As Date = Nothing)
    Dim saDestination() As String = {}
    Dim lOrderNumber, lTradeNumber, lOpposingBrokerNumber, lTradeVolume As Long
    Dim sTemp, sExchange, sDestination, sTradeMarkers, sCurrency, sAccountCode, sSecurityCode As String
    Dim iIndex As Integer
    Dim lorders() As Long = {OrderNo}
    Dim dtAuditLogTime As Date = Now
    Dim OrderFlags As Long = 0
    Dim TradeDateTime As Date
    Dim sTradeNo As String
    Dim decSourcePrice, decTradePrice, DFXrate As Decimal

    Try
      If OrderNo <> RootParentOrderNo Then
        Exit Sub
      End If

      ' Retrieve information using TradeGetByUserInput method
      Dim pqgeRequest As IOSPlus.AuditTrailGetByOrderNumberInput = New IOSPlus.AuditTrailGetByOrderNumberInput

      ' Initialize the header, use the IOS Plus service session key we got earlier
      pqgeRequest.Header = New IOSPlus.AuditTrailGetByOrderNumberInputHeader
      pqgeRequest.Header.ServiceSessionKey = myIOSPlusServiceSessionKey

      Dim guid As New Guid

      ' Initialize the request options
      pqgeRequest.Header.Timeout = 25             ' Timeout after 25 seconds
      pqgeRequest.Header.PageSize = 100           ' Recommended maximum page size
      pqgeRequest.Header.Updates = False          ' Don't watch for updates
      pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
      pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

      ' Initialize the request parameters - set the parameter to retrieve information for
      pqgeRequest.Parameters = New IOSPlus.AuditTrailGetByOrderNumberInputParameters
      If StartDateTime = Nothing Then
        StartDateTime = New DateTime(gdDate.Year, gdDate.Month, gdDate.Day)

      End If
      If EndDateTime = Nothing Then
        EndDateTime = StartDateTime.AddDays(1)
      End If

      pqgeRequest.Parameters.AuditLogDateTimeFrom = StartDateTime
      pqgeRequest.Parameters.AuditLogDateTimeTo = EndDateTime


      pqgeRequest.Parameters.OrderNumberArray = lorders

      Dim pqgeResult As IOSPlus.AuditTrailGetByOrderNumberOutput

      ' Call the OrderPadGetByAccount method until we have received all pages of data
      Do
        pqgeResult = wsIOSPlus.AuditTrailGetByOrderNumber(pqgeRequest)

        For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)
          Try
            lOrderNumber = pqgeResult.Result.DataRows(iIndex).OrderNumber
            OrderFlags = pqgeResult.Result.DataRows(iIndex).OrderFlagsMask
            sTemp = pqgeResult.Result.DataRows(iIndex).EventDescription.ToString.Trim.ToUpper.Replace(" ", "")

            If (OrderFlags And 1) = 1 Then

              If sTemp.Contains("DONEFORDAY") Then

                'Call AddWSTrades(lTradeNumber, lOrderNumber, RootOrderNo, sExchange, sDestination, lTradeVolume, decTradePrice, DFXrate,
                'TradeDateTime, lOpposingBrokerNumber, sTradeMarkers, decSourcePrice, sCurrency, sAccountCode, sSecurityCode)

                dtAuditLogTime = pqgeResult.Result.DataRows(iIndex).AuditLogDateTime

                'RL003
                'If Not htDNDOrdersTable.ContainsKey(lOrderNumber) Then
                '    htDNDOrdersTable.Add(lOrderNumber, dtAuditLogTime)
                'End If

                'If Not htDNDOrdersTable.ContainsKey(RootParentOrderNo) Then
                '    htDNDOrdersTable.Add(RootParentOrderNo, dtAuditLogTime)
                'End If
              End If
              Exit Do
            End If

            'If sTemp.StartsWith("TRADENO") Then
            'ff()
            '    'RLU003
            '    sTradeNo = sTemp.Substring(0, sTemp.IndexOf(":")).Replace("TRADENO", "").Replace("(", "").Replace(")", "").Replace(":", "")
            '    Try
            '        If Not gDicDeskTradeNo.ContainsKey(sTradeNo) Then
            '            gDicDeskTradeNo.Add(sTradeNo, OrderNo)
            '        End If

            '    Catch ex As Exception

            '        Call LogToFile("  Error: GetWebServiceIOSAuditTrailByOrderNo -" & OrderNo & "-" & sTradeNo & "-" & ex.Message)
            '    End Try
            'End If


          Catch ex As Exception
            Call LogToFile("  Error: GetWebServiceIOSAuditTrailByOrderNo - " & ex.Message)
          End Try
        Next



        Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
      Loop While pqgeResult.Result.Header.StatusCode = 1


    Catch ex As Exception
      Call LogToFile("  Error: GetWebServiceIOSTradesByOrderNo - " & ex.Message)
    End Try
  End Sub


  Private Function GetWebServiceSecurity(ByVal SecurityCode() As String, ByVal Exchange() As String, ByRef CUSIP As String, ByRef ISIN As String, ByRef SEDOL As String,
                                         ByRef Description As String, ByRef UkIrishStampDutyReserveTaxMarker As String, ByRef PtmLevyIndicator As String, ByRef CurrencyDenomination As String) As Boolean
    Dim sErrorMessage As String = ""

    GetWebServiceSecurity = False

    'RL002
    Try
      ' Retrieve quote information using the PricingQuoteGet method
      Dim pqgeRequest As IRESS.PricingWatchListGetInput = New IRESS.PricingWatchListGetInput

      ' Initialize the header, use the IRESS session key we got earlier
      pqgeRequest.Header = New IRESS.PricingWatchListGetInputHeader
      pqgeRequest.Header.SessionKey = myIRESSSessionKey

      Dim guid As New Guid

      ' Initialize the request options
      pqgeRequest.Header.Timeout = glWSTimeout
      pqgeRequest.Header.PageSize = 1000          ' Recommended maximum page size
      pqgeRequest.Header.Updates = False          ' Don't watch for updates
      pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
      pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

      ' Initialize the request parameters
      pqgeRequest.Parameters = New IRESS.PricingWatchListGetInputParameters

      pqgeRequest.Parameters.SecurityCodeArray = SecurityCode
      pqgeRequest.Parameters.ExchangeArray = Exchange
      pqgeRequest.Parameters.ColumnGroupArray = {"SecurityExtra", "SecInfoEx"}

      Dim pqgeResult As IRESS.PricingWatchListGetOutput
      Dim iIndex As Integer
      ' Call the GetWebServiceQuotes method until we have received all pages of data
      Do
        pqgeResult = wsIRESS.PricingWatchListGet(pqgeRequest)

        For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)
          If iIndex = 0 Then
            If pqgeResult.Result.DataRows(0).SecurityExtra Is Nothing Or pqgeResult.Result.DataRows(0).SecurityExtra Is System.DBNull.Value Then
              Exit Do
            End If
            UkIrishStampDutyReserveTaxMarker = pqgeResult.Result.DataRows(0).SecurityExtra.UkIrishStampDutyReserveTaxMarker
            If UkIrishStampDutyReserveTaxMarker Is System.DBNull.Value Or UkIrishStampDutyReserveTaxMarker Is Nothing Then
              UkIrishStampDutyReserveTaxMarker = ""
            End If

            PtmLevyIndicator = pqgeResult.Result.DataRows(0).SecurityExtra.PtmLevyIndicator
            If PtmLevyIndicator Is System.DBNull.Value Or PtmLevyIndicator Is Nothing Then
              PtmLevyIndicator = ""
            End If
            'RL008
            CurrencyDenomination = pqgeResult.Result.DataRows(0).SecInfoEx.CurrencyDenomination
            If CurrencyDenomination Is System.DBNull.Value Or CurrencyDenomination Is Nothing Then
              CurrencyDenomination = ""
            End If


            GetWebServiceSecurity = True
          End If
        Next

        Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
      Loop While pqgeResult.Result.Header.StatusCode = 1
    Catch ex As Exception
      Call LogToFile("  Error: GetWebServiceSecurity-IRESSAPI-PricingWatchListGet - " & ex.Message)

      ' Close bad IRESS webservices
      ' Call EndIRESSSession()

      ' Start IRESS webservices
      'Call CreateIRESSSession(giIressSession)
    End Try

    Try
      ' Retrieve quote information using the PricingQuoteGet method
      Dim pqgeRequest As IRESS.SecurityInformationGetInput = New IRESS.SecurityInformationGetInput

      ' Initialize the header, use the IRESS session key we got earlier
      pqgeRequest.Header = New IRESS.SecurityInformationGetInputHeader
      pqgeRequest.Header.SessionKey = myIRESSSessionKey

      Dim guid As New Guid

      ' Initialize the request options
      pqgeRequest.Header.Timeout = glWSTimeout
      pqgeRequest.Header.PageSize = 1000          ' Recommended maximum page size
      pqgeRequest.Header.Updates = False          ' Don't watch for updates
      pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
      pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

      ' Initialize the request parameters
      pqgeRequest.Parameters = New IRESS.SecurityInformationGetInputParameters

      pqgeRequest.Parameters.SecurityCodeArray = SecurityCode
      pqgeRequest.Parameters.ExchangeArray = Exchange

      Dim pqgeResult As IRESS.SecurityInformationGetOutput
      Dim iIndex As Integer

      ' Call the GetWebServiceQuotes method until we have received all pages of data
      Do
        pqgeResult = wsIRESS.SecurityInformationGet(pqgeRequest)

        For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)
          ISIN = pqgeResult.Result.DataRows(iIndex).ISIN
          If ISIN Is System.DBNull.Value Or ISIN Is Nothing Then
            ISIN = ""
          End If
          Description = pqgeResult.Result.DataRows(iIndex).SecurityDescription
          If Description Is System.DBNull.Value Or Description Is Nothing Then
            Description = ""
          End If
          SEDOL = pqgeResult.Result.DataRows(iIndex).SEDOL
          If SEDOL Is System.DBNull.Value Or SEDOL Is Nothing Then
            SEDOL = ""
          End If
          'YearlyDividend = pqgeResult.Result.DataRows(iIndex).YearlyDividend
          'If YearlyDividend = 0 Then
          '    YearlyDividend = 0
          'End If
          GetWebServiceSecurity = True
        Next

        Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
      Loop While pqgeResult.Result.Header.StatusCode = 1
    Catch ex As Exception
      Call LogToFile("  Error: GetWebServiceSecurity-IRESSAPI-SecurityInformationGet - " & ex.Message)

      ' Close bad IRESS webservices
      Call EndIRESSSession()

      ' Start IRESS webservices
      Call CreateIRESSSession(giIressSession)
    End Try
    Try
      ' Retrieve quote information using the PricingQuoteGet method
      Dim pqgeRequest As IRESS.SecuritySearchGetInput = New IRESS.SecuritySearchGetInput

      ' Initialize the header, use the IRESS session key we got earlier
      pqgeRequest.Header = New IRESS.SecuritySearchGetInputHeader
      pqgeRequest.Header.SessionKey = myIRESSSessionKey

      Dim guid As New Guid

      ' Initialize the request options
      pqgeRequest.Header.Timeout = glWSTimeout
      pqgeRequest.Header.PageSize = 1000          ' Recommended maximum page size
      pqgeRequest.Header.Updates = False          ' Don't watch for updates
      pqgeRequest.Header.WaitForResponse = False  ' Request to IRESS asynchronously to ensure method call does not timeout
      pqgeRequest.Header.RequestID = guid.NewGuid.ToString() ' Set the request identifier to identify our request when we page through results

      ' Initialize the request parameters
      pqgeRequest.Parameters = New IRESS.SecuritySearchGetInputParameters

      pqgeRequest.Parameters.ExchangeArray = Exchange
      pqgeRequest.Parameters.SecurityCode = SecurityCode.First
      pqgeRequest.Parameters.SecurityCodeSearchType = 3

      Dim pqgeResult As IRESS.SecuritySearchGetOutput
      Dim iIndex As Integer

      ' Call the PricingQuoteGet method until we have received all pages of data
      Do
        pqgeResult = wsIRESS.SecuritySearchGet(pqgeRequest)

        For iIndex = pqgeResult.Result.DataRows.GetLowerBound(0) To pqgeResult.Result.DataRows.GetUpperBound(0)
          If ISIN = "" Then
            ISIN = pqgeResult.Result.DataRows(iIndex).ISIN
          End If
          If ISIN Is System.DBNull.Value Or ISIN Is Nothing Then
            ISIN = ""
          End If
          If Description = "" Then
            Description = pqgeResult.Result.DataRows(iIndex).SecurityDescription
          End If
          If Description Is System.DBNull.Value Or Description Is Nothing Then
            Description = ""
          End If
          If SEDOL = "" Then
            SEDOL = pqgeResult.Result.DataRows(iIndex).SEDOL
          End If
          If SEDOL Is System.DBNull.Value Or SEDOL Is Nothing Then
            SEDOL = ""
          End If
          CUSIP = pqgeResult.Result.DataRows(iIndex).CUSIP
          If CUSIP Is System.DBNull.Value Or CUSIP Is Nothing Then
            CUSIP = ""
          End If
          GetWebServiceSecurity = True
        Next

        Application.DoEvents() ' Since we're in a loop allow processing of all windows messages in message queue
      Loop While pqgeResult.Result.Header.StatusCode = 1
    Catch ex As Exception
      Call LogToFile("  Error: GetWebServiceSecurity-IRESSAPI-SecuritySearchGet  - " & ex.Message)

      ' Close bad IRESS webservices
      Call EndIRESSSession()

      ' Start IRESS webservices
      Call CreateIRESSSession(giIressSession)

    End Try




  End Function

End Class