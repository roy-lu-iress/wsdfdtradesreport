Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain

    Private Function ConvertDestinations(ByVal Destinations As String) As String
        If Destinations <> "*ALL" And Destinations <> "*NONE" Then
            Destinations = ReplaceString(Destinations, "TSE", "TSX")
            Destinations = ReplaceString(Destinations, "MN", "TCM")
            Destinations = ReplaceString(Destinations, "PUR", "PURE")
            Destinations = ReplaceString(Destinations, "CHI", "CHIX")
            Destinations = ReplaceString(Destinations, "ALF", "ALPHA")
            Destinations = ReplaceString(Destinations, "OMG", "OMEGA")
            Destinations = ReplaceString(Destinations, "SEL", "TMX-SELECT")
            Destinations = ReplaceString(Destinations, "CX2", "CXX")
        End If

        ConvertDestinations = Destinations
    End Function

    Private Function ConvertExchanges(ByVal Exchanges As String) As String

        If Exchanges <> "*ALL" And Exchanges <> "*NONE" Then
            Exchanges = ReplaceString(Exchanges, "TSE", "TSX")
            Exchanges = ReplaceString(Exchanges, "CDNX", "TSXV")
            Exchanges = ReplaceString(Exchanges, "TCM", "MN")
            Exchanges = ReplaceString(Exchanges, "PURE", "PUR")
            Exchanges = ReplaceString(Exchanges, "CHIX", "CHI")
            Exchanges = ReplaceString(Exchanges, "ALPHA", "ALF")
            Exchanges = ReplaceString(Exchanges, "OMEGA", "OMG")
            Exchanges = ReplaceString(Exchanges, "TMX-SELECT", "SEL")
            Exchanges = ReplaceString(Exchanges, "CX2", "CXX")
        End If

        ConvertExchanges = Exchanges
    End Function

    Private Function ConvertCurrency(ByVal Currency As String, ByVal DefaultCurrency As String) As String
        Dim iLen As Integer

        If Currency <> DefaultCurrency Then
            iLen = Len(Currency)
            If Currency = Microsoft.VisualBasic.Left("CAD", iLen) Then
                ConvertCurrency = "CAD"
            ElseIf Currency = Microsoft.VisualBasic.Left("USD", iLen) Then
                ConvertCurrency = "USD"
            Else
                Call LogToFile("  Error: ConvertCurrency - Currency (" & Currency & ") is invalid. Using default (" & _
                  DefaultCurrency & ")")
                ConvertCurrency = DefaultCurrency
            End If
        Else
            ConvertCurrency = DefaultCurrency
        End If
    End Function

    Private Sub ConvertDestinationListToArray(ByVal DestinationList As String, ByRef DestinationArray() As String)
        Dim strSplit As String()
        Dim iIndex As Integer

        strSplit = DestinationList.Split(",")

        ReDim DestinationArray(strSplit.GetUpperBound(0))
        For iIndex = strSplit.GetLowerBound(0) To strSplit.GetUpperBound(0)
            DestinationArray(iIndex) = strSplit(iIndex).Trim
        Next
    End Sub

    Private Function GetAccountType(ByVal ExecutionInstructions As String) As String
        GetAccountType = UCase(Trim(GetTaggedFieldFromTMXI("", "C", "Account Type", ExecutionInstructions)))

        'If GetAccountType = "" Then
        '    GetAccountType = "NC"
        'ElseIf GetAccountType = "RT" Then
        '    GetAccountType = "ST"
        'End If

        If GetAccountType = "CL" Then
            GetAccountType = "Client"
        ElseIf GetAccountType = "" Then
            GetAccountType = "Market"
        ElseIf GetAccountType = "RT" Then
            GetAccountType = "ST"
        End If
    End Function

    Private Function MatchReportAccount(ByVal AccCode As String, ByVal Destination As String, ByVal Exchange As String, _
      ByVal Currency As String, ByVal AccountType As String, ByVal AccountId As String, ByVal Symbol As String) As Boolean
        Dim iIndex As Integer

        MatchReportAccount = False

        If htReportAccountsTable.ContainsKey(AccCode) Or htReportAccountsTable.ContainsKey("*ALL") Then
            iIndex = htReportAccountsTable.Item(AccCode)
            If MatchAccount(iIndex, Destination, Exchange, Currency, AccountType, AccountId, Symbol) Then
                MatchReportAccount = True
                Exit Function
            End If

        End If
    End Function

    Private Function MatchAccount(ByVal Index As Integer, ByVal Destination As String, ByVal Exchange As String, _
      ByVal Currency As String, ByVal AccountType As String, ByVal AccountId As String, ByVal Symbol As String) As Boolean

        MatchAccount = False

        If Not MatchType(gaReportAccountsList(Index).Destinations, Destination) Then
            Exit Function
        End If

        If MatchType(gaReportAccountsList(Index).ExceptionDestinations, Destination) Then
            Exit Function
        End If

        If Not MatchExchange(Exchange, gaReportAccountsList(Index).Exchanges) Then
            Exit Function
        End If

        If MatchExchange(Exchange, gaReportAccountsList(Index).ExceptionExchanges) Then
            Exit Function
        End If

        If Not MatchType(gaReportAccountsList(Index).Currency, Currency) Then
            Exit Function
        End If

        If Not MatchType(gaReportAccountsList(Index).AccountTypes, AccountType) Then
            Exit Function
        End If

        If Not MatchType(gaReportAccountsList(Index).AccountIds, AccountId) Then
            Exit Function
        End If

        If MatchType(gaReportAccountsList(Index).ExceptionAccountIds, AccountId) Then
            Exit Function
        End If

        If Not MatchType(gaReportAccountsList(Index).Symbols, Symbol) Then
            Exit Function
        End If

        If MatchType(gaReportAccountsList(Index).ExceptionSymbols, Symbol) Then
            Exit Function
        End If

        MatchAccount = True
    End Function

    Private Function MatchType(ByVal SearchStr As String, ByVal FindStr As String) As Boolean
        Dim sToken As String

        MatchType = False

        Do
            sToken = Trim(StrTok(",", SearchStr))
            If sToken = "" Then
                Exit Function
            ElseIf ((sToken = "*ALL") Or (sToken = FindStr) Or ((sToken = "*BLANK") And (FindStr = ""))) Then
                MatchType = True
                Exit Function
            End If
        Loop
    End Function

    Private Function MatchExchange(ByVal Exchange As String, ByVal SearchStr As String) As Boolean
        Dim sToken As String

        MatchExchange = False

        If Exchange <> "" Then
            Do
                sToken = Trim(StrTok(",", SearchStr))
                If sToken = "" Then
                    Exit Function
                ElseIf (sToken = "*ALL") Then
                    MatchExchange = True
                    Exit Function
                ElseIf (sToken = Exchange) Then
                    MatchExchange = True
                    Exit Function
                ElseIf (sToken = "*CDN") And IsCDNExchange(Exchange) Then
                    MatchExchange = True
                    Exit Function
                ElseIf (sToken = "*US") And (Not IsCDNExchange(Exchange)) Then
                    MatchExchange = True
                    Exit Function
                End If
            Loop
        End If
    End Function

    Private Function IsCDNExchange(ByVal Exchange As String) As Boolean
        If IsCDNInStkStat(Exchange) Or Exchange = "CNQ" Or Exchange = "MX" Or Exchange = "OMG" Or Exchange = "LYNX" Then
            IsCDNExchange = True
        Else
            IsCDNExchange = False
        End If
    End Function

    Private Function IsCDNInStkStat(ByVal Exchange As String) As Boolean
        If Exchange = "TSX" Or Exchange = "TSXV" Or Exchange = "MN" Or Exchange = "PUR" Or Exchange = "CHI" Or _
          Exchange = "ALF" Or Exchange = "SEL" Or Exchange = "ICX" Or Exchange = "CXX" Then
            IsCDNInStkStat = True
        Else
            IsCDNInStkStat = False
        End If
    End Function

    Private Function MapPostTradeStatus(ByVal PostTradeStatusNumber As Integer) As String
        Select Case PostTradeStatusNumber
            Case 1
                MapPostTradeStatus = ""
            Case 2
                MapPostTradeStatus = "Open"
            Case 3
                MapPostTradeStatus = "Partial Fill"
            Case 4
                MapPostTradeStatus = "Filled"
            Case 5
                MapPostTradeStatus = "Held"
            Case 6
                MapPostTradeStatus = "NOE Sent"
            Case 7
                MapPostTradeStatus = "Matched"
            Case 8
                MapPostTradeStatus = "Ticketed"
            Case 9
                MapPostTradeStatus = "Confirmed"
            Case Else
                Call LogToFile("  Error: MapPostTradeStatus - PostTradeStatusNumber (" & CStr(PostTradeStatusNumber) & _
                  ") is invalid")
                MapPostTradeStatus = ""
        End Select
    End Function

    Private Function CheckHoliday(ByVal exchange As String, ByVal SettlementDate As Date) As Boolean
        Dim sDate As String
        Dim sKey As String
        Dim sCountry As String
        CheckHoliday = False

        sDate = SettlementDate.Month.ToString & "/" & SettlementDate.Day.ToString & "/" & SettlementDate.Year.ToString


        If dicExchangeDictionary.ContainsKey(exchange) Then
            sCountry = dicExchangeDictionary.Item(exchange)
            sKey = sCountry & "." & sDate
            If htHolidays.ContainsKey(sKey) Then
                Return True
            End If
        End If

        sKey = exchange & "." & sDate
        If htHolidays.ContainsKey(sKey) Then
            Return True
        End If

        Return False
    End Function



    Private Function IsWorkDays(ByVal SettlementDate As Date) As Boolean
        IsWorkDays = True
        If SettlementDate.DayOfWeek = DayOfWeek.Saturday Or SettlementDate.DayOfWeek = DayOfWeek.Sunday Then
            IsWorkDays = False
        End If
    End Function


    Private Function GetSettlementDate(ByVal exchange As String, ByVal TradeDate As Date, ByVal SettlementDays As Int32) As Date
        Dim dtSettlement As Date
        Dim bWorkDays As Boolean
        Dim bHolidays As Boolean
        dtSettlement = TradeDate

        While SettlementDays <> 0
            dtSettlement = dtSettlement.AddDays(1)
            bWorkDays = IsWorkDays(dtSettlement)
            bHolidays = CheckHoliday(exchange, dtSettlement)
            If Not bWorkDays Then
                SettlementDays = SettlementDays
            ElseIf bHolidays Then
                SettlementDays = SettlementDays
            Else
                SettlementDays = SettlementDays - 1
            End If
        End While

        Return dtSettlement
    End Function
End Class