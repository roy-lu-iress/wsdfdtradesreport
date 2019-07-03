Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain

    Private Sub ReadIniSource()
        Dim TempFileExist As Boolean
    'Dim reportType As String
        Dim tempFileName As String
        Dim dicInfo As IO.DirectoryInfo
        Dim fInfo As IO.FileInfo()
        Dim sToday As String
        Dim sTime As String()
        Dim dTempTime As Date
        Dim dLastRunTime As Date
        Dim sHour, sMin, sSec As String
        Dim iHour, iMin, iSec As Int32
        Dim sLastTempFile, sgdDate As String
        Dim iTimeZone As Int64

        ' Source date
        gdDate = ReadIniDate(gsIniFile, APP_NAME, "SourceDate", Today)



    If gdDate <> Today Then
      If MsgBox("Are you sure you want to use SourceDate of " &
        Format(gdDate, "yyyyMMdd") & "?", vbYesNo, APP_NAME) = MsgBoxResult.No Then
        Call Finish()
      End If
    Else
      ' KC002
            Dim dPreviousDayEndTime As Date

            Dim dTime As Date


      dPreviousDayEndTime = ReadIniTime(gsIniFile, APP_NAME, "PreviousDayEndTime", "01:30")
      dTime = TimeOfDay
      If dTime.TimeOfDay <= dPreviousDayEndTime.TimeOfDay Then
        gdDate = gdDate.AddDays(-1)
        gbModifiedDate = True
      End If
    End If




    ' DOS 8.3 file name format
    gbDOSFileNameFormat = ReadIniBoolean(gsIniFile, APP_NAME, "DOSFileNameFormat", False)

    ' Add computer name to WebServices Application Id
    gbAddComputerName = ReadIniBoolean(gsIniFile, APP_NAME, "AddComputerName", True)
    If gbAddComputerName Then
      gsComputerName = Environment.MachineName
    End If

    ' SQL no lock
    gbSQLNoLock = ReadIniBoolean(gsIniFile, APP_NAME, "SQLNoLock", True)

    ' SQL time out
    glSQLTimeOut = ReadIniLong(gsIniFile, APP_NAME, "SQLTimeOut", 0, 300)

    ' Webservice timeout
    glWSTimeout = ReadIniLong(gsIniFile, APP_NAME, "WebserviceTimeout", 1, 60)



    'Check Temp File Existing 
    Dim tempFolder As String
    tempFolder = ReadIniFullPath(gsIniFile, APP_NAME, "TempFolder", gdDate, False, gsAppPath & "\Temp")

    tempFileName = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-" & Now.Hour & "-" & Now.Minute & "-" & Now.Second & ".csv"

    gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, tempFolder & "\" & tempFileName)

    sToday = Now.Year & "-" & Now.Month & "-" & Now.Day & "-"
    sgdDate = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day

    'If TEMP Directory Not  Exists Then Create
    If Not IO.Directory.Exists(tempFolder) Then
      IO.Directory.CreateDirectory(tempFolder)
    End If

    gsReportingType = ReadIniString(gsIniFile, APP_NAME, "ReportType", 8, 0, False, "DELTA")

    ' KC002
    'If gdDate < Today Then
    If Not gbModifiedDate And gdDate < Today Then
      gsReportingType = "SNAPSHOT"
    End If

    '#If DEBUG Then
    '        reportType = "DELTA"
    '        tempFileName = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-12-0-0"
    '        gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, True, False, gsAppPath & "\Temp\" & tempFileName & ".csv")
    '#End If

    gbTempToFinal = True
    If gsReportingType.StartsWith("SNAP") Then
      'IF SNAPSHOT Then Report WHOLE DAY
      gdStartTime = New DateTime(gdDate.Year, gdDate.Month, gdDate.Day, 0, 0, 1)

      'gdEndTime = New DateTime(gdDate.Year, gdDate.Month, gdDate.Day + 1, 0, 0, 1)

      gdEndTime = gdStartTime.AddDays(1)

      ''gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, gsAppPath & "\Temp\" & sgdDate & ".csv")


      tempFileName = sgdDate & ".csv"

      gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, tempFolder & "\" & tempFileName)
      gbCsvHeader = True

    Else
      'IF DELTA  
      dicInfo = New DirectoryInfo(tempFolder)
      fInfo = dicInfo.GetFiles
      dLastRunTime = gdDate
      Dim bDelta As Boolean = False

      '#If DEBUG Then
      '            sToday = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-"
      '#End If

      If fInfo.Count > 0 Then
        For Each fileInfoItem As FileInfo In fInfo
          If fileInfoItem.CreationTime > Today And fileInfoItem.Name.Contains(sToday) Then
            sTime = fileInfoItem.Name.Replace(sToday, "").Replace(".csv", "").Split("-")
            sHour = sTime(0)
            sMin = sTime(1)
            sSec = sTime(2)
            Integer.TryParse(sHour, iHour)
            Integer.TryParse(sMin, iMin)
            Integer.TryParse(sSec, iSec)
            dTempTime = gdDate.AddHours(iHour).AddMinutes(iMin).AddSeconds(iSec)

            If dTempTime >= dLastRunTime Then
              dLastRunTime = dTempTime
              sLastTempFile = fileInfoItem.FullName
              bDelta = True
            End If
          Else
            File.Delete(fileInfoItem.FullName)
          End If
        Next


      End If
      gdEndTime = Now
      Dim gbDeltaMerge As Boolean = ReadIniBoolean(gsIniFile, APP_NAME, "MergeDelta", False)

      If bDelta And gbDeltaMerge Then
        'Copy Latest Delta File to new Temp output
        gdStartTime = dTempTime


        '#If DEBUG Then
        '                gdEndTime = gdDate.AddHours(20)
        '#End If



        tempFileName = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-" & gdEndTime.Hour & "-" & gdEndTime.Minute & "-" & gdEndTime.Second & ".csv"
        gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, tempFolder & "\" & tempFileName)
        File.Copy(sLastTempFile, gsTempOutputName)
      End If

      If bDelta And Not gbDeltaMerge Then

        'if Not merge previous Delta file in temp folder
        gbCsvHeader = True
        gdStartTime = dTempTime


        tempFileName = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-" & gdEndTime.Hour & "-" & gdEndTime.Minute & "-" & gdEndTime.Second & ".csv"
        gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, tempFolder & "\" & tempFileName)
      End If

      If Not bDelta Then
        'if No Delta file exist in temp folder
        gbCsvHeader = True
        gdStartTime = New DateTime(gdDate.Year, gdDate.Month, gdDate.Day, 0, 0, 1)

        '#If DEBUG Then
        '                gdEndTime = gdDate.AddHours(12)
        '#End If


        tempFileName = gdDate.Year & "-" & gdDate.Month & "-" & gdDate.Day & "-" & gdEndTime.Hour & "-" & gdEndTime.Minute & "-" & gdEndTime.Second & ".csv"
        gsTempOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "Temp", gdDate, False, tempFolder & "\" & tempFileName)
      End If


    End If


    If gbTempToFinal And File.Exists(gsOutputName) Then
      File.Delete(gsOutputName)
    End If


    If gbCsvHeader And IO.File.Exists(gsTempOutputName) Then
      File.Delete(gsTempOutputName)
    End If


    gbRemoveTemp = ReadIniBoolean(gsIniFile, APP_NAME, "RemoveTemp", False)

    gbDFD = ReadIniBoolean(gsIniFile, APP_NAME, "DoneForDay", True)

    gbTradedOrderOnly = ReadIniBoolean(gsIniFile, APP_NAME, "TradedOrderOnly", True)

    ' Source Details
    Call ReadSourceDetails()
    Dim sysTime As DateTime
    Dim sysUTCTime As DateTime
    Dim timeZone As Long
    sysTime = DateTime.Now
    sysUTCTime = sysTime.ToUniversalTime

    timeZone = sysTime.Hour - sysUTCTime.Hour
    'If bdelta Then

    'End If
    'giReportServerTimeZone = ReadIniLong(gsIniFile, APP_NAME, "ReportServerTimeZone", -1, timeZone)
    'giTradeServerTimeZone = ReadIniLong(gsIniFile, APP_NAME, "TradeServerTimeZone", -1, 1)

    'If giReportServerTimeZone <> giTradeServerTimeZone Then
    '    gdStartTime = gdStartTime.AddHours(giTradeServerTimeZone - giReportServerTimeZone)
    '    gdEndTime = gdEndTime.AddHours(giTradeServerTimeZone - giReportServerTimeZone)
    'End If

    End Sub

    Private Sub ReadSourceDetails()
        Dim bReadSQLSourceConfig As Boolean
        Dim iIndex As Integer
        Dim sLogMsg As String

        '' Broker number
        giBrokerNumber = ReadIniLong(gsIniFile, APP_NAME, "BrokerNumber", -1, -1)
        'If giBrokerNumber < 0 Then
        '    Call LogToFile("  Error: ReadSourceDetails - BrokerNumber (" & CStr(giBrokerNumber) & ") is invalid")
        '    Call Finish()
        'End If

        If gbReadConfigFromINI Then
            Call LogToFile("  Info: ReadSourceDetails - Reading Source Config from Ini")
            Call ReadIniSourceConfig()
        Else
            Call LogToFile("  Info: ReadSourceDetails - Reading Source Config from SQL")
            bReadSQLSourceConfig = ReadSQLSourceConfig()
            If Not bReadSQLSourceConfig Then
                Call LogToFile("  Info: ReadSourceDetails - Reading Source Config from Ini")
                Call ReadIniSourceConfig()
            End If
        End If

        ' "Log source"
        For iIndex = 0 To (MAX_SOURCES - 1)
            If gaSourcesList(iIndex).DatabaseType = "WS" Then
                sLogMsg = "  Info: ReadIniSourceDetails - Source" & CStr(iIndex) & " URL: " & gaSourcesList(iIndex).WSUrl & _
                  ", UserName: " & gaSourcesList(iIndex).WSUserName & ", CompanyName: " & gaSourcesList(iIndex).WSCompanyName & _
                  ", Password: " & gaSourcesList(iIndex).WSPassword
                If gaSourcesList(iIndex).WSIOS <> "" Then
                    sLogMsg = sLogMsg & ", IOS: " & gaSourcesList(iIndex).WSIOS
                End If
                sLogMsg = sLogMsg & ", ApplicationId: " & gaSourcesList(iIndex).WSApplicationId

                Call LogToFile(sLogMsg)
            Else
                Exit For
            End If
        Next

        If bReadSQLSourceConfig Then
            Call SaveSource()
        End If
    End Sub

    Private Sub ReadIniOutput()
        Dim sReportType(1) As String
        Dim bDefault(1) As Boolean
        Dim iNum As Integer
        Dim TempFileName As String
        Dim bArchive As Boolean

        sReportType(0) = "xls"
        sReportType(1) = "csv"

        bDefault(0) = True
        bDefault(1) = False




        gsOutputName = ReadIniFullPath(gsIniFile, APP_NAME, "ReportName", gdDate, False, gsAppPath & "\" & APP_NAME & "-" & Format(gdDate, "yyyyMMdd") & ".csv")

        bArchive = ReadIniBoolean(gsIniFile, APP_NAME, "Archive", False)

        Try
            If File.Exists(gsOutputName) Then
                If bArchive Then
                    iNum = 1
                    TempFileName = gsOutputName.Replace(".csv", "-1.csv")

                    Do
                        If File.Exists(TempFileName) Then
                        Else
                            Exit Do
                        End If

                        TempFileName = TempFileName.Replace(iNum.ToString & ".csv", (iNum + 1).ToString & ".csv")

                        iNum = iNum + 1
                    Loop
                    File.Copy(gsOutputName, TempFileName)

                End If
                File.Delete(gsOutputName)
            End If
        Catch ex As Exception
            Call LogToFile("  Error: ReadIniOutput -  " & ex.Message)

            gsOutputName = gsOutputName.Replace(Format(gdDate, "yyyyMMdd"), Format(gdDate, "yyyyMMdd") & "-" & Format(TimeOfDay, "hhmmss"))
            Call LogToFile("  Info: New Output -  " & gsOutputName)
        End Try


        gsCsvDelimiter = ReadIniString(gsIniFile, APP_NAME, "csvDelimiter", 1, 0, False, ",")


    End Sub

  Private Sub ReadIniOptions()

    giOrderFilter = ReadIniLong(gsIniFile, APP_NAME, "OrderFilter", 1, 7)

    Dim Muliplier As Decimal
    Dim Currency As String
    Dim Exchanges As String
    'RL009
    For i As Int32 = 0 To 100
      Exchanges = ReadIniString(gsIniFile, "CurrencyOverride" & CStr(i), "Exchanges", 0, 1, False, "*END")
      Muliplier = ReadIniDecimal(gsIniFile, "CurrencyOverride" & CStr(i), "Multiplier", 3, 1)
      Currency = ReadIniString(gsIniFile, "CurrencyOverride" & CStr(i), "Currency", 0, 1, False, "")
      For Each exchange As String In Exchanges.Split(",")
        If Not htExchangeCurrencys.ContainsKey(exchange) Then
          htExchangeCurrencys.Add(exchange, Currency)
        End If
        If Not htExchangeMultipliers.ContainsKey(exchange) Then
          htExchangeMultipliers.Add(exchange, Muliplier)
        End If
      Next
      If Exchanges = "*ALL" Or Exchanges = "*END" Then
        Exit For
      End If
    Next


    '        Order Filter Codes	
    '1 - Active orders with remaining volume	
    '2 - Active orders with remaining volume And/Or any orders traded today	
    '3 - All orders for the last 3 days	
    '4 - Active orders with remaining volume And/Or any parent orders traded today	
    '5 - All active orders	
    '6 - All orders for today And yesterday	
    '7 - All orders for today	
    '8 - Purged at-risk orders for today And previous trading day 	

    '        If Now.DayOfYear = gdDate.DayOfYear Then
    '            giOrderFilter = 5
    '#If DEBUG Then
    '            giOrderFilter = 5
    '            'giOrderFilter = 3
    '#End If

    '        ElseIf Now.DayOfYear - gdDate.DayOfYear < 3 Then
    '            giOrderFilter = 3
    '        Else
    '            Call LogToFile("  Error: ReadIniOptions - Webservice cannot read order older than 3 days")
    '        End If

    gbNetting = ReadIniBoolean(gsIniFile, APP_NAME, "Netting", False)

    ' KC003
    gsTplus3Exchanges = ReadIniString(gsIniFile, APP_NAME, "T+3Exchanges", 0, 1, False, "")

    gsSelttlementPath = ReadIniFullPath(gsIniFile, APP_NAME, "SettlementHolidaysPath", gdDate, False, gsAppPath & "\SettlementHolidays")
    gsExchangeToCountryPath = ReadIniFullPath(gsIniFile, APP_NAME, "ExchangeToCountryPath", gdDate, False, gsAppPath & "\Exchange2Country.txt")
    Call ReadSettlementHolidays()
    Call ReadCountries()

    ' KC006
    gsSettleTradeDateFormat = ReadIniString(gsIniFile, APP_NAME, "SettleTradeDateFormat", 0, 0, False, "dd/MM/yyyy")

    gbGroupSort = ReadIniBoolean(gsIniFile, APP_NAME, "GroupSort", False)

    gsSortOrder = ReadIniString(gsIniFile, APP_NAME, "SortOrder", 0, 0, False, "OrdNo")
  End Sub

    Private Sub ReadIniReportAccounts()
        Dim sTemp As String
        Dim iIndex As Integer = 0
        Dim iNum As Integer = 0

        Call ReadReportAccountsDefaults()

        Do
            ' Account Code
            sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(iIndex), REPORT_ID_KEY, 0, 1, False, "")
            If sTemp = "*END" Then
                Exit Do
            End If

            If sTemp <> "" Then
                If iNum < MAX_REPORT_ACCOUNTS Then
                    Try
                        htReportAccountsTable.Add(sTemp, iNum)

                        gaReportAccountsList(iNum).AccCode = sTemp
                        Call ReadReportAccountsDetails(iIndex, iNum)

                        iNum = iNum + 1
                    Catch ex As Exception
                        Call LogToFile("  Error: ReadIniReportAccounts (" & REPORT_USER_SECTION & CStr(iIndex) & _
                          ") - Unable to add to table (" & sTemp & ") - " & ex.Message)
                    End Try
                Else
                    Call LogToFile("  Error: ReadIniReportAccounts (" & REPORT_USER_SECTION & CStr(iIndex) & ") - Table full (" & _
                      sTemp & ")")
                End If
            End If

            If iIndex < (MAX_ENTRIES - 1) Then
                iIndex = iIndex + 1
            Else
                Call LogToFile("  Error: ReadIniReportAccounts - Maximum entries reached (" & CStr(MAX_ENTRIES) & ")")
                Exit Do
            End If
        Loop

        If iNum = 0 Then
            Call LogToFile("  Error: ReadIniReportAccounts - No valid " & REPORT_ID_KEY)
            Call Finish()
        End If
    End Sub

    Private Sub ReadReportAccountsDefaults()
        Dim sTemp As String

        ' Destinations
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_DESTINATIONS, 0, 1, False, "*ALL")
        gsDefaultReportDestinations = ConvertDestinations(sTemp)
        ' Exception destinations
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_EXCEPTION_DESTINATIONS, 0, 1, False, "*NONE")
        gsDefaultReportExceptionDestinations = ConvertDestinations(sTemp)

        ' Exchanges
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_EXCHANGES, 0, 1, False, "*ALL")
        gsDefaultReportExchanges = ConvertExchanges(sTemp)
        ' Exception exchanges
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_EXCEPTION_EXCHANGES, 0, 1, False, "*NONE")
        gsDefaultReportExceptionExchanges = ConvertExchanges(sTemp)

        ' Currency
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_CURRENCY, 0, 1, False, "*ALL")
        gsDefaultReportCurrency = ConvertCurrency(sTemp, "*ALL")
        ' Account types
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_ACCOUNT_TYPES, 0, 1, False, "*ALL")
        gsDefaultReportAccountTypes = ReplaceString(sTemp, "RT", "ST")

        ' Account ids
        gsDefaultReportAccountIds = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_ACCOUNT_IDS, 0, 1, False, "*ALL")
        ' Exception account ids
        gsDefaultReportExceptionAccountIds = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_EXCEPTION_ACCOUNT_IDS, _
          0, 1, False, "*NONE")

        ' Symbols
        gsDefaultReportSymbols = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_SYMBOLS, 0, 1, False, "*ALL")

        ' Exception symbols
        gsDefaultReportExceptionSymbols = ReadIniString(gsIniFile, REPORT_USER_SECTION, REPORT_KEY_EXCEPTION_SYMBOLS, 0, 1,
          False, "*NONE")


        ' ExceptionOrderTypes
        gsDefaultReportExceptionOrderTypes = ReadIniString(gsIniFile, REPORT_USER_SECTION, "ExceptionOrderTypes", 0, 1,
          False, "*NONE").ToString.Trim.ToUpper.Replace(" ", "")

        If gsDefaultReportExceptionOrderTypes <> "*NONE" Then
            gsExceptionOrderTypes = gsDefaultReportExceptionOrderTypes.Split(",")
            For Each item As String In gsExceptionOrderTypes
                htExceptionOrderTypes.Add(item, "")
            Next
        End If

        ' OrderTypes
        gsDefaultReportOrderTypes = ReadIniString(gsIniFile, REPORT_USER_SECTION, "OrderTypes", 0, 1, False, "*ALL").ToString.Trim.ToUpper.Replace(" ", "")

        If gsDefaultReportOrderTypes <> "*ALL" Then
            gsOrderTypes = gsDefaultReportOrderTypes.Split(",")
            For Each item As String In gsOrderTypes
                htOrderTypes.Add(item, "")
            Next
        End If


    End Sub
	
    Private Sub ReadReportAccountsDetails(ByVal Index As Integer, ByVal Num As Integer)
        Dim sTemp As String

        ' Destinations
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_DESTINATIONS, 0, 1, False, _
          gsDefaultReportDestinations)
        gaReportAccountsList(Num).Destinations = ConvertDestinations(sTemp)
        ' Exception destinations
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_EXCEPTION_DESTINATIONS, 0, 1, False, _
          gsDefaultReportExceptionDestinations)
        gaReportAccountsList(Num).ExceptionDestinations = ConvertDestinations(sTemp)

        ' Exchanges
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_EXCHANGES, 0, 1, False, _
          gsDefaultReportExchanges)
        gaReportAccountsList(Num).Exchanges = ConvertExchanges(sTemp)
        ' Exception exchanges
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_EXCEPTION_EXCHANGES, 0, 1, False, _
          gsDefaultReportExceptionExchanges)
        gaReportAccountsList(Num).ExceptionExchanges = ConvertExchanges(sTemp)

        ' Currency
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_CURRENCY, 0, 1, False, _
          gsDefaultReportCurrency)
        gaReportAccountsList(Num).Currency = ConvertCurrency(sTemp, "*ALL")
        ' Account types
        sTemp = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_ACCOUNT_TYPES, 0, 1, False, _
          gsDefaultReportAccountTypes)
        gaReportAccountsList(Num).AccountTypes = ReplaceString(sTemp, "RT", "ST")

        ' Account ids
        gaReportAccountsList(Num).AccountIds = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index),
          REPORT_KEY_ACCOUNT_IDS, 0, 1, False, gsDefaultReportAccountIds)

        ' Exception account ids
        gaReportAccountsList(Num).ExceptionAccountIds = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index),
          REPORT_KEY_EXCEPTION_ACCOUNT_IDS, 0, 1, False, gsDefaultReportExceptionAccountIds)

        ' Symbols
        gaReportAccountsList(Num).Symbols = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), REPORT_KEY_SYMBOLS, _
          0, 1, False, gsDefaultReportSymbols)
        ' Exception symbols
        gaReportAccountsList(Num).ExceptionSymbols = ReadIniString(gsIniFile, REPORT_USER_SECTION & CStr(Index), _
          REPORT_KEY_EXCEPTION_SYMBOLS, 0, 1, False, gsDefaultReportExceptionSymbols)
    End Sub


    Private Sub ReadSettlementHolidays()
        Dim di As DirectoryInfo
        Dim fileName As String
        Dim skey As String
        htHolidays = New Hashtable
        If IO.Directory.Exists(gsSelttlementPath) Then
            di = New DirectoryInfo(gsSelttlementPath)
            If di.GetFiles.Count = 0 Then
                Call LogToFile("  Error: SelttlementPath does not  have any files")
            End If
            For Each fi In di.GetFiles
                ' Open the stream and read it back.
                Dim sr As StreamReader = fi.OpenText()

                Dim dt As String = ""
                fileName = fi.Name.Substring(0, 4)

                While sr.EndOfStream = False
                    dt = sr.ReadLine()
                    skey = fileName & dt
                    htHolidays.Add(skey, 0)
                End While
                sr.Close()

            Next
        Else
            Call LogToFile("  Error: SelttlementPath does not existed")
        End If
    End Sub

    Private Sub ReadIniOrders()
        Dim iIndex As Long = 0
        Dim iOrderNo As Long = 0
        Dim dDate As Date
        Do
            iOrderNo = ReadIniLong(gsIniFile, "Orders", "OrderNo" & iIndex, 0, 0)
            If iOrderNo = 0 Then
                Exit Do
            End If
            dDate = ReadIniDate(gsIniFile, "Orders", "Date" & iIndex, Now)

            htReportOrdersTable.Add(iOrderNo, dDate)
            iIndex = iIndex + 1
        Loop

    End Sub
    Private Sub WriteIniOrders()
        Dim iIndex As Long = 0
        Dim iOrderNo As Long = 0
        Dim dDate As Date
        Do

        Loop

        For Each key As Long In htReportOrdersTable.Keys
            iOrderNo = key

            If iOrderNo = 0 Then
                Exit For
            End If

            dDate = htReportOrdersTable.Item(iOrderNo)

            WriteINI(gsIniFile, "Orders", "OrderNo" & iIndex, iOrderNo)
            WriteINI(gsIniFile, "Orders", "Date" & iIndex, dDate)

            iIndex = iIndex + 1
        Next
    End Sub

    Private Sub CleanIniOrders()
        Dim iIndex As Long = 0
        Dim iOrderNo As Long = 0

        Do
            iOrderNo = ReadIniLong(gsIniFile, "Orders", "OrderNo" & iIndex, 0, 0)

            If iOrderNo = 0 Then
                Exit Do
            End If

            WriteINI(gsIniFile, "Orders", "OrderNo" & iIndex, 0)
            WriteINI(gsIniFile, "Orders", "Date" & iIndex, 0)

            iIndex = iIndex + 1
        Loop

    End Sub

    Dim dicExchangeDictionary As Dictionary(Of String, String)
  Private Sub ReadCountries()

    Dim fileName As String
    Dim skey As String
    Dim sValue As String

    dicExchangeDictionary = New Dictionary(Of String, String)

    If IO.File.Exists(gsExchangeToCountryPath) Then
      Dim fi As FileInfo
      fi = New FileInfo(gsExchangeToCountryPath)

      ' Open the stream and read it back.
      Dim sr As StreamReader = fi.OpenText()

      Dim dt As String = ""
      fileName = fi.Name.Substring(0, 4)

      While sr.EndOfStream = False
        dt = sr.ReadLine().Trim
        If dt.Length < 2 Then
          Continue While
        End If

        ' KC005
        'skey = dt.Split(",")(0)
        'sValue = dt.Split(",")(1)
        skey = dt.Split(",")(0).Trim
        sValue = dt.Split(",")(1).Trim

        dicExchangeDictionary.Add(skey, sValue)
      End While
      sr.Close()
    Else
      Call LogToFile("  Error: ExchangeToCountryPath does not existed")
    End If
  End Sub

End Class