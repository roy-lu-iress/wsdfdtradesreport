Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain


    Private Sub BackupLog()
        If (Weekday(gdDate) <> vbMonday) Then
            Exit Sub
        End If

        ' Check for existing log file
        If Not File.Exists(gsLogFile) Then
            Exit Sub
        End If

        ' Check log file update time stamp
        If FileModifiedToday(gsLogFile) Then
            Exit Sub
        End If

        If DeleteFile(gsBackupLog) Then
            ' Rename log file
            Try
                File.Move(gsLogFile, gsBackupLog)
            Catch ex As Exception
                Call LogToFile("  Error: BackupLog - Unable to rename file (" & gsLogFile & ") - " & ex.Message)
            End Try
        End If
    End Sub

    Private Function FileModifiedToday(ByVal FileName As String) As Boolean
        Dim objFileInfo As New FileInfo(FileName)
        Dim dtLastWriteTime As DateTime = objFileInfo.LastWriteTime

        FileModifiedToday = False

        If dtLastWriteTime.Date = gdDate.Date Then
            FileModifiedToday = True
        End If
    End Function

    Private Function DeleteFile(ByVal Filename As String) As Boolean
        DeleteFile = True

        If File.Exists(Filename) Then
            Try
                File.Delete(Filename)
            Catch ex As Exception
                Call LogToFile("  Error: DeleteFile - Unable to delete file (" & Filename & ") - " & ex.Message)
                DeleteFile = False
            End Try
        End If
    End Function

    Private Sub LogToFile(ByVal LogMessage As String)
        Dim iFileNumber As Integer
        Dim dTimeStamp As Date

        iFileNumber = FreeFile()
        Try
            FileOpen(iFileNumber, gsLogFile, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)
        Catch ex As Exception
            MsgBox("Error: LogToFile - Unable to open file (" & gsLogFile & ") - " & ex.Message, , APP_NAME)
            Exit Sub
        End Try

        dTimeStamp = Now
        PrintLine(iFileNumber, Format(dTimeStamp, "MM/dd/yy HH:mm:ss") & Space(1) & LogMessage)
        FileClose(iFileNumber)
    End Sub

    Private Sub ParseCommandLine()
        Dim cmdLine As String()
        Dim iNum As Integer
        Dim iUpper As Integer

        cmdLine = GetCommandLineArgs()
        iNum = cmdLine.GetLowerBound(0)
        iUpper = cmdLine.GetUpperBound(0)

        While iNum <= iUpper
            Select Case cmdLine(iNum).ToLower
                Case "-r"   ' To specify instance
                    If iNum <= (iUpper - 1) Then
                        If Not cmdLine(iNum + 1).StartsWith("-") Then
                            gsInstanceName = cmdLine(iNum + 1)
                            iNum = iNum + 1
                        End If
                    End If
            End Select

            iNum = iNum + 1
        End While
    End Sub

    Private Function GetCommandLineArgs() As String()
        Dim sCommands As String = Microsoft.VisualBasic.Command()
        Dim sSeparators As String = " "

        GetCommandLineArgs = sCommands.Split(sSeparators.ToCharArray)
    End Function

    Private Sub ReadRegistry()
        Dim sBaseKey As String
        Dim sSQLServer As String
        Dim sError As String = ""
        Dim sSQLUserName As String
        Dim sSQLPassword As String
        Dim sSQLDatabase As String
        Dim sPrimary As String

        ' Read base key from INI
        sBaseKey = ReadIniString(gsIniFile, APP_NAME, "RegistryBaseKey", 0, 0, False, "SOFTWARE\DFS\EOD")

        sSQLServer = RegValue(RegistryHive.CurrentUser, sBaseKey, "SQLServer", sError)
        If sSQLServer = "" Then
            gbReadConfigFromINI = True
            Exit Sub
        End If

        sSQLUserName = RegValue(RegistryHive.CurrentUser, sBaseKey, "SQLUserName", sError)
        If sSQLUserName = "" Then
            gbReadConfigFromINI = True
            Exit Sub
        End If

        sSQLPassword = RegValue(RegistryHive.CurrentUser, sBaseKey, "SQLPassword", sError)

        sSQLDatabase = RegValue(RegistryHive.CurrentUser, sBaseKey, "SQLDatabase", sError)
        If sSQLDatabase = "" Then
            gbReadConfigFromINI = True
            Exit Sub
        End If

        gsConfigSQLConnectStr = CreateSQLConnStr(sSQLUserName, sSQLPassword, sSQLServer, sSQLDatabase)

        gbPrimary = True
        sPrimary = RegValue(RegistryHive.CurrentUser, sBaseKey, "PrimaryDataSources", sError)
        If sPrimary.ToUpper = "N" Then
            gbPrimary = False
        End If
    End Sub

    Private Function ReadIniString(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String, _
      ByVal Truncate As Integer, ByVal StringCase As Integer, ByVal CriticalError As Boolean, _
      ByVal DefaultString As String) As String
        Dim sTemp As String

        ReadIniString = ""

        sTemp = Trim(ReadINI(IniFile, IniSection, IniKey, ""))
        If sTemp <> "" Then
            If UCase(sTemp) = "*BLANK*" Then
                Exit Function
            End If

            If Truncate > 0 Then
                If Len(sTemp) > Truncate Then
                    Call LogToFile("  Error: ReadIniString - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                      ") will be truncated")
                    ReadIniString = Microsoft.VisualBasic.Left(sTemp, Truncate)
                Else
                    ReadIniString = sTemp
                End If
            Else
                ReadIniString = sTemp
            End If

            Select Case StringCase
                Case 0
                    ' Leave as is
                Case 1
                    ReadIniString = UCase(ReadIniString)
                Case 2
                    ReadIniString = LCase(ReadIniString)
                Case Else
                    Call LogToFile("  Error: ReadIniString - StringCase (" & CStr(StringCase) & ") is invalid")
            End Select
        Else
            If Not CriticalError Then
                ReadIniString = DefaultString
            Else
                Call LogToFile("  Error: ReadIniString - [" & IniSection & "], " & IniKey & " is missing")
                Call Finish()
            End If
        End If
    End Function


    Private Function ReadIniStringArray(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String,
      ByVal Truncate As Integer, ByVal StringCase As Integer, ByVal CriticalError As Boolean,
      ByVal DefaultString As String, ByVal SplitChar As String) As String()
        Dim sTemp As String



        sTemp = ReadIniString(IniFile, IniSection, IniKey, Truncate, StringCase, CriticalError, DefaultString).Trim.ToUpper

        ReadIniStringArray = sTemp.Split(SplitChar)
    End Function

    Private Function ReadIniDate(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String,
    ByVal DefaultDate As Date) As Date
        Dim sTemp As String

        sTemp = Trim(ReadINI(IniFile, IniSection, IniKey, ""))
        If sTemp <> "" Then
            Try
                If sTemp.Substring(1, 1) = "/" Then
                    sTemp = sTemp.Insert(0, "0")
                End If
                If sTemp.Substring(4, 1) = "/" Then
                    sTemp = sTemp.Insert(3, "0")
                End If
                ReadIniDate = Date.ParseExact(sTemp, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo,
        Globalization.DateTimeStyles.None)
            Catch ex As Exception
                Call LogToFile("  Error: ReadIniDate - [" & IniSection & "], " & IniKey & " (" & sTemp &
        ") cannot be converted. Using default (" & Format(DefaultDate, "MM/dd/yyyy") & ") - " & ex.Message)
                ReadIniDate = DefaultDate
            End Try

            '  If IsDate(sTemp) Then
            '  Else
            '      Call LogToFile("  Error: ReadIniDate - [" & IniSection & "], " & IniKey & " (" & sTemp &
            '") is invalid. Using default (" & Format(DefaultDate, "MM/dd/yyyy") & ")")
            '      ReadIniDate = DefaultDate
            '  End If

        Else
            ReadIniDate = DefaultDate
        End If
  End Function

  ' KC002
  Private Function ReadIniTime(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String, _
    ByVal DefaultTime As String) As Date
    Dim sTemp As String

    sTemp = Trim(ReadINI(IniFile, IniSection, IniKey, ""))
    If sTemp <> "" Then
      If IsDate(sTemp) Then
        Try
          ReadIniTime = TimeValue(sTemp)
        Catch ex As Exception
          Call LogToFile("  Error: ReadIniTime - [" & IniSection & "], " & IniKey & " (" & sTemp & _
            ") cannot be converted. Using default (" & DefaultTime & ") - " & ex.Message)
          ReadIniTime = TimeValue(DefaultTime)
        End Try
      Else
        Call LogToFile("  Error: ReadIniTime - [" & IniSection & "], " & IniKey & " (" & sTemp & _
          ") is invalid. Using default (" & DefaultTime & ")")
        ReadIniTime = TimeValue(DefaultTime)
      End If
    Else
      ReadIniTime = TimeValue(DefaultTime)
    End If
  End Function

    Private Function ReadIniBoolean(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String, _
      ByVal DefaultBoolean As Boolean) As Boolean
        Dim sTemp As String
        Dim iLen As Integer

        ReadIniBoolean = DefaultBoolean

        sTemp = UCase(Trim(ReadINI(IniFile, IniSection, IniKey, "")))
        If sTemp <> "" Then
            iLen = Len(sTemp)

            If DefaultBoolean Then
                If sTemp = Microsoft.VisualBasic.Left("NO", iLen) Then
                    ReadIniBoolean = False
                End If
            ElseIf sTemp = Microsoft.VisualBasic.Left("YES", iLen) Then
                ReadIniBoolean = True
            End If
        End If
    End Function

    Private Function ReadIniDecimal(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String, _
      ByVal RangeFilter As Integer, ByVal DefaultDecimal As Decimal) As Decimal
        Dim sTemp As String

        sTemp = Trim(ReadINI(IniFile, IniSection, IniKey, ""))
        If sTemp <> "" Then
            If IsNumeric(sTemp) Then
                Try
                    ReadIniDecimal = CDec(sTemp)
                    Select Case RangeFilter
                        Case 0
                            If ReadIniDecimal < 0 Then
                                Call LogToFile("  Error: ReadIniDecimal - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                                  ") is invalid (< 0). Using default (" & CStr(DefaultDecimal) & ")")
                                ReadIniDecimal = DefaultDecimal
                            End If
                        Case 1
                            If ReadIniDecimal < 1 Then
                                Call LogToFile("  Error: ReadIniDecimal - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                                  ") is invalid (< 1). Using default (" & CStr(DefaultDecimal) & ")")
                                ReadIniDecimal = DefaultDecimal
                            End If
                    End Select
                Catch ex As Exception
                    Call LogToFile("  Error: ReadIniDecimal - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                      ") cannot be converted. Using default (" & CStr(DefaultDecimal) & ") - " & ex.Message)
                    ReadIniDecimal = DefaultDecimal
                End Try
            Else
                Call LogToFile("  Error: ReadIniDecimal - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                  ") is invalid. Using default (" & CStr(DefaultDecimal) & ")")
                ReadIniDecimal = DefaultDecimal
            End If
        Else
            ReadIniDecimal = DefaultDecimal
        End If
    End Function

    Private Function ReadIniLong(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String, _
      ByVal RangeFilter As Integer, ByVal DefaultLong As Long) As Long
        Dim sTemp As String

        sTemp = Trim(ReadINI(IniFile, IniSection, IniKey, ""))
        If sTemp <> "" Then
            If IsNumeric(sTemp) Then
                Try
                    ReadIniLong = CLng(sTemp)
                    Select Case RangeFilter
                        Case 0
                            If ReadIniLong < 0 Then
                                Call LogToFile("  Error: ReadIniLong - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                                  ") is invalid (< 0). Using default (" & CStr(DefaultLong) & ")")
                                ReadIniLong = DefaultLong
                            End If
                        Case 1
                            If ReadIniLong < 1 Then
                                Call LogToFile("  Error: ReadIniLong - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                                  ") is invalid (< 1). Using default (" & CStr(DefaultLong) & ")")
                                ReadIniLong = DefaultLong
                            End If
                    End Select
                Catch ex As Exception
                    Call LogToFile("  Error: ReadIniLong - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                      ") cannot be converted. Using default (" & CStr(DefaultLong) & ") - " & ex.Message)
                    ReadIniLong = DefaultLong
                End Try
            Else
                Call LogToFile("  Error: ReadIniLong - [" & IniSection & "], " & IniKey & " (" & sTemp & _
                  ") is invalid. Using default (" & CStr(DefaultLong) & ")")
                ReadIniLong = DefaultLong
            End If
        Else
            ReadIniLong = DefaultLong
        End If
    End Function

    Private Sub Finish()
        Call LogToFile("End of process")
        End
    End Sub

    Private Function RegValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String, _
      ByRef ErrInfo As String) As String
        Dim objParent As RegistryKey
        Dim objSubkey As RegistryKey
        Dim sAns As String = ""

        RegValue = ""

        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.DynData
                objParent = Registry.DynData
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
            Case Else
                Exit Function
        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            ' If can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If
        Catch ex As Exception
            ErrInfo = ex.Message
        Finally
            ' If no error but value is empty, populate ErrInfo
            If ErrInfo = "" And sAns = "" Then
                ErrInfo = "No value found for requested registry key"
            End If
        End Try

        Return sAns
    End Function

    Private Function CreateSQLConnStr(ByVal UserId As String, ByVal Password As String, ByVal Server As String, _
      ByVal DatabaseName As String) As String

        CreateSQLConnStr = ""

        If UserId <> "" Then
            CreateSQLConnStr = "User Id=" & UserId
        End If
        If Password <> "" Then
            CreateSQLConnStr = CreateSQLConnStr & ";Password=" & Password
        End If
        If Server <> "" Then
            CreateSQLConnStr = CreateSQLConnStr & ";Data Source=" & Server
        End If
        If DatabaseName <> "" Then
            CreateSQLConnStr = CreateSQLConnStr & ";Initial Catalog=" & DatabaseName
        End If
    End Function

    Private Function ReplaceString(ByVal InString As String, ByVal RemoveString As String, _
  ByVal AddString As String) As String
        Dim iPosition As Integer
        Dim sFront As String
        Dim sBack As String

        iPosition = FindString(InString, RemoveString)
        If iPosition > 0 Then
            sFront = Microsoft.VisualBasic.Left(InString, iPosition - 1)
            sBack = Mid(InString, iPosition + Len(RemoveString))
            ReplaceString = sFront & AddString & sBack
        Else
            ReplaceString = InString
        End If
    End Function

    Private Function FindString(ByVal SearchStr As String, ByVal SubStr As String) As Integer
        Dim iPosition As Integer

        iPosition = InStr(SearchStr, SubStr)
        Do While iPosition > 0
            If Len(SearchStr) - (iPosition - 1) = Len(SubStr) Then    ' Look for end of string
                FindString = iPosition
                Exit Do
            ElseIf Mid(SearchStr, iPosition, Len(SubStr) + 1) = SubStr & "," Then   ' Look for comma delimiter
                FindString = iPosition
                Exit Do
            Else
                iPosition = InStr(iPosition + 1, SearchStr, SubStr)
            End If
        Loop
    End Function

    Private Function ConvertDBType(ByVal DatabaseType As String) As String
        Dim iLen As Integer

        ConvertDBType = ""

        If DatabaseType <> "" Then
            iLen = Len(DatabaseType)
            If DatabaseType = Microsoft.VisualBasic.Left("WS", iLen) Then
                ConvertDBType = "WS"
            Else
                Call LogToFile("  Error: ConvertDBType - DatabaseType (" & DatabaseType & ") is invalid")
            End If
        End If
    End Function

    Private Function ReadOrdinals(ByRef Ordinals As ConfigOrdinals, _
        ByVal Reader As System.Data.SqlClient.SqlDataReader) As Boolean

        ReadOrdinals = False

        Try
            Ordinals.SourceType = Reader.GetOrdinal("Type")

            Ordinals.SQLServer = Reader.GetOrdinal("SQLServer")
            Ordinals.SQLUserName = Reader.GetOrdinal("SQLUserName")
            Ordinals.SQLPassword = Reader.GetOrdinal("SQLPassword")
            Ordinals.SQLDBName = Reader.GetOrdinal("SQLDBName")

            Ordinals.ClientFirm = Reader.GetOrdinal("ClientFirm")

            ReadOrdinals = True
        Catch ex As Exception
            Call LogToFile("  Error: ReadOrdinals - Unable to read ordinals - " & ex.Message)
        End Try
    End Function
    Private Sub ReadIniSourceConfig()
        Dim iIndex As Integer
        Dim sTemp As String
        Dim bValidWS As Boolean

        For iIndex = 0 To (MAX_SOURCES - 1)
            sTemp = ReadIniString(gsIniFile, "Source" & CStr(iIndex), "DatabaseType", 0, 1, False, "")
            gaSourcesList(iIndex).DatabaseType = ConvertDBType(sTemp)

            If gaSourcesList(iIndex).DatabaseType = "WS" Then
                gaSourcesList(iIndex).WSUrl = ReadIniString(gsIniFile, "Source" & CStr(iIndex), "WebServiceURL", 0, 0, _
                  True, "")
                gaSourcesList(iIndex).WSUserName = ReadIniString(gsIniFile, "Source" & CStr(iIndex), "WebServiceUserName", _
                  0, 0, True, "")
                gaSourcesList(iIndex).WSCompanyName = ReadIniString(gsIniFile, "Source" & CStr(iIndex), _
                  "WebServiceCompanyName", 0, 0, True, "")
                gaSourcesList(iIndex).WSPassword = ReadIniString(gsIniFile, "Source" & CStr(iIndex), "WebServicePassword", _
                  0, 0, True, "")
                gaSourcesList(iIndex).WSIOS = ReadIniString(gsIniFile, "Source" & CStr(iIndex), "WebServiceIOS", 0, 0, True, "")
                gaSourcesList(iIndex).WSBaseApplicationId = ReadIniString(gsIniFile, "Source" & CStr(iIndex), _
                  "WebServiceApplicationID", 0, 0, True, "")
                If gbAddComputerName Then
                    gaSourcesList(iIndex).WSApplicationId = gsComputerName & "-" & gaSourcesList(iIndex).WSBaseApplicationId
                Else
                    gaSourcesList(iIndex).WSApplicationId = gaSourcesList(iIndex).WSBaseApplicationId
                End If

                bValidWS = True
            Else
                Exit For
            End If
        Next

        ' Check for WS connection
        If Not bValidWS Then
            Call LogToFile("  Error: ReadIniSourceConfig - No valid WS Source")
            Call Finish()
        End If
    End Sub

    Private Function ReadSQLSourceConfig() As Boolean
        Dim myConnection As New SqlClient.SqlConnection(gsConfigSQLConnectStr)
        Dim myCommand As New SqlClient.SqlCommand("GET_DATA_SOURCES", myConnection)
        myCommand.CommandType = CommandType.StoredProcedure
        myCommand.Parameters.Add("@AppType", SqlDbType.VarChar, 50).Value = APP_NAME
        myCommand.Parameters.Add("@Instance", SqlDbType.VarChar, 50).Value = gsInstanceName
        myCommand.Parameters.Add("@Firm", SqlDbType.VarChar, 50).Value = giBrokerNumber.ToString
        myCommand.Parameters.Add("@PrimarySource", SqlDbType.VarChar, 1).Value = IIf(gbPrimary, "Y", "N")
        Dim myReader As SqlClient.SqlDataReader
        Dim SourceOrdinals As ConfigOrdinals
        Dim iIndex As Integer
        Dim bValidWS As Boolean
        Dim bError As Boolean

        ReadSQLSourceConfig = False

        Try
            myConnection.Open()
        Catch ex As Exception
            Call LogToFile("  Error: ReadSQLSourceConfig - Unable to open SQL connection - " & ex.Message)
            Exit Function
        End Try

        myCommand.CommandTimeout = glSQLTimeOut
        Try
            myReader = myCommand.ExecuteReader()
        Catch ex As Exception
            Call LogToFile("  Error: ReadSQLSourceConfig - Unable to open SQL reader - " & ex.Message)
            myConnection.Close()
            Exit Function
        End Try

        If Not ReadOrdinals(SourceOrdinals, myReader) Then
            myReader.Close()
            myConnection.Close()
            Exit Function
        End If

        Try
            While myReader.Read() And iIndex < MAX_SOURCES
                If ProcessSourceRecord(myReader, SourceOrdinals, iIndex, bValidWS) Then
                    iIndex = iIndex + 1
                Else
                    bError = True
                    Exit While
                End If
            End While
        Catch ex As Exception
            Call LogToFile("  Error: ReadSQLSourceConfig - Unable to read SQL record - " & ex.Message)
            bError = True
        End Try

        Try
            myReader.Close()
        Catch ex As Exception
            Call LogToFile("  Error: ReadSQLSourceConfig - Unable to close SQL reader - " & ex.Message)
        End Try
        myReader = Nothing

        Try
            myConnection.Close()
        Catch ex As Exception
            Call LogToFile("  Error: ReadSQLSourceConfig - Unable to close SQL connection - " & ex.Message)
        End Try
        myConnection = Nothing

        If Not bError And bValidWS Then
            ReadSQLSourceConfig = True
        End If
    End Function

    Private Function ProcessSourceRecord(ByVal Reader As System.Data.SqlClient.SqlDataReader, _
      ByVal Ordinals As ConfigOrdinals, ByVal Session As Integer, ByRef ValidWS As Boolean) As Boolean
        Dim sTemp As String = ""

        ProcessSourceRecord = False

        If Not Reader.IsDBNull(Ordinals.SourceType) Then
            sTemp = UCase(Trim(Reader.GetString(Ordinals.SourceType)))
        End If
        If sTemp = "WS" Then
            If Not ProcessWSConfigRecord(Reader, Ordinals, Session) Then
                Exit Function
            End If

            ValidWS = True
        Else
            Exit Function
        End If

        ProcessSourceRecord = True
    End Function

    Private Function ProcessWSConfigRecord(ByVal Reader As System.Data.SqlClient.SqlDataReader, _
      ByVal Ordinals As ConfigOrdinals, ByVal Session As Integer) As Boolean
        Dim sWSUserName As String = ""
        Dim sWSPassword As String = ""
        Dim sWSURL As String = ""
        Dim sWSCompanyName As String = ""
        Dim sWSIOS As String = ""

        ProcessWSConfigRecord = False

        If Not Reader.IsDBNull(Ordinals.SQLUserName) Then
            sWSUserName = Trim(Reader.GetString(Ordinals.SQLUserName))
        End If
        If sWSUserName = "" Then
            Call LogToFile("  Error: ProcessWSConfigRecord (Source" & CStr(Session) & ") - WSUserName is invalid")
            Exit Function
        End If
        If Not Reader.IsDBNull(Ordinals.SQLPassword) Then
            sWSPassword = Trim(Reader.GetString(Ordinals.SQLPassword))
        End If
        If Not Reader.IsDBNull(Ordinals.SQLServer) Then
            sWSURL = Trim(Reader.GetString(Ordinals.SQLServer))
        End If
        If sWSURL = "" Then
            Call LogToFile("  Error: ProcessWSConfigRecord (Source" & CStr(Session) & ") - WSUrl is invalid")
            Exit Function
        End If
        If Not Reader.IsDBNull(Ordinals.SQLDBName) Then
            sWSCompanyName = Trim(Reader.GetString(Ordinals.SQLDBName))
        End If
        If sWSCompanyName = "" Then
            Call LogToFile("  Error: ProcessWSConfigRecord (Source" & CStr(Session) & ") - WSCompanyName is invalid")
            Exit Function
        End If
        If Not Reader.IsDBNull(Ordinals.ClientFirm) Then
            sWSIOS = Trim(Reader.GetString(Ordinals.ClientFirm))
        End If
        If sWSIOS = "" Then
            Call LogToFile("  Error: ProcessWSConfigRecord (Source" & CStr(Session) & ") - WSIOS is invalid")
            Exit Function
        End If

        gaSourcesList(Session).DatabaseType = "WS"
        gaSourcesList(Session).WSUrl = sWSURL
        gaSourcesList(Session).WSUserName = sWSUserName
        gaSourcesList(Session).WSCompanyName = sWSCompanyName
        gaSourcesList(Session).WSPassword = sWSPassword
        gaSourcesList(Session).WSIOS = sWSIOS

        gaSourcesList(Session).WSBaseApplicationId = APP_NAME & "-" & gsInstanceName
        If gbAddComputerName Then
            gaSourcesList(Session).WSApplicationId = gsComputerName & "-" & gaSourcesList(Session).WSBaseApplicationId
        Else
            gaSourcesList(Session).WSApplicationId = gaSourcesList(Session).WSBaseApplicationId
        End If

        ProcessWSConfigRecord = True
    End Function

    Private Sub SaveSource()
        Dim iIndex As Integer

        For iIndex = 0 To (MAX_SOURCES - 1)
            If gaSourcesList(iIndex).DatabaseType = "WS" Then
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "DatabaseType", "WS")

                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServiceURL", gaSourcesList(iIndex).WSUrl)
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServiceUserName", gaSourcesList(iIndex).WSUserName)
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServiceCompanyName", gaSourcesList(iIndex).WSCompanyName)
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServicePassword", gaSourcesList(iIndex).WSPassword)
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServiceIOS", gaSourcesList(iIndex).WSIOS)
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), "WebServiceApplicationID", _
                  gaSourcesList(iIndex).WSBaseApplicationId)
            Else
                ' Clear section
                Call WriteINI(gsIniFile, "Source" & CStr(iIndex), Nothing, Nothing)
            End If
        Next
    End Sub

    Private Function ReadIniFullPath(ByVal IniFile As String, ByVal IniSection As String, ByVal IniKey As String,
    ByVal InDate As Date, ByVal CriticalError As Boolean, ByVal DefaultFullPath As String) As String
        Dim sTemp As String
        Dim sFullPath As String
        Dim sDirectory As String
        Dim vDirInfo As DirectoryInfo

        ReadIniFullPath = ""

        sTemp = ReadIniString(IniFile, IniSection, IniKey, 0, 0, False, "")
        If sTemp <> "" Then
            sFullPath = CheckForDateVariables(sTemp, InDate)

            ' Check directory
            sDirectory = Path.GetDirectoryName(sFullPath)

            If sDirectory <> "" Then
                vDirInfo = New DirectoryInfo(sDirectory)

                If vDirInfo.Exists Then
                    ReadIniFullPath = sFullPath
                ElseIf Not CriticalError Then
                    ReadIniFullPath = DefaultFullPath
                Else
                    Call LogToFile("  Error: ReadIniFullPath - [" & IniSection & "], " & IniKey & " (" & sFullPath &
            ") is invalid")
                    Call Finish()
                End If
            Else
                ' Add app path
                ReadIniFullPath = gsAppPath & "\" & sFullPath
            End If
        ElseIf Not CriticalError Then
            ReadIniFullPath = DefaultFullPath
        Else
            Call LogToFile("  Error: ReadIniFullPath - [" & IniSection & "], " & IniKey & " (" & sTemp & ") is invalid")
            Call Finish()
        End If
    End Function

  Private Function CheckForDateVariables(ByVal InString As String, ByVal InDate As Date) As String
    CheckForDateVariables = InString

    If InString <> "" Then
      CheckForDateVariables = InString.Replace("<D1>", Format(InDate, "yyMMdd")).Replace("<D2>", Format(InDate,
        "MMdd")).Replace("<D3>", Format(InDate, "yyMMdd") & "-" & Format(TimeOfDay, "hhmmss")).Replace("<D4>",
        Format(InDate, "yyyyMMdd")).Replace("<D5>", Format(InDate, "MMddyy")).Replace("<D6>", Format(InDate,
        "MM_dd")).Replace("<D7>", Format(InDate, "yyyyMM")).Replace("<D8>", Format(InDate, "MMM dd yyyy"))
    End If
  End Function

    Private Function CheckPath(ByRef FullPath As String, ByVal DefaultPath As String, ByVal InDate As Date, _
      ByVal CheckType As Integer, ByVal CreatePath As Boolean, ByVal ValidatePath As Boolean, _
      ByVal CriticalError As Boolean) As String
        Dim sFullPath As String
        Dim bUNCPath As Boolean
        Dim sTemp As String
        Dim sDirectory As String

        sFullPath = FullPath
        If Microsoft.VisualBasic.Left(sFullPath, 2) = "\\" Then
            bUNCPath = True
        End If
        sTemp = StrTok("\", sFullPath)
        If sFullPath = "" Then
            CheckPath = DefaultPath
            bUNCPath = False
        Else
            CheckPath = CheckVariable(sTemp, InDate)
            sTemp = StrTok("\", sFullPath)
            Do While sFullPath <> ""
                CheckPath = CheckPath & "\" & CheckVariable(sTemp, InDate)
                sTemp = StrTok("\", sFullPath)
            Loop
            ' CheckType: 0 - directory, 1 - fullpath file
            If CheckType = 0 Then
                CheckPath = CheckPath & "\" & CheckVariable(sTemp, InDate)
            End If
        End If

        If bUNCPath Then
            CheckPath = "\\" & CheckPath
        End If
        Try
            sDirectory = Dir(CheckPath, vbDirectory)
        Catch ex As Exception
            sDirectory = ""
        End Try
        If sDirectory = "" Then
            If CreatePath Then
                Call LogToFile("  Info: CheckPath - Directory (" & CheckPath & ") does not exist. Creating directory")
                Try
                    Directory.CreateDirectory(CheckPath)
                Catch ex As Exception
                    Call LogToFile("  Error: CheckPath - Unable to create Directory (" & CheckPath & ") - " & ex.Message)
                End Try
            ElseIf ValidatePath Then
                If Not CriticalError Then
                    Call LogToFile("  Error: CheckPath - Directory (" & CheckPath & ") is invalid. Using default (" & _
                      DefaultPath & ")")
                    CheckPath = DefaultPath
                Else
                    Call LogToFile("  Error: CheckPath - (" & CheckPath & ") is invalid")
                    Call Finish()
                End If
            End If
        End If

        If CheckType = 1 Then
            ' Retain filename
            FullPath = sTemp
        End If
    End Function

    Private Function StrTok(ByVal Delimiters As String, ByRef CallString As String) As String
        Dim iDelimLen As Integer
        Dim iMinPos As Integer
        Dim iIndex As Integer
        Dim sDelimChar As String
        Dim iDelimPos As Integer

        iDelimLen = Len(Delimiters)
        If iDelimLen < 1 Then
            Call LogToFile("  Error: StrTok - Delimiter is missing")
            StrTok = ""
            Exit Function
        End If

        iMinPos = Len(CallString) + 1
        If iMinPos > 1 Then
            ' Find the earliest occurrence of a valid delimiter
            For iIndex = 1 To iDelimLen
                sDelimChar = Mid(Delimiters, iIndex, 1)
                iDelimPos = InStr(CallString, sDelimChar)
                If iDelimPos < iMinPos And iDelimPos > 0 Then
                    iMinPos = iDelimPos
                End If
            Next
            If iMinPos = 1 Then
                ' Ignore preceding delimiters
                CallString = Mid(CallString, 2)
                StrTok = StrTok(Delimiters, CallString)
            Else
                StrTok = Microsoft.VisualBasic.Left(CallString, iMinPos - 1)
                CallString = Mid(CallString, iMinPos + 1)
            End If
        Else
            StrTok = ""
        End If
    End Function

    Private Function CheckVariable(ByVal CallString As String, ByVal InDate As Date) As String
        Dim iPos As Integer
        Dim sToken As String
        Dim sFront As String

        Do
            iPos = InStr(1, CallString, "<")
            If iPos = 0 Then
                ' No variable
                Exit Do
            ElseIf iPos = 1 Then
                ' Variable in front
                sFront = StrTok("<", CallString)
                sToken = StrTok(">", sFront)
                CallString = ConvertVariable(sToken, InDate) & sFront
            Else
                ' Variable in middle or end
                sFront = StrTok("<", CallString)
                sToken = StrTok(">", CallString)
                CallString = sFront & ConvertVariable(sToken, InDate) & CallString
            End If
        Loop

        CheckVariable = CallString
    End Function

    Private Function ConvertVariable(ByVal Variable As String, ByVal InDate As Date) As String
        Dim sVariable As String

        sVariable = UCase(Variable)
        Select Case sVariable
            Case "D1"
                ConvertVariable = Format(InDate, "yyMMdd")
            Case "D2"
        ConvertVariable = Format(InDate, "MMdd")
            Case "D3"
                ConvertVariable = Format(InDate, "yyMMdd") & "-" & Format(TimeOfDay, "hhmmss")
            Case "D4"
                ConvertVariable = Format(InDate, "yyyyMMdd")
            Case "D5"
                ConvertVariable = Format(InDate, "MMddyy")
            Case "D6"
                ConvertVariable = Format(InDate, "MM_dd")
            Case "D7"
                ConvertVariable = Format(InDate, "yyyyMM")
            Case "B1"
                ConvertVariable = Format(giBrokerNumber, "000")
            Case Else
                Call LogToFile("  Error: ConvertVariable - Variable (" & Variable & ") is invalid")
                ConvertVariable = ""
        End Select
    End Function

    Private Function CheckName(ByVal Name As String, ByVal InDate As Date) As String
        Dim sTemp As String

        If gbDOSFileNameFormat Then
            sTemp = StrTok(".", Name)
            CheckName = CheckVariable(sTemp, InDate)
            ' Extension
            If Name <> "" Then
                CheckName = CheckName & "." & Name
            End If
        Else
            CheckName = CheckVariable(Name, InDate)
        End If
    End Function

    Private Function GetTaggedFieldFromTMXI(ByVal TradeMarkers As String, ByVal OrderAttribute As String, _
  ByVal AttributeDescription As String, ByVal ExecutionInstructions As String) As String

        GetTaggedFieldFromTMXI = ""

        If TradeMarkers <> "" Then
            GetTaggedFieldFromTMXI = GetTaggedFieldFromColumn(OrderAttribute, TradeMarkers, AttributeDescription)
        End If

        If GetTaggedFieldFromTMXI = "" Then
            If ExecutionInstructions <> "" Then
                GetTaggedFieldFromTMXI = GetTaggedFieldFromColumn(OrderAttribute, ExecutionInstructions, AttributeDescription)
            End If
        End If
    End Function

    Private Function GetTaggedFieldFromColumn(ByVal OrderAttribute As String, ByVal Column As String, _
      ByVal AttributeDescription As String) As String

        GetTaggedFieldFromColumn = ""

        If OrderAttribute <> "" Then
            GetTaggedFieldFromColumn = ParseTaggedString(OrderAttribute, Column)
        End If

        If GetTaggedFieldFromColumn = "" Then
            If AttributeDescription <> "" Then
                GetTaggedFieldFromColumn = ParseTaggedString(AttributeDescription, Column)
            End If
        End If
    End Function

    Private Function ParseTaggedString(ByVal Tag As String, ByVal TaggedString As String) As String
        Dim iTagLen As Integer
        Dim sTemp As String
        Dim iLen As Integer

        ParseTaggedString = ""

        iTagLen = Len(Tag)
        sTemp = Trim(StrTok(",", TaggedString))
        Do While sTemp <> ""
            If (Microsoft.VisualBasic.Left(sTemp, iTagLen + 1) = Tag & "(") And _
              (Microsoft.VisualBasic.Right(sTemp, 1) = ")") Then
                iLen = Len(sTemp)
                ParseTaggedString = Mid(sTemp, iTagLen + 2, iLen - (iTagLen + 2))
                Exit Function
            End If
            sTemp = Trim(StrTok(",", TaggedString))
        Loop
    End Function




End Class