Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain
#Region "Attributes"
    Const APP_NAME As String = "WSDFDTradesReport"
    Const MAX_SOURCES As Integer = 20

    Const REPORT_USER_SECTION As String = "AccountId"

    Const REPORT_KEY_DESTINATIONS As String = "Destinations"
    Const REPORT_KEY_EXCEPTION_DESTINATIONS As String = "ExceptionDestinations"
    Const REPORT_KEY_EXCHANGES As String = "Exchanges"
    Const REPORT_KEY_EXCEPTION_EXCHANGES As String = "ExceptionExchanges"

    Const REPORT_KEY_CURRENCY As String = "Currency"
    Const REPORT_KEY_ACCOUNT_TYPES As String = "AccountTypes"
    Const REPORT_KEY_ACCOUNT_IDS As String = "AccountIds"
    Const REPORT_KEY_EXCEPTION_ACCOUNT_IDS As String = "ExceptionAccountIds"

    Const REPORT_KEY_SYMBOLS As String = "Symbols"
    Const REPORT_KEY_EXCEPTION_SYMBOLS As String = "ExceptionSymbols"

    Const REPORT_ID_KEY As String = "AccountCode"
    Const MAX_REPORT_ACCOUNTS As Integer = 300
    Const MAX_ENTRIES As Integer = 600

    Const MAX_ARRAY_COLUMNS As Integer = 18
    Const MAX_ARRAY_ROWS As Integer = 1000

  Dim gsAppPath As String
  Dim gdDate As Date
  Dim gbModifiedDate As Boolean   ' KC002

    Dim gsIniFile As String
    Dim gsLogFile As String
    Dim gsBackupLog As String

    Dim gsInstanceName As String
    Dim gbReadConfigFromINI As Boolean

    Dim gsConfigSQLConnectStr As String
    Dim gbPrimary As Boolean

    Dim gdStartDay As DateTime
    Dim gdStartTime As DateTime
    Dim gdEndTime As DateTime
	
    Dim gbDOSFileNameFormat As Boolean
    Dim gbAddComputerName As Boolean
    Dim gsComputerName As String

    Dim gbSQLNoLock As Boolean
    Dim glSQLTimeOut As Long
    Dim giBrokerNumber As Integer
    Dim gbDFD As Boolean
    Dim gbTradedOrderOnly As Boolean
    Dim gbTempToFinal As Boolean
    Dim gbNetting As Boolean
    Dim gbRemoveTemp As Boolean


  Structure Source
    Dim DatabaseType As String
    Dim WSUrl As String
    Dim WSUserName As String
    Dim WSCompanyName As String
    Dim WSPassword As String
    Dim WSIOS As String
    Dim WSBaseApplicationId As String
    Dim WSApplicationId As String
    Dim TradesFile As String    ' KC001
  End Structure
  Dim gaSourcesList(MAX_SOURCES - 1) As Source

    Structure ConfigOrdinals
        Dim SourceType As Integer
        Dim SQLServer As Integer
        Dim SQLUserName As Integer
        Dim SQLPassword As Integer
        Dim SQLDBName As Integer
        Dim ClientFirm As Integer
    End Structure

    Dim gbOutput(1) As Boolean

    Dim gsOutputName As String
    Dim gsTempOutputName As String

  Dim gsReportingType As String   ' KC001

    Dim gsCsvDelimiter As String
    Dim gbCsvHeader As Boolean

  Dim giOrderFilter As Integer
  Dim gsTplus3Exchanges As String   ' KC003
  Dim gsSelttlementPath As String
  Dim gsExchangeToCountryPath As String

  Dim gsSettleTradeDateFormat As String   ' KC006
  Dim gbGroupSort As Boolean
    Dim gsSortOrder As String

    Dim gsDefaultReportDestinations As String
    Dim gsDefaultReportExceptionDestinations As String
    Dim gsDefaultReportExchanges As String
    Dim gsDefaultReportExceptionExchanges As String

    Dim gsDefaultReportCurrency As String
    Dim gsDefaultReportAccountTypes As String
    Dim gsDefaultReportAccountIds As String
    Dim gsDefaultReportExceptionAccountIds As String

    Dim gsDefaultReportSymbols As String
    Dim gsDefaultReportExceptionSymbols As String

    Dim gsDefaultReportOrderTypes As String
    Dim gsDefaultReportExceptionOrderTypes As String

    Dim gsOrderTypes As String()
    Dim gsExceptionOrderTypes As String()
    Dim htOrderTypes As New Hashtable
    Dim htExceptionOrderTypes As New Hashtable

    Dim htExchangeCurrencys As New Hashtable
    Dim htExchangeMultipliers As New Hashtable



    Dim htReportAccountsTable As New Hashtable
    Dim htReportOrdersTable As New Hashtable
    Dim htDNDOrdersTable As New Hashtable

    Dim htHolidays As Hashtable
    Dim giReportServerTimeZone As Int64
    Dim giTradeServerTimeZone As Int64

    Structure ReportAccount
        Dim AccCode As String
        Dim Destinations As String
        Dim ExceptionDestinations As String
        Dim Exchanges As String
        Dim ExceptionExchanges As String
        Dim Currency As String
        Dim AccountTypes As String
        Dim AccountIds As String
        Dim ExceptionAccountIds As String
        Dim Symbols As String
        Dim ExceptionSymbols As String
    End Structure

    Dim gaReportAccountsList(MAX_REPORT_ACCOUNTS - 1) As ReportAccount

    Dim dtWSIOSOrdersTable As Data.DataTable
    Dim dtReportOrdersTable As Data.DataTable
    Dim dtNettingReportOrdersTable As Data.DataTable
    Dim dtWSTradesTable As Data.DataTable
    Dim giWsSecIndex As Int32

    Dim htWSSecTable As New Hashtable

    Structure SecInfo
        Dim CUSIP As String
        Dim ISIN As String
        Dim SEDOL As String
        Dim Description As String
        Dim UkIrishStampDutyReserveTaxMarker As String
        Dim PtmLevyIndicator As String
        Dim CurrencyDenomination As String
    End Structure
    Dim gaWSSecList(100000 - 1) As SecInfo


    Dim wsIRESS As New IRESS.IRESSSoapClient
    Dim myIRESSSessionKey As String
    Dim wsIOSPlus As New IOSPlus.IOSPLUSSoapClient
    Dim myIOSPlusSessionKey As String
    Dim myIOSPlusServiceSessionKey As String
    Dim glWSTimeout As Long
  Dim giIressSession As Int32

  Dim htProcessedTradesTable As New Hashtable   ' KC001

    Dim glRow As Long
    Dim giFileNumber As Integer
    Dim DataArray(,) As Object
    Dim giArrayRow As Integer

#End Region

End Class