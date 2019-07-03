Imports System.IO
Imports Microsoft.Win32
Imports System.Globalization
Partial Class frmMain
#Region "API Connection"

    Private Function CreateIRESSSession(ByVal Session As Integer) As Boolean
        CreateIRESSSession = False
        CreateIRESSSession = CreateIRESSSession(gaSourcesList(Session).WSUrl, gaSourcesList(Session).WSUserName, gaSourcesList(Session).WSCompanyName, gaSourcesList(Session).WSPassword, gaSourcesList(Session).WSApplicationId)
    End Function

    Private Function CreateIOSPlusSession(ByVal Session As Integer) As Boolean
        CreateIOSPlusSession = False
        CreateIOSPlusSession = CreateIOSPlusSession(gaSourcesList(Session).WSUrl, gaSourcesList(Session).WSUserName, gaSourcesList(Session).WSCompanyName, gaSourcesList(Session).WSPassword, gaSourcesList(Session).WSApplicationId, myIOSPlusSessionKey)
    End Function

    Private Function CreateIRESSSession(ByVal URL As String, ByVal UserName As String, ByVal CompanyName As String, _
    ByVal Password As String, ByVal ApplicationId As String) As Boolean

        CreateIRESSSession = False

        Try
            ' Set the webservice url
            wsIRESS.Endpoint.Address = New System.ServiceModel.EndpointAddress(URL)
            wsIRESS.Endpoint.Binding.OpenTimeout = New TimeSpan(0, 0, glWSTimeout)
            wsIRESS.Endpoint.Binding.SendTimeout = New TimeSpan(0, 0, glWSTimeout)
            wsIRESS.Endpoint.Binding.ReceiveTimeout = New TimeSpan(0, 0, glWSTimeout)

            ' Create an IRESS session using the IRESSSessionStart method
            Dim issRequest As IRESS.IRESSSessionStartInput = New IRESS.IRESSSessionStartInput

            ' Initialize the parameters of our IRESS session
            issRequest.Parameters = New IRESS.IRESSSessionStartInputParameters
            issRequest.Parameters.UserName = UserName
            issRequest.Parameters.CompanyName = CompanyName
            issRequest.Parameters.Password = Password
            issRequest.Parameters.ApplicationID = ApplicationId
            issRequest.Parameters.SessionNumberToKick = ReadIniLong(gsIniFile, APP_NAME, "IRESSSessionNumber", 0, 0)

            ' Call the IRESSSessionStart method to create the IRESS session (equivalent to logging in via the front-end)
            Dim issResult As IRESS.IRESSSessionStartOutput = wsIRESS.IRESSSessionStart(issRequest)

            If IsNothing(issResult) Then
                Throw New Exception("Failed to start IRESS session")
            End If

            ' Obtain the IRESS session key from the response of the IRESSSessionStart method
            myIRESSSessionKey = issResult.Result.DataRows(0).IRESSSessionKey

            'WriteINI(gsIniFile, APP_NAME, "IRESSSessionNumber", myIRESSSessionKey.ToString)

            CreateIRESSSession = True
        Catch ex As Exception
            Call LogToFile("  Error: StartIRESSSession - IRESS.IRESSSessionStart - " & ex.Message)
        End Try
    End Function


    Private Function CreateIOSPlusSession(ByVal URL As String, _
      ByVal UserName As String, ByVal CompanyName As String, ByVal Password As String, ByVal ApplicationId As String, _
      ByRef myIOSPlusSessionKey As String) As Boolean

        CreateIOSPlusSession = False

        Try
            ' Set the webservice url
            wsIOSPlus.Endpoint.Address = New System.ServiceModel.EndpointAddress(URL)
            wsIOSPlus.Endpoint.Binding.OpenTimeout = New TimeSpan(0, 0, glWSTimeout)
            wsIOSPlus.Endpoint.Binding.SendTimeout = New TimeSpan(0, 0, glWSTimeout)
            wsIOSPlus.Endpoint.Binding.ReceiveTimeout = New TimeSpan(0, 0, glWSTimeout)

            ' Create an IRESS session using the IRESSSessionStart method
            Dim issRequest As IOSPlus.IRESSSessionStartInput = New IOSPlus.IRESSSessionStartInput

            ' Initialize the parameters of our IRESS session
            issRequest.Parameters = New IOSPlus.IRESSSessionStartInputParameters
            issRequest.Parameters.UserName = UserName
            issRequest.Parameters.CompanyName = CompanyName
            issRequest.Parameters.Password = Password
            issRequest.Parameters.ApplicationID = ApplicationId
            issRequest.Parameters.SessionNumberToKick = ReadIniLong(gsIniFile, APP_NAME, "IOSPlusSessionNumber", 0, 0)
            ' Call the IRESSSessionStart method to create the IRESS session (equivalent to logging in via the front-end)
            Dim issResult As IOSPlus.IRESSSessionStartOutput = wsIOSPlus.IRESSSessionStart(issRequest)

            If IsNothing(issResult) Then
                Throw New Exception("Failed to start IRESS session")
            End If

            ' Obtain the IRESS Session Key from the response of the IRESSSessionStart method
            myIOSPlusSessionKey = issResult.Result.DataRows(0).IRESSSessionKey
            'WriteINI(gsIniFile, APP_NAME, "IOSPlusSessionNumber", myIOSPlusSessionKey.ToString)

            CreateIOSPlusSession = True
        Catch ex As Exception
            Call LogToFile("  Error: CreateIOSPlusSession - IOSPlus.IRESSSessionStart - " & ex.Message)
        End Try
    End Function

    Private Function CreateIOSPlusService(ByVal Session As Integer) As Boolean
        CreateIOSPlusService = False

        Try
            ' Create an IOS Plus service using the ServiceSessionStart method
            Dim sssRequest As IOSPlus.ServiceSessionStartInput = New IOSPlus.ServiceSessionStartInput

            ' Initialize the parameters of our IOS Plus service
            sssRequest.Parameters = New IOSPlus.ServiceSessionStartInputParameters
            sssRequest.Parameters.Service = "IOSPlus"
            sssRequest.Parameters.Server = gaSourcesList(Session).WSIOS
            sssRequest.Parameters.IRESSSessionKey = myIOSPlusSessionKey

            ' Call the ServiceSessionStart method to Create the IOS Plus service
            Dim sssResult As IOSPlus.ServiceSessionStartOutput = wsIOSPlus.ServiceSessionStart(sssRequest)

            If IsNothing(sssResult) Then
                Throw New Exception("Failed to start IOS Plus service")
            End If

            ' Obtain the IOS Plus service session key from the response of the ServiceSessionStart method
            myIOSPlusServiceSessionKey = sssResult.Result.DataRows(0).ServiceSessionKey

            CreateIOSPlusService = True
        Catch ex As Exception
            Call LogToFile("  Error: CreateIOSPlusService (Source" & CStr(Session) & ") - " & ex.Message)
        End Try
    End Function


    Private Sub EndIOSPlusService(ByVal Session As Integer)
        Try
            Dim sssRequest = New IOSPlus.ServiceSessionEndInput

            sssRequest.Header = New IOSPlus.ServiceSessionEndInputHeader()
            sssRequest.Header.ServiceSessionKey = myIOSPlusServiceSessionKey

            Dim sssResult As IOSPlus.ServiceSessionEndOutput = wsIOSPlus.ServiceSessionEnd(sssRequest)
        Catch ex As Exception
            Call LogToFile("  Error: EndIOSPlusService (Source" & CStr(Session) & ") - " & ex.Message)
        End Try
    End Sub

    Private Sub EndIOSPlusSession(ByVal Session As Integer)
        Try
            ' Create an IRESS Session End request
            Dim issRequest As IOSPlus.IRESSSessionEndInput = New IOSPlus.IRESSSessionEndInput

            issRequest.Header = New IOSPlus.IRESSSessionEndInputHeader
            issRequest.Header.SessionKey = myIOSPlusSessionKey

            Dim issResult As IOSPlus.IRESSSessionEndOutput = wsIOSPlus.IRESSSessionEnd(issRequest)
        Catch ex As Exception
            Call LogToFile("  Error: EndIOSPlusSession (Source" & CStr(Session) & ") - " & ex.Message)
        End Try
    End Sub

    Private Sub EndIRESSSession()
        Try
            ' Create an IRESS Session End request
            Dim issRequest As IRESS.IRESSSessionEndInput = New IRESS.IRESSSessionEndInput

            issRequest.Header = New IRESS.IRESSSessionEndInputHeader
            issRequest.Header.SessionKey = myIRESSSessionKey
            issRequest.Header.Timeout = glWSTimeout

            Dim issResult As IRESS.IRESSSessionEndOutput = wsIRESS.IRESSSessionEnd(issRequest)
        Catch ex As Exception
            Call LogToFile("  Error: EndIRESSSession - " & ex.Message)
        End Try
    End Sub
#End Region


End Class