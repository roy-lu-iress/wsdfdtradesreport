Module Globals

  Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
  Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Integer

  Public Function ReadINI(ByVal sINIFile As String, ByVal sSection As String, ByVal sKey As String, _
    ByVal sDefault As String) As String
        Dim sTemp As String = Space(2550)
    Dim iLength As Integer

        iLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 2550, sINIFile)
    Return sTemp.Substring(0, iLength)
  End Function

  Public Sub WriteINI(ByVal sINIFile As String, ByVal sSection As String, ByVal sKey As String, ByVal sValue As String)
    ' Write information to INI file
    WritePrivateProfileString(sSection, sKey, sValue, sINIFile)
  End Sub

End Module
