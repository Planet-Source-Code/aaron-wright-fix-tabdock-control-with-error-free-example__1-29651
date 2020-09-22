Attribute VB_Name = "modIni"
''''''''''''''''''''''''''''''''''''''''''''
'   White Blotter, Inc.                    '
'       Easy INI Module                    '
'                                          '
'   Name: EasyINI.bas                      '
'                                          '
'   This is available "as is" freeware.    '
'   No help will be provided for this      '
'   module.                                '
'                                          '
' http://www.geocities.com/whiteblotterinc '
''''''''''''''''''''''''''''''''''''''''''''

    #If Win16 Then
        Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
        Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal Default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
    #Else
        Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
    #End If

Function ReadINI(iSection, iKeyName, iFileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Function WriteINI(iSection As String, iKeyName As String, iNewString As String, iFileName) As Integer
    Dim r
    r = WritePrivateProfileString(iSection, iKeyName, iNewString, iFileName)
End Function
