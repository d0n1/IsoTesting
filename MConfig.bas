Attribute VB_Name = "MConfig"
Option Explicit
'Write ini file
Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long
'read ini file
Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long
'get computername
Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" _
                        (ByVal lpBuffer As String, nSize As Long) As Long
                        
'Delay by milliseconds
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





'Delay by seconds
Public Sub mWait(ByVal pSecond As Single)
    Dim sngStopTime As Single

    sngStopTime = Timer + pSecond

    Do While Timer <= sngStopTime
        DoEvents
    Loop

End Sub
'ini Write
Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
  
  Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
  INIWrite = (Err.Number = 0)
End Function
'ini read
Public Function INIRead(sSection As String, sKeyName As String, sINIFileName As String) As String
Dim sRet As String

  sRet = String(255, Chr(0))
  INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "NULL", sRet, Len(sRet), sINIFileName))
End Function

Public Function mMachineName() As String
    Dim sBuffer As String * 255
    If GetComputerName(sBuffer, 255&) <> 0 Then
        mMachineName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
End Function

