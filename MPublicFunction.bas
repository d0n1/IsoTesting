Attribute VB_Name = "MPubFunction"
Option Explicit

Public mLogLevel As Integer

Public Function SaveData(ByRef strPoint() As String)
Dim fileHandle As Integer
Dim fileName As String
Dim i As Long

fileHandle = FreeFile
fileName = App.Path & "\testData" & Replace(CStr(Now), ":", "") & ".txt"

Open fileName For Output As #fileHandle
    For i = 0 To UBound(strPoint())
        Write #fileHandle, strPoint(i)
    Next i
Close #fileHandle

End Function




Public Sub mMakeServerLog(ByVal pClientIP As String, ByVal pLogDesc As String, ByVal pSub As String, ByVal pLogLevel As Integer)
'Make STS Server log
    Dim strDateTime As String
    Dim strDir As String
    Dim strPathFile As String
    Dim intFileNum As Integer
    Dim strLogMsg As String
    
    If pLogLevel <= mLogLevel Then
        'Check Dir \STSServerLog exist or not?
        strDir = App.Path & "\STSServerLog"
        If Dir(strDir, vbDirectory) = "" Then
            MkDir strDir
        End If
        
        'Set Error message
        strDateTime = Format(Date, "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss")
        strLogMsg = "[" & pClientIP & "], " & strDateTime & "-->" & pLogDesc _
                  & "--> From: " & pSub
        
        'Open file
        intFileNum = FreeFile()
        strPathFile = strDir & "\STSServerLog" & Trim(Format(Date, "yyyy")) _
                    & Trim(Format(Date, "mm")) & ".txt"
        
        Open strPathFile For Append As #intFileNum
        
        'Write Error message
        Write #intFileNum, strLogMsg
        
        'Close file
        Close #intFileNum
    
  End If
End Sub

Public Sub WaitSecond(ByVal pSecond As Single)
'Wait some second
    Dim sngStopTime As Single
    
    sngStopTime = Timer() + pSecond
    
    Do While Timer() <= sngStopTime
        DoEvents
    Loop

End Sub
