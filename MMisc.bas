Attribute VB_Name = "MMisc"
Option Explicit


'Constants for color
Public Const mButtonGray = &H8000000F
Public Const mMiddleGray = &HC0C0C0
Public Const mLightGreen = &H80FF80
Public Const mMiddleGreen = &HFF00&
Public Const mLightRed = &H8080FF
Public Const mLightYellow = &HFFFF&
Public Const mNormalYellow = &HFFFF&
Public Const mCBlack = &H0&
Public Const mCWhite = &HFFFFFF




Public mAccessSpecSettingPath As String
Public mAccessSpecSettingPath_T As String
Public mHelpWebPath As String
Public mPOSAMESPath As String
Public mPOSADataSheetPath As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9




Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp" _
                (ByVal PathName As String) As Long
                
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
                
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public StrMode As String






'Make Error log
Public Sub mMakeErrorLog(ByVal pErrNum As Variant, ByVal pErrDesc As Variant, ByVal pSub As Variant, ByVal pLineNum As Variant)
  Dim strDateTime As String
  Dim strDir As String
  Dim strPathFile As String
  Dim intFileNum As Integer
  Dim strErrMsg As String
  Dim strMsg As String
  
  'Check Dir \ErrorLog exist or not?
  strDir = App.Path & "\ErrorLog"
  If Dir(strDir, vbDirectory) = "" Then
      MkDir strDir
  End If
  
  'Set Error message
  strDateTime = Format(Date, "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss")
  strErrMsg = strDateTime & " --> Error: " & pErrNum & " , " & pErrDesc _
            & "--> From: " & pSub & " , @Line: " & pLineNum
  
  'Open file
  intFileNum = FreeFile()
  strPathFile = strDir & "\ErrorLog" & Trim(Format(Date, "yyyy")) _
              & Trim(Format(Date, "mm")) & ".txt"
  
  Open strPathFile For Append As #intFileNum
  
  'Write Error message
  Write #intFileNum, strErrMsg
  
  'Close file
  Close #intFileNum
  
  'Display Error message
  strMsg = "Error: " & pErrNum & " , " & pErrDesc _
         & vbCrLf & "From: " & pSub & " , @Line: " & pLineNum

  MsgBox strMsg, vbCritical
  
End Sub

'Public Function mCopyRecordset(ByRef pRsSource As ADODB.Recordset, ByRef pRsTarget As ADODB.Recordset) As Boolean
'          Dim tName As String 'field name
'          Dim tType 'field type
'          Dim tSize As Long 'field size
'          Dim i As Long
'
'10    On Error GoTo ErrorHandle
'
'20        mCopyRecordset = False
'
'          'Close traget recordset
'30        If pRsTarget.State = adStateOpen Then
'40            pRsTarget.Close
'50        End If
'
'          'Append fields------------------------------------------------------------
'60        For i = 0 To (pRsSource.Fields.count - 1)
'70            tName = pRsSource.Fields(i).Name
'80            tType = pRsSource.Fields(i).Type
'              'Because can't set value in adNumeric type,
'              'change to adDouble
'90            If tType = adNumeric Or tType = adVarNumeric Then
'100               tType = adDouble
'110           End If
'
'120           tSize = pRsSource.Fields(i).DefinedSize
'130           pRsTarget.Fields.Append tName, tType, tSize, adFldIsNullable
'140       Next
'
'          'Open recordset
'150       pRsTarget.Open
'          '-------------------------------------------------------------------------
'          'Copy data----------------------------------------------------------------
'160       If pRsSource.RecordCount <> 0 Then
'170           pRsSource.MoveFirst
'180           Do While Not pRsSource.EOF
'190               pRsTarget.AddNew
'200               For i = 0 To (pRsSource.Fields.count - 1)
'
'210                   pRsTarget.Fields(i).value = pRsSource.Fields(i).value
'
'220               Next
'230               pRsTarget.UpdateBatch adAffectCurrent
'240               pRsSource.MoveNext
'250           Loop
'260       End If
'          '-------------------------------------------------------------------------
'
'270       mCopyRecordset = True
'
'280       Exit Function
'
'ErrorHandle:
'290       Call mMakeErrorLog(Err.Number, Err.Description, "mCopyRecordset()", Erl)
'300       Err.Clear
'
'End Function

Public Sub mLoadFilePathSetup()
'Load file path setup
    Dim strPathFile As String
    Dim iFile As Long
    Dim iCount As Integer
    Dim strTemp As String

    strPathFile = App.Path & "\FilePathSetup.txt"
    'Check and create default file-----------------------
'    If Dir(strPathFile, vbArchive) = "" Then
        iFile = FreeFile
        Open strPathFile For Output As #iFile
            Write #iFile, "\\Filesrv\AFOP_PRC\Pub\Prc_Database\Software exe file\DY_SpecBasic2\DY_SpecBasic2.mdb" 'Access spec setting Jian Ti
            Write #iFile, "\\Filesrv\AFOP_PRC\Pub\Prc_Database\Software exe file\DY_T_SpecBasic2\DY_T_SpecBasic2.mdb" 'Access spec setting Fan Ti
            Write #iFile, "\\Filesrv\AFOP_PRC\Pub\Prc_Database\Software exe file\MFGI_POSA_MES\HelpWeb" 'System Help Web path
            Write #iFile, "\\Filesrv\AFOP_PRC\Pub\Prc_Database\Software exe file\MFGI_POSA_MES" 'POSA MES path
        Close #iFile
'    End If
    '----------------------------------------------------
    
    'Check data count------------------------------------------------------------------
    iFile = FreeFile
    Open strPathFile For Input As #iFile
        'Initialize count
        iCount = 0
        'Get data count
        Do Until EOF(iFile)
            iCount = iCount + 1
            Input #iFile, strTemp
        Loop
    Close #iFile
    '----------------------------------------------------------------------------------
    
    'If data count <5, then add data sheet-----------------------------------
'    If iCount < 3 Then
'        iFile = FreeFile
'        Open strPathFile For Append As #iFile
'            Write #iFile, "\\Filesrv\afop_prc\PUB\Prc_Database\Software exe file\MFGI_TAP-PD_MES\HelpWeb" 'System Help Web path
'        Close #iFile
'    End If
'    '----------------------------------------------------------------------------------
'
'    'If data count <4, then add Channel MES Path---------------------------------------
'    If iCount < 4 Then
'        iFile = FreeFile
'        Open strPathFile For Append As #iFile
'            Write #iFile, "\\Filesrv\afop_prc\PUB\Prc_Database\Software exe file\MFGI_TAP-PD_MES" 'Tap-PD MES path
'        Close #iFile
'    End If
'    '----------------------------------------------------------------------------------
    
     'If data count <5, then add Channel MES Path---------------------------------------
    If iCount < 5 Then
        iFile = FreeFile
        Open strPathFile For Append As #iFile
            Write #iFile, "\\Filesrv\afop_prc\PUB\Prc_Database\Software exe file\MFGI_POSA_MES" 'Excel Tap-PD Data Sheet
        Close #iFile
    End If
    '----------------------------------------------------------------------------------
    
    'Load file path setup--------------------------------
    iFile = FreeFile
    Open strPathFile For Input As #iFile
        Input #iFile, mAccessSpecSettingPath, mAccessSpecSettingPath_T
        Input #iFile, mHelpWebPath, mPOSAMESPath
        Input #iFile, mPOSADataSheetPath
    Close #iFile
    '----------------------------------------------------

End Sub

Public Function mDemo() As Boolean
'A demo program format
    Dim strSQL As String
  
On Error GoTo ErrorHandle

    mDemo = False




    mDemo = True
  
    Exit Function
  
ErrorHandle:
    Call mMakeErrorLog(Err.Number, Err.Description, "mDemo()", Erl)
    Err.Clear
  
End Function


'Public Function mExportRecordsetToTxt(ByRef pRsData As ADODB.Recordset, ByVal pPathFile As String) As Boolean
'      'Export data to a path\file
'          Dim strPath As String, strPathFile As String, strTestPath As String
'          Dim strMsg As String
'          Dim iFile As Long
'          Dim i As Long
'
'10    On Error GoTo ErrorHandle
'
'20        mExportRecordsetToTxt = False
'
'          'Check recordset-----------------------------------------------
'30        If pRsData.State = adStateClosed Then
'40            MsgBox "No data to export !!", vbCritical, "Error"
'50            Exit Function
'60        End If
'
'70        If pRsData.RecordCount < 1 Then
'80            MsgBox "No data to export !!", vbCritical, "Error"
'90            Exit Function
'100       End If
'          '--------------------------------------------------------------
'
'          'Set Path\FileName
'110       strPathFile = pPathFile
'
'          'Export data===================================================
'120       iFile = FreeFile
'130       Open strPathFile For Output As #iFile
'140           With pRsData
'                  'Write field names---------------------------
'150               For i = 0 To .Fields.count - 1
'160                   If i < .Fields.count - 1 Then
'170                       Write #iFile, .Fields(i).Name;
'180                   Else
'190                       Write #iFile, .Fields(i).Name
'200                   End If
'210               Next i
'                  '--------------------------------------------
'
'                  'Write data----------------------------------
'220               .MoveFirst
'230               Do While Not .EOF
'240                   For i = 0 To .Fields.count - 1
'250                       If i < .Fields.count - 1 Then
'260                           If Not IsNull(.Fields(i).value) Then
'270                               Write #iFile, .Fields(i).value;
'280                           Else
'290                               Write #iFile, "";
'300                           End If
'310                       Else
'320                           If Not IsNull(.Fields(i).value) Then
'330                               Write #iFile, .Fields(i).value
'340                           Else
'350                               Write #iFile, ""
'360                           End If
'370                       End If
'380                   Next i
'
'390                   .MoveNext
'400               Loop
'                  '--------------------------------------------
'
'410           End With
'420       Close #iFile
'          '==============================================================
'
'430       strMsg = "Export data...OK.--> To:" & vbCrLf & vbCrLf _
'                 & strPathFile
'440       MsgBox strMsg, vbInformation, "Export data"
'
'450       mExportRecordsetToTxt = True
'
'460       Exit Function
'
'ErrorHandle:
'470       Call mMakeErrorLog(Err.Number, Err.Description, "mExportRecordsetToTxt()", Erl)
'480       Err.Clear
'
'End Function

'call this function in KeyPress event method
Public Function mAutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
          Dim strFindThis As String, bContinueSearch As Boolean
          Dim lResult As Long, lStart As Long, lLength As Long
          
10        mAutoMatchCBBox = 0 ' block cbBox since we handle everything
20        bContinueSearch = True
30        lStart = cbBox.SelStart
40        lLength = cbBox.SelLength

50    On Error GoTo ErrHandle
              
60        If KeyAscii < 32 Then 'control char
70            bContinueSearch = False
80            cbBox.SelLength = 0 'select nothing since we will delete/enter
90            If KeyAscii = Asc(vbBack) Then 'take care BackSpace and Delete first
100               If lLength = 0 Then 'delete last char
110                   If Len(cbBox) > 0 Then ' in case user delete empty cbBox
120                       cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
130                   End If
140               Else 'leave unselected char(s) and delete rest of text
150                   cbBox.Text = Left(cbBox.Text, lStart)
160               End If
170               cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
180           ElseIf KeyAscii = vbKeyReturn Then  'user select this string
190               cbBox.SelStart = Len(cbBox)
200               lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
210               mAutoMatchCBBox = KeyAscii 'let caller a chance to handle "Enter"
220           End If
230       Else 'generate searching string
240           If lLength = 0 Then
250               strFindThis = cbBox.Text & Chr(KeyAscii) 'No selection, append it
260           Else
270               strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
280           End If
290       End If
          
300       If bContinueSearch Then 'need to search
310           Call VBComBoBoxDroppedDown(cbBox)  'open dropdown list
320           lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
330           If lResult = CB_ERR Then 'not found
340               cbBox.Text = strFindThis 'set cbBox as whatever it is
350               cbBox.SelLength = 0 'no selected char(s) since not found
360               cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
370           Else
                  'found string, highlight rest of string for user
380               cbBox.SelStart = Len(strFindThis)
390               cbBox.SelLength = Len(cbBox) - cbBox.SelStart
400           End If
410       End If
420       On Error GoTo 0
430       Exit Function
          
ErrHandle:
          'got problem, simply return whatever pass in
440       Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
450       Debug.Assert False
460       mAutoMatchCBBox = KeyAscii
470       On Error GoTo 0
          
End Function


'open dorpdown list
Private Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)

    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
    
End Sub



Public Sub mMakeLog2(ByVal pErrNum As Long, ByVal pErrDesc As String, ByVal pSub As String, ByVal pLineNum As Variant)
    Call mMakeLog("Error: " & pErrNum & " , " & pErrDesc _
                 & "-->From: " & pSub & " , @Line: " & pLineNum)
End Sub

Public Sub mMakeLog(pDescription As String)
    Dim strDir As String
    Dim strPathFile As String
    Dim intFileNum As Integer
    Dim strMsg As String

    'Check Dir \BackupLog exist or not?
    strDir = App.Path & "\Log"
    If Dir(strDir, vbDirectory) = "" Then
        MkDir strDir
    End If
    
    'Set message
    strMsg = Format(Date, "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") _
                & " --> " & pDescription
    
    'Open file
    intFileNum = FreeFile()
    strPathFile = strDir & "\Log" & Trim(Format(Date, "yyyy")) _
                & Trim(Format(Date, "mm")) & ".log"
    
    Open strPathFile For Append As #intFileNum
    
    'Write Error message
    Write #intFileNum, strMsg
    
    'Close file
    Close #intFileNum

    Exit Sub
End Sub


   
