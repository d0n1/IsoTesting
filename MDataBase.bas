Attribute VB_Name = "MDataBase"
Option Explicit

Public mCnMES As New ADODB.Connection

Public Function mOpenCnMES() As Boolean
          Dim strSQL As String
          Dim strSite As String
        
10    On Error GoTo Errorhandle
        
20        mOpenCnMES = False
          
30        If mCnMES.State = adStateClosed Then
          
              'Set the properties of connection to MF2000 Oracle Server---------
              'ODBC Provider = "MSDASQL"
              'OLE DB Oracle Provider = "MSDAORA"
            
              'Set site area strSite: TC, DY, LocalTest
              'strSite = "DY"
              'strSite = "DY_OracleOLEDb"
              'strSite = "TC"
              'strSite = "AFOPDB"
              'strSite = "NewTest"
40            strSite = "DY_OracleOLEDb_45"
              'strSite = "DY_ODBC_2"
              
50            If strSite = "DY_OracleOLEDb_45" Then
                  '------New DY MES connection for Oracle OLE DB--------
60                mCnMES.CommandTimeout = 20
70                mCnMES.Provider = "OraOLEDB.Oracle"
80                mCnMES.Properties("Data Source").value = "DB45"
90                mCnMES.Properties("User ID").value = "mes"
100               mCnMES.Properties("Password").value = "mes"
                  
110               mCnMES.CursorLocation = adUseClient
                  '-----------------------------------------------------
              
120           ElseIf strSite = "DY_ODBC_2" Then
                  '------New DY MES connection for ODBC-----------------
130               mCnMES.CommandTimeout = 20
140               mCnMES.Provider = "MSDASQL" 'for ODBC
150               mCnMES.Properties("Data Source").value = "MESDY2" 'ODBC Name
160               mCnMES.Properties("User ID").value = "mes"
170               mCnMES.Properties("Password").value = "mes"
                  
180               mCnMES.CursorLocation = adUseClient
                  '-----------------------------------------------------
              
190           ElseIf strSite = "DY" Then
                  '------DY MF2000 connection -----------------------------
200               mCnMES.CommandTimeout = 20
210               mCnMES.Provider = "MSDASQL"
220               mCnMES.Properties("Data Source").value = "MESDY"
230               mCnMES.Properties("User ID").value = "mes"
240               mCnMES.Properties("Password").value = "mes"
                  
250               mCnMES.CursorLocation = adUseClient
                  '-----------------------------------------------------
            
260           ElseIf strSite = "DY_OracleOLEDb" Then
                  '------DY MF2000 connection for Oracle OLE DB---------
270               mCnMES.CommandTimeout = 20
280               mCnMES.Provider = "OraOLEDB.Oracle"
290               mCnMES.Properties("Data Source").value = "AFOPDG"
300               mCnMES.Properties("User ID").value = "mes"
310               mCnMES.Properties("Password").value = "mes"
                  
320               mCnMES.CursorLocation = adUseClient
                  '-----------------------------------------------------
            
330           ElseIf strSite = "TC" Then
                  '------TC MF2000 connection -----------------------------
340               mCnMES.CommandTimeout = 20
350               mCnMES.Provider = "MSDASQL"
360               mCnMES.Properties("Data Source").value = "MESTC"
370               mCnMES.Properties("User ID").value = "mestc"
380               mCnMES.Properties("Password").value = "mestc"
                  
390               mCnMES.CursorLocation = adUseClient
                  '-----------------------------------------------------
            
400           ElseIf strSite = "AFOPDB" Then
              
                  '------ Local connection (for testing)----------------
410               mCnMES.CommandTimeout = 20
420               mCnMES.Provider = "MSDASQL"
                  'mCnMES.Properties("Data Source").value = "ENGTEST"
                  'mCnMES.Properties("User ID").value = "ENGTEST"
                  'mCnMES.Properties("Password").value = "ENGTEST"
                  
430               mCnMES.Properties("Data Source").value = "AFOPDB"
440               mCnMES.Properties("User ID").value = "mes"
450               mCnMES.Properties("Password").value = "mes"
                  
460               mCnMES.CursorLocation = adUseClient
                  '------------------------------------------------------
470           ElseIf strSite = "NewTest" Then
              
                  '------ Local connection (for testing)----------------
480               mCnMES.CommandTimeout = 20
490               mCnMES.Provider = "MSDASQL"
                  'mCnMES.Properties("Data Source").value = "ENGTEST"
                  'mCnMES.Properties("User ID").value = "ENGTEST"
                  'mCnMES.Properties("Password").value = "ENGTEST"
                  
500               mCnMES.Properties("Data Source").value = "NewTest"
510               mCnMES.Properties("User ID").value = "mes"
520               mCnMES.Properties("Password").value = "mes"
                  
530               mCnMES.CursorLocation = adUseClient
                  '------------------------------------------------------
540           End If
            
550           mCnMES.Open
          
560       End If
          
570       mOpenCnMES = True
          
580       Exit Function
        
Errorhandle:
590       Call mMakeServerLog("mOpenCnMES", Erl & "-->" & Err.Number & "-->" & Err.Description, "mOpenCnMES()", 2)
600       Err.Clear

End Function


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

