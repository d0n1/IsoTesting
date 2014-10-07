Attribute VB_Name = "OPMmodule"
Option Explicit
'****************************************************
'
'   This module includes Sub procedures for Optical Power Meter
'
'****************************************************


Public mSlotInfo() As Long          'Module information


Public Function OPM_Open_Session(resourceName As String, OPMhandle As Long, OPM_ID As String) As Long
'****************************************************
'Opens VISA session and reads out instruments ID.
'
'   input:  resourceName    VISA resource name for the OPM
'   output: OPMhandle       Session handle to control the OPM
'           OPM_ID          ID of the OPM
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
Dim status As Long
Dim lLength As Long
Dim strAns As String * 256

    ' open session, no verification, no reset
    status = hp816x_init(resourceName, 0, 0, OPMhandle)
    Call checkStatus(OPMhandle, status)
    
    If (status <> 0) Then
        ' connection failed
        strAns = "Connection Failed!"
    Else
        ' query instrument ID string
        lLength = 256
        status = hp816x_cmdString_Q(OPMhandle, "*IDN?", lLength, strAns)
        Call checkStatus(OPMhandle, status)
        
    End If
    
    OPM_ID = strAns
    OPM_Open_Session = status
    
End Function
Public Function OPM_Init_settings(OPMhandle As Long, pOPM_ID As String, pOPMSlots As Long, _
                                pOPMResponsData() As Double, pArrSize() As Long) As Long
'****************************************************
'Updated by Azen, 2012/12/24,read back pOPMResponsData() for dual channel power meter
'Initializes the OPM and reads out the response data of power sensors.
'The response data is used in data process
'   input:  OPMhandle
'           pOPM_ID
'   output: pOPMSlots     Number of OPM slots
'           pOPMResponsData()    Response data of power sensors（Channels x Num of data x 2）
'                               1チャネルあたりのデータは､波長と感度データが交互に並んだ1次元配列
'           pArrSize()       Number of array elements
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

Dim status As Long
'Dim ArrSize() As Long   '波長感度データのデータ数　slotごとの配列
Dim CSVSize As Long     'CSV形式の波長感度データのデータ数(値は未使用)
Dim MaxArrSize As Long  'ArrSize() 配列に含まれる数値の最大値
Dim slot As Long        'SLOT番号
Dim readBuf1() As Double
Dim readBuf2() As Double
Dim datnum As Long
'Dim channel As Long  'Azen 2012/12/23
Dim lLength As Long
Dim strReturn As String * 256
Dim strCmd As String
lLength = 256


'1.Configure trigger passthrogh
    status = hp816x_standardTriggerConfiguration(OPMhandle, hp816x_TRIG_PASSTHROUGH, 0, 0, 0)
    Call checkStatus(OPMhandle, status)

'2.Read wavelength response data
    'number of slots depends on mainframe
    
    If InStr(1, pOPM_ID, "8163", vbTextCompare) <> 0 Then
        pOPMSlots = 3 'real 2 slots
    ElseIf InStr(1, pOPM_ID, "8164", vbTextCompare) <> 0 Then
        pOPMSlots = 5 'real 4 slots
    ElseIf InStr(1, pOPM_ID, "8166", vbTextCompare) <> 0 Then
        pOPMSlots = 18 'real 17 slots
    ElseIf InStr(1, pOPM_ID, "7744", vbTextCompare) <> 0 Then
        pOPMSlots = 5 'real 4 slots
    ElseIf InStr(1, pOPM_ID, "7745", vbTextCompare) <> 0 Then
        pOPMSlots = 9 'real 8 slots
    End If
    
    'Read SlotInfo
    ReDim mSlotInfo(pOPMSlots)
    ReDim pArrSize(pOPMSlots, 2)
    status = hp816x_getSlotInformation_Q(OPMhandle, pOPMSlots, mSlotInfo(0))
    Call checkStatus(OPMhandle, status)
    

'===========================================Get ArraySize=======================================
'Add Optical Head check function
    For slot = 1 To pOPMSlots - 1
    
    Select Case mSlotInfo(slot)
    
            Case hp816x_SINGLE_SENSOR
            'For single sensor
            'Read number of data and set array size
                strCmd = "SLOT" & slot & ":HEAD:EMPTY?"
                status = hp816x_cmdString_Q(OPMhandle, strCmd, 256, strReturn)
                If strReturn Like "0*" Then
                    status = hp816x_getWlRespTblSizeEx(OPMhandle, slot, hp816x_CHAN_1, pArrSize(slot, hp816x_CHAN_1), CSVSize)
                    Call checkStatus(OPMhandle, status)
                    
                    If MaxArrSize < pArrSize(slot, hp816x_CHAN_1) Then
                        MaxArrSize = pArrSize(slot, hp816x_CHAN_1)
                    End If
                End If
            
            Case hp816x_DUAL_SENSOR
            'For dual sensor
            'Read number of data and set array size
            
            'Check sensor 1 if it is empty.
                strCmd = "SLOT" & slot & ":HEAD1:EMPTY?"
                status = hp816x_cmdString_Q(OPMhandle, strCmd, lLength, strReturn)
                Call checkStatus(OPMhandle, status)
                If strReturn Like "0*" Then
                    
                    status = hp816x_getWlRespTblSizeEx(OPMhandle, slot, hp816x_CHAN_1, pArrSize(slot, hp816x_CHAN_1), CSVSize)
                    Call checkStatus(OPMhandle, status)
                    
                    If MaxArrSize < pArrSize(slot, hp816x_CHAN_1) Then
                        MaxArrSize = pArrSize(slot, hp816x_CHAN_1)
                    End If

                    'Check sensor2 if it is empty?
                    strCmd = "SLOT" & slot & ":HEAD2:EMPTY?"
                    status = hp816x_cmdString_Q(OPMhandle, strCmd, lLength, strReturn)
                    Call checkStatus(OPMhandle, status)
                    
                    If strReturn Like "0*" Then
                        status = hp816x_getWlRespTblSizeEx(OPMhandle, slot, hp816x_CHAN_2, pArrSize(slot, hp816x_CHAN_2), CSVSize)
                        Call checkStatus(OPMhandle, status)
                        
                        If MaxArrSize < pArrSize(slot, hp816x_CHAN_2) Then
                            MaxArrSize = pArrSize(slot, hp816x_CHAN_2)
                        End If
                    End If
                End If
                    
            Case Else
                pArrSize(slot, 0) = 0
                pArrSize(slot, 1) = 0
                
      End Select
      
    Next slot
    
    ReDim pOPMResponsData(pOPMSlots, 2, 2 * MaxArrSize)
'=====================================================================================================

'***********************************Read OPM Responsive Data ******************************************
    'Read response data
    For slot = 1 To pOPMSlots - 1
        If pArrSize(slot, 0) Or pArrSize(slot, 1) <> 0 Then
            Select Case mSlotInfo(slot)
            
                Case hp816x_SINGLE_SENSOR
                
                    strCmd = "SLOT" & slot & ":HEAD:EMPTY?"
                    status = hp816x_cmdString_Q(OPMhandle, strCmd, 256, strReturn)
                    If strReturn Like "0*" Then
                        ReDim readBuf1(pArrSize(slot, hp816x_CHAN_1))
                        ReDim readBuf2(pArrSize(slot, hp816x_CHAN_1))
                        status = hp816x_readWlRespTableEx(OPMhandle, slot, hp816x_CHAN_1, readBuf1(0), readBuf2(0))
                        Call checkStatus(OPMhandle, status)
                        pArrSize(slot, hp816x_CHAN_1) = 2 * pArrSize(slot, hp816x_CHAN_1)
                        For datnum = 0 To pArrSize(slot, hp816x_CHAN_1) - 2 Step 2
                            pOPMResponsData(slot, hp816x_CHAN_1, datnum) = readBuf1(datnum / 2) * 1000000000# 'nm単位へ変換して格納
                            pOPMResponsData(slot, hp816x_CHAN_1, datnum + 1) = 10 * Log10(readBuf2(datnum / 2)) 'dB単位へ変換して格納
                        Next datnum
                    End If
                    
                Case hp816x_DUAL_SENSOR
                    strCmd = "SLOT" & slot & ":HEAD1:EMPTY?"
                    status = hp816x_cmdString_Q(OPMhandle, strCmd, lLength, strReturn)
                    Call checkStatus(OPMhandle, status)
                    If strReturn Like "0*" Then
                        ReDim readBuf1(pArrSize(slot, hp816x_CHAN_1))
                        ReDim readBuf2(pArrSize(slot, hp816x_CHAN_1))
                    
                        status = hp816x_readWlRespTableEx(OPMhandle, slot, hp816x_CHAN_1, readBuf1(0), readBuf2(0))
                        Call checkStatus(OPMhandle, status)
                        
                        pArrSize(slot, 0) = 2 * pArrSize(slot, 0)
                        
                        For datnum = 0 To pArrSize(slot, 0) - 2 Step 2
                            pOPMResponsData(slot, hp816x_CHAN_1, datnum) = readBuf1(datnum / 2) * 1000000000# 'nm単位へ変換して格納
                            pOPMResponsData(slot, hp816x_CHAN_1, datnum + 1) = 10 * Log10(readBuf2(datnum / 2)) 'dB単位へ変換して格納
                        Next datnum
                        
                        strCmd = "SLOT" & slot & ":HEAD2:EMPTY?"
                        status = hp816x_cmdString_Q(OPMhandle, strCmd, lLength, strReturn)
                        Call checkStatus(OPMhandle, status)
                        
                        If strReturn Like "0*" Then
                        'Read wavelength response data
                        ReDim readBuf1(pArrSize(slot, hp816x_CHAN_2))
                        ReDim readBuf2(pArrSize(slot, hp816x_CHAN_2))
                        
                        status = hp816x_readWlRespTableEx(OPMhandle, slot, hp816x_CHAN_2, readBuf1(0), readBuf2(0))
                        Call checkStatus(OPMhandle, status)
                        
                        pArrSize(slot, hp816x_CHAN_2) = 2 * pArrSize(slot, hp816x_CHAN_2)
                        
                        For datnum = 0 To pArrSize(slot, hp816x_CHAN_2) - 2 Step 2
                            pOPMResponsData(slot, hp816x_CHAN_2, datnum) = readBuf1(datnum / 2) * 1000000000# 'nm単位へ変換して格納
                            pOPMResponsData(slot, hp816x_CHAN_2, datnum + 1) = 10 * Log10(readBuf2(datnum / 2)) 'dB単位へ変換して格納
                        Next datnum
                        
                        End If
                        
                    End If
                    
            End Select
            
        End If
        
    Next slot
    
'***************************************************************************************************
    OPM_Init_settings = status
    
End Function
'Public Function OPM_Set_Wavlength(OPMhandle As Long, OPMSlots As Long, CentWav As Double) As Long
'****************************************************
'Sets the wavelength of the OPM
'
'input: OPMhandle
'       OPMSlots     Number of OPM slots
'       CentWav         Wavelength  [nm]
'output:None
'Function return:status (0:No error), (Non 0 value:Error code)
'****************************************************
'Dim status As Long
'Dim slot As Long
'Dim lLength As Long
'Dim strReturn As String * 256
'Dim strCmd As String
'lLength = 256
'
'    For slot = 1 To OPMSlots - 1
'       Select Case mSlotInfo(slot)
'            Case hp816x_SINGLE_SENSOR
'                '1.Set wavelength to center of scan range
'                status = hp816x_set_PWM_wavelength(OPMhandle, slot, hp816x_CHAN_1, CentWav * 0.000000001)
'                Call checkStatus(OPMhandle, status)
'
'            Case hp816x_DUAL_SENSOR
'                strCmd = "SLOT" & slot & ":HEAD1:EMPTY?"
'                status = hp816x_cmdString_Q(OPMhandle, strCmd, 256, strReturn)
'                If strReturn Like "0*" Then
'                    status = hp816x_set_PWM_wavelength(OPMhandle, slot, hp816x_CHAN_1, CentWav * 0.000000001)
'                    Call checkStatus(OPMhandle, status)
'
'                    strCmd = "SLOT" & slot & ":HEAD2:EMPTY?"
'                    status = hp816x_cmdString_Q(OPMhandle, strCmd, 256, strReturn)
'
'                    If strReturn Like "0*" Then
'                        status = hp816x_set_PWM_wavelength(OPMhandle, slot, hp816x_CHAN_2, CentWav * 0.000000001)
'                        Call checkStatus(OPMhandle, status)
'                    End If
'                Else
'                    MsgBox "There is NO Optical Head in Slot & " & slot & "! Please CHECK The Optical Head Connection!!"
'                End If
'            End Select
'    Next slot
Public Function OPM_Set_Wavlength(OPMhandle As Long, slot As Long, sensor As Long, CentWav As Double) As Long
    Dim status As Long
    Dim lLength As Long
    Dim strReturn As String * 256
    Dim strCmd As String
    lLength = 256
    strCmd = "SLOT" & slot & ":HEAD" & sensor & ":EMPTY?"
    status = hp816x_cmdString_Q(OPMhandle, strCmd, 256, strReturn)
    If strReturn Like "0*" Then
    
        status = hp816x_set_PWM_wavelength(OPMhandle, slot, sensor, CentWav * 0.000000001)
        Call checkStatus(OPMhandle, status)
    Else
    
        MsgBox "There is NO Optical Head in Slot: &" & slot & " Sensor: " & sensor & "! Please CHECK The Slot Settings!!"
        
    End If
    OPM_Set_Wavlength = status
    
End Function


Public Function OPM_Start_Logging(OPMhandle As Long, slot As Long, sensor As Long, powerRange As Double, _
                                sweepTime As Double, AvgTime As Double, NumofValues As Long) As Long
'****************************************************
'Starts data logging function.
'
'   input:  OPMhandle
'           OPMSlots Number of OPM channels
'           powerRange  Power range     [dBm]
'           SweepTime   Time duration for measurement   [sec]
'           AvgTime     Averaging Time  [sec]
'   output: NumofValues Number of logging data
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'
'
'****************************************************
Dim status As Long

Dim timeOut As Long


    NumofValues = CLng(sweepTime / 0.9 / AvgTime) + 1   '0.9 is margin for sweep speed

    
        If mSlotInfo(slot) = hp816x_SINGLE_SENSOR Or mSlotInfo(slot) = hp816x_DUAL_SENSOR Then
        '1.Set power range hold
            status = hp816x_set_PWM_powerRange(OPMhandle, slot, sensor, hp816x_PWM_RANGE_AUTO_OFF, powerRange)
            Call checkStatus(OPMhandle, status)
        '2.Set input trigger mode to CME(Complete measurement)
            status = hp816x_set_PWM_triggerConfiguration(OPMhandle, slot, hp816x_PWM_TRIGIN_CME, hp816x_PWM_TRIGOUT_NONE)
            Call checkStatus(OPMhandle, status)
        '3.Set Logging function parameters and start function
            status = hp816x_set_PWM_logging(OPMhandle, slot, sensor, AvgTime, NumofValues, timeOut)
            Call checkStatus(OPMhandle, status)
        End If

    
    OPM_Start_Logging = status

End Function


'Original:
'Public Function OPM_Read_LogData(OPMhandle As Long, OPMSlots As Long, NumofOPMdata As Long, PwrDataArr() As Double) As Long
Public Function OPM_Read_LogData(OPMhandle As Long, slot As Long, sensor As Long, NumofOPMdata As Long, pPwrDataArr() As Double) As Long
'****************************************************
'Reads out power data after logging is finished.
'
'input: OPMhandle
'output:pPwrDataArr() Array of power data [dBm]
'                   パワーメータチャネル数×取得データ数の2次元配列
'Function return:status (0:No error), (Non 0 value:Error code)
'****************************************************
Dim status As Long
Dim waitforCompletion As Integer
Dim LogStatus As Integer
Dim readBuff() As Double
Dim Cnt As Long
'Dim lngChannel As Long

'1.Wait for logging complete
    waitforCompletion = 1   '0: Not wait, 1:Wait
    ReDim readBuff(NumofOPMdata)

            'Read out log data(only one channel)
            status = hp816x_get_PWM_loggingResults_Q(OPMhandle, slot, sensor, waitforCompletion, hp816x_PU_DBM, LogStatus, readBuff(0))
            Call checkStatus(OPMhandle, status)
            
            For Cnt = 0 To NumofOPMdata
                pPwrDataArr(slot, sensor, Cnt) = readBuff(Cnt)
            Next Cnt

    OPM_Read_LogData = status

End Function
Public Function OPM_Stop_Logging(OPMhandle As Long, slot As Long, sensor As Long) As Long
'****************************************************
'Stops logging function that is started by OPM_Start_Logging
'
'input: OPMhandle
'output:None
'Function return:status (0:No error), (Non 0 value:Error code)
'****************************************************

Dim status As Long
    

        If mSlotInfo(slot) = hp816x_SINGLE_SENSOR Or mSlotInfo(slot) = hp816x_DUAL_SENSOR Then

        '1.Stop logging function
            status = hp816x_PWM_functionStop(OPMhandle, slot, sensor)
            Call checkStatus(OPMhandle, status)
        
        '2.Set trigger mode IGNORE
            status = hp816x_set_PWM_triggerConfiguration(OPMhandle, slot, hp816x_PWM_TRIGIN_IGN, hp816x_PWM_TRIGOUT_NONE)
            Call checkStatus(OPMhandle, status)
            
        '3.Set power range AUTO
            status = hp816x_set_PWM_powerRange(OPMhandle, slot, sensor, hp816x_PWM_RANGE_AUTO_ON, 0)
            Call checkStatus(OPMhandle, status)
        End If

    
    OPM_Stop_Logging = status
    
End Function

Public Function OPM_Close_Session(OPMhandle As Long) As Long
'****************************************************
'Closes VISA session that is opened by OPM_Open_Session
'
'input: OPMhandle
'output:None
'Function return:status (0:No error), (Non 0 value:Error code)
'****************************************************
' close session
Dim status As Long
    If OPMhandle <> 0 Then
        status = hp816x_close(OPMhandle)
        Call checkStatus(OPMhandle, status)
    End If
    
    OPM_Close_Session = status
    
End Function

Public Sub checkStatus(OPMhandle As Long, ByRef status As Long)
'****************************************************
'Checks return value of every function in OPMmodule.bas
'
'   input:  OPMhandle
'           status
'   output:
'
'****************************************************

Dim message As String * 256

    If status <> 0 Then

        If (status = hp816x_INSTR_ERROR_DETECTED) Then
            status = hp816x_error_query(OPMhandle, status, message)
        Else
            status = hp816x_error_message(OPMhandle, status, message)
        End If
        
        MsgBox message, vbOKOnly + vbExclamation, "VxiPnP Driver"
        Err.Raise status, , message
        
    End If
    
End Sub

