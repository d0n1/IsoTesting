Attribute VB_Name = "TSLModule"
Option Explicit
'
'This module includes procedures for TSL-510
'

Private strAns As String * 256

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function TSL_Open_Session(resourceName As String, TSLhandle As Long, TSL_ID As String) As Long

'****************************************************
'Opens VISA session and reads out instruments ID.
'プログラム開始時に必ず1回実行します。
'   input:  resourceName    VISA resource name for TSL-510
'   output: TSL_ID          ID of TSL-510
'           TSLhandle       Session handle to control TSL-510.
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
Dim status As Long
Dim lLength As Long

    On Error GoTo ErrorHandler

    ' open session,
    lLength = 128
    TSLErrChk TSL510_Init(resourceName, 1, TSLhandle, strAns, lLength)
    TSL_ID = strAns
    TSL_Open_Session = 0
Exit Function
    
ErrorHandler:
    TSL_ID = "Connection Failed!"
    TSL_Open_Session = Err.Number
    
End Function
Public Function m2TSL_Init_settings(TSLhandle As Long, _
                                    pMinWav As Double, _
                                    pMaxWav As Double, _
                                    pWavRespDataArr() As Double, _
                                    pWavRespArrSize As Long) As Long

'****************************************************
'Initializes TSL-510 and reads out the response data of power monitor.
'The response data is used in data process.
'   input:  TSLhandle
'   Output: pMinWav      Minimum tunable wavelength [nm]
'           pMaxWav      Maximum tunable wavelength [nm]
'           pWavRespDataArr()    Response data of power monitor
'           pWavRespArrSize      Number of elements of pWavRespDataArr()
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
'    Dim pWavRespArrSize As Long  'Number of elements of Power monitor response dara array
    Dim MinDataWav As Double    'Minimum wavelength of Power monitor response dara array
    Dim MaxDataWav As Double    'Maximum wavelength of Power monitor response dara array
    Dim DataStep As Double      'Waveelngth step of Power monitor response dara array

    On Error GoTo ErrorHandler


    'Set command set to SCPI(1)
    TSLErrChk TSL510_set_CommSet(TSLhandle, TSL510_Command_SCPI)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

    'Reset
    TSLErrChk TSL510_Reset(TSLhandle)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

    'Set Input trigger disabled, and output trigger mode to STEP.
    TSLErrChk TSL510_set_TrigConfig(TSLhandle, TSL510_InpTrig_DISABLE, TSL510_OutpTrig_STEP)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

    'Set to Manual(ACC) mode
    TSLErrChk TSL510_set_AttnAuto(TSLhandle, TSL510_Attn_MANUAL)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

    'Read Max Min wavelength
    TSLErrChk TSL510_get_WavMinMax_Q(TSLhandle, pMinWav, pMaxWav)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
    'For debug, by Azen, 2012/12/24------------------------------
    Debug.Print "------TSL-510, From: m2TSL_Init_settings()------"
    Debug.Print "pMinWav= " & pMinWav
    Debug.Print "pMaxWav= " & pMaxWav
    '------------------------------------------------------------

    'Readout Power monitor response dara
    pWavRespArrSize = 600
    ReDim pWavRespDataArr(pWavRespArrSize)
    
    TSLErrChk TSL510_get_Pwr_Moni_Wav_Dep_Q(TSLhandle, pWavRespDataArr(0), pWavRespArrSize, _
    MinDataWav, MaxDataWav, DataStep)
    'For debug, by Azen, 2012/12/24------------------------------
    Debug.Print "pWavRespArrSize(default=600)= " & pWavRespArrSize
    Debug.Print "MinDataWav= " & MinDataWav
    Debug.Print "MaxDataWav= " & MaxDataWav
    Debug.Print "DataStep= " & DataStep
    '------------------------------------------------------------
    
    m2TSL_Init_settings = 0
Exit Function
    
ErrorHandler:
    m2TSL_Init_settings = Err.Number
    
End Function

Public Function TSL_Set_Parameters(TSLhandle As Long, _
                                SourcePower As Double, _
                                startWav As Double, _
                                stopWav As Double, _
                                samplStep As Double, _
                                speed As Double, _
                                AvgTime As Double, _
                                intpolNum As Long, _
                                powerRange As Long) As Long

'****************************************************
'Checks scan parameters, and sets to instrument
'
'   input:  TSLhandle
'           SourcePower     Output power of TSL-510 [dBm]
'           startWav        Start wavelength of scan[nm]
'           stopWav         Stop wavelength of scan [nm]
'           samplStep       Sampling step           [nm]
'           speed           Scan Speed              [nm/s]
'   output: AvgTime         Averaging Time of OPM   [sec]
'           intpolNum       Number for internal data process
'           powerRange      Power range of power monitor (0 to 3)
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

Dim TrigStep As Double
Dim targetPwr As Double
Dim resultPwr As Double
    
    On Error GoTo ErrorHandler


    '1.Chack sweep parameters
    TSLErrChk TSL510_chk_Sweep_Config(startWav, stopWav, samplStep, speed, intpolNum, TrigStep, AvgTime)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
    
    '2.Set sweep parameters
    TSLErrChk TSL510_set_Sweep_Config(TSLhandle, startWav, stopWav, TrigStep, speed)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

    '3.Set TSL-510 to start wavelength
    TSLErrChk TSL510_set_Wavelength(TSLhandle, startWav)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
    
    '4.Set TSL-510 source power (in ACC mode)
    TSLErrChk TSL510_set_Power_ACC(TSLhandle, SourcePower, resultPwr)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
'    Sleep (200)
    
    'Read Power monitor range
    TSLErrChk TSL510_get_Pwr_Range_Q(TSLhandle, powerRange)
    
    TSL_Set_Parameters = 0
    
Exit Function
    
ErrorHandler:
    TSL_Set_Parameters = Err.Number
    
End Function

Public Function TSL_Standby_Logging(TSLhandle As Long, powerRange As Long) As Long
'****************************************************
'Sets TSL-510 to wavelength scan mode.
'
'   input:  TSLhandle
'           powerRange  Power range of power monitor (0 to 3)
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

    On Error GoTo ErrorHandler
    
    'Set power monitor range HOLD
    TSLErrChk TSL510_set_Pwr_Range_Auto(TSLhandle, TSL510_Range_HOLD, powerRange)
    TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)


    'Start sweep
    TSLErrChk TSL510_set_Trig_Stdby(TSLhandle)

    TSL_Standby_Logging = 0
    
Exit Function
    
ErrorHandler:
    TSL_Standby_Logging = Err.Number
        
End Function
Public Function TSL_Start_Logging(TSLhandle As Long) As Long
'****************************************************
'Starts wavelength scan.
'
'   input:  TSLhandle
'   output:None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

    On Error GoTo ErrorHandler

    TSLErrChk TSL510_set_SoftTrigger(TSLhandle)
    
    TSL_Start_Logging = 0

Exit Function
    
ErrorHandler:
    
    TSL_Start_Logging = Err.Number
    
End Function


Public Function TSL_Stop_Logging(TSLhandle As Long) As Long
'****************************************************
'Sets TSL-510 to normal operation mode from wavelength scan mode.
'
'   input:  TSLhandle
'   output:None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
    
    On Error GoTo ErrorHandler
    
    '1.Set power monitor range AUTO
    TSLErrChk TSL510_set_Pwr_Range_Auto(TSLhandle, TSL510_Range_AUTO, 0)
    
    TSL_Stop_Logging = 0 'status
    
Exit Function
        
ErrorHandler:
    
    TSL_Stop_Logging = Err.Number
    
End Function

Public Function TSL_set_LaserState(TSLhandle As Long, State As Long) As Long
'****************************************************
'Sets LD current ON or OFF
'
'   input:  TSLhandle
'           State 0;OFF, 1;ON
'   output:
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

    On Error GoTo ErrorHandler

    TSLErrChk TSL510_set_LD_OnOff(TSLhandle, State)
    
    TSL_set_LaserState = 0

Exit Function
    
ErrorHandler:
    
    TSL_set_LaserState = Err.Number

End Function

    

Public Function TSL_Close_Session(TSLhandle As Long) As Long

'****************************************************
'Closes VISA session
'
'   input:  TSLhandle
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
' close session
    If TSLhandle <> 0 Then
        TSLErrChk TSL510_Close(TSLhandle)
    End If
    TSL_Close_Session = 0
    
End Function

Public Sub TSLErrChk(status)
Dim ErrCode As Long
Dim ErrorMsg As String * 256
Dim MsgLen As Long
'****************************************************
'TChecks return value of every function in TSLModule.bas.
'   input:  status　    Return value of function
'****************************************************

    If status <> 0 Then
        MsgLen = 256
        ErrCode = TSL510_Error_Message(status, ErrorMsg, MsgLen)
        MsgBox ErrorMsg, vbOKOnly + vbExclamation
        Err.Raise ErrCode, , ErrorMsg
    End If

End Sub



