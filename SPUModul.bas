Attribute VB_Name = "SPUModule"
Option Explicit
'****************************************************
'
'This module includes procedures for SPU-100
'
'****************************************************

Private strAns As String * 256
Private status As Long 'function return
Private data() As Double
Public taskIsRunning As Boolean

Public Function SPU_Open_Session(DevName As String, taskHandle As Long) As Long
'****************************************************
'Open the DAQmx session and create task
'プログラム実行時に必ず1回実行します。
'   input: DevName      Device number of SPU-100
'   output: taskHnadle    Session handle to control SPU-100
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'
'****************************************************

Dim physicalChannel As String
Dim minValue As Double
Dim maxValue As Double

    On Error GoTo ErrorHandler
    
    'Create the DAQmx task.
    DAQmxErrChk DAQmxCreateTask("", taskHandle)
    taskIsRunning = True

    physicalChannel = DevName & "/ai0:1"    'Use Analog input channel 0 and 1
    minValue = -5 'Volts
    maxValue = 5 'Volts

    'Add an analog input channel to the task.
    '   Terminal configration: Differential
    '   Input signal units: Volts
    DAQmxErrChk DAQmxCreateAIVoltageChan(taskHandle, physicalChannel, "", _
                    10106, minValue, maxValue, _
                    DAQmx_Val_VoltageUnits1_Volts, "")

    SPU_Open_Session = 0
Exit Function
    
ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Open_Session = Err.Number
    
End Function

Public Function SPU_Add_AI_Chan(taskHandle As Long, DevName As String) As Long
'****************************************************
'Configures analog input ports.
'プログラム実行時に必ず1回実行します。Execute once at the begining of program.
'
'   Input:  taskHnadle
'           DevName Device number of SPU-100
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'
'****************************************************

Dim physicalChannel As String
Dim minValue As Double
Dim maxValue As Double

    On Error GoTo ErrorHandler
       
    physicalChannel = DevName & "/ai0:1"    'Use Analog input channel 0 and 1
    minValue = -5 'Volts
    maxValue = 5 'Volts

    'Add an analog input channel to the task.
    '   Terminal configration: Differential
    '   Input signal units: Volts
    DAQmxErrChk DAQmxCreateAIVoltageChan(taskHandle, physicalChannel, "", _
                    DAQmx_Val_InputTermCfg_Diff, minValue, maxValue, _
                    DAQmx_Val_VoltageUnits1_Volts, "")
    SPU_Add_AI_Chan = 0
Exit Function
    
ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Add_AI_Chan = Err.Number
    
End Function


Public Function SPU_Config_task(taskHandle As Long, SamplingRate As Double, SamplesPerChannel As Long) As Long
'****************************************************
'Configures data sampling timing
'
'   input:  taskHnadle
'           SamplingRate        sampling rate  [Hz]
'           SamplesPerChannel   Number of data of 1 channel
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

    On Error GoTo ErrorHandler

    'Configure task for finite sample acquisition
    DAQmxErrChk DAQmxCfgSampClkTiming(taskHandle, "OnboardClock", _
                    SamplingRate, DAQmx_Val_Rising, _
                    DAQmx_Val_AcquisitionType_FiniteSamps, _
                    SamplesPerChannel)
    SPU_Config_task = 0
Exit Function
    
ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Config_task = Err.Number
    
End Function
Public Function SPU_Start_task(taskHandle As Long) As Long
'****************************************************
'Starts data sampling.
'
'   input:  taskHnadle
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'
'****************************************************
    On Error GoTo ErrorHandler
    
    DAQmxErrChk DAQmxStartTask(taskHandle)
    
    SPU_Start_task = 0
Exit Function
    
ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Start_task = Err.Number
    
End Function
Public Function SPU_Read_Data(taskHandle As Long, SamplesPerChannel As Long, data() As Double) As Long
'****************************************************
'Reads out data after sampling is finished
'
'   input:  taskHnadle
'           SamplesPerChannel   Number of data of 1 channel.
'   output: data()              Data array (1D)
'                               Number of elements of the array is SamplesPerChannel x number of channels.
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************

Dim sampsPerChanRead As Long
Dim numChannels As Long
Dim fillMode As DAQmxFillMode
Dim arraySizeInSamps As Long
Dim count As Long
Dim i As Long
Dim channel As Long
    
    On Error GoTo ErrorHandler
            
    fillMode = DAQmx_Val_GroupByScanNumber
    'DAQmx_Val_GroupByScanNumber - Use this fill mode when possible since this is faster
    ' Chan0Data0, Chan1Data0, Chan2Data0...,Chan0Data1, Chan1Data1, Chan2Data1...,Chan0Data2,Chan1Data2,Chan2Data2...
    
    'Set array size
    DAQmxErrChk DAQmxGetTaskNumChans(taskHandle, numChannels)
    arraySizeInSamps = SamplesPerChannel * numChannels
    ReDim data(arraySizeInSamps)
    
    'Read data
    DAQmxErrChk DAQmxReadAnalogF64(taskHandle, SamplesPerChannel, 10#, _
                    fillMode, data(0), arraySizeInSamps, sampsPerChanRead, ByVal 0&)
    
    ReDim DataByChan(numChannels, sampsPerChanRead)
    i = 0
    For count = 0 To sampsPerChanRead - 1
        For channel = 0 To numChannels - 1
            DataByChan(channel, count) = data(i)
            i = i + 1
        Next channel
    Next count
 
    SPU_Read_Data = 0
Exit Function
    
ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Read_Data = Err.Number
    
End Function
    
Public Function SPU_Stop_Task(taskHandle As Long) As Long
'****************************************************
'Stops the task that is started by SPU_Start_task.
'
'   input:  taskHnadle
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'
'****************************************************
    On Error GoTo ErrorHandler

    DAQmxErrChk DAQmxStopTask(taskHandle)
    taskIsRunning = False
    SPU_Stop_Task = 0

Exit Function

ErrorHandler:
    If taskIsRunning = True Then
        DAQmxStopTask taskHandle
        DAQmxClearTask taskHandle
        taskIsRunning = False
    End If
    MsgBox "Error: " & Err.Number & " " & Err.Description, , "SPU-100 Error"
    SPU_Stop_Task = Err.Number
    
End Function

Public Function SPU_Close_Session(taskHandle As Long) As Long
'****************************************************
'Closes task that is opend by SPU_Open_Session
'
'   input:  taskHnadle
'   output: None
'   Function return:
'           status (0:No error), (Non 0 value:Error code)
'****************************************************
    If taskIsRunning = True Then
        SPU_Stop_Task (taskHandle)
    End If
    If taskHandle <> 0 Then
        DAQmxErrChk DAQmxClearTask(taskHandle)
    End If
    SPU_Close_Session = 0
    
End Function

Public Sub DAQmxErrChk(errorCode As Long)
'****************************************************
'
'   Checks return value of every function in SPUModule.bas
'
'   input:  errorCode
'
'****************************************************

    Dim errorString As String
    Dim bufferSize As Long
    Dim status As Long
    If (errorCode < 0) Then
        ' Find out the error message length.
        bufferSize = DAQmxGetErrorString(errorCode, 0, 0)
        ' Allocate enough space in the string.
        errorString = String$(bufferSize, 0)
        ' Get the actual error message.
        status = DAQmxGetErrorString(errorCode, errorString, bufferSize)
        ' Trim it to the actual length, and display the message
        errorString = Left(errorString, InStr(errorString, Chr$(0)))
        Err.Raise errorCode, , errorString
    End If

    
End Sub



