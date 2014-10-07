Attribute VB_Name = "STSctrlModule"
Option Explicit

Private TSLhandle As Long
Private OPMhandle As Long
Private SPUHandle As Long

Private mPMCalDataArr(3) As Double      'TSL power monitor calibration data
Private mTSLResponseData() As Double    'TSL power monitor response data
Private mOPMResponseData() As Double    'OPM power monitor response data
Private mTSLResArrSize As Long          'Array size TSL power monitor response data
Private mOPMResArrSize() As Long        'Array size OPM power monitor response data
Private mOPMSlots As Long               'Number of slots in the Mainframe

Private mSlot As Long
Private mSensor As Long

Private mStartWav As Double             'Sweep start wavelength
Private mStopWav As Double              'Sweep stop wavelength
Private mTrigStep As Double             'Trigger output step
Private mSwpSpeed As Double             'Sweep speed
Private mAvgTime As Double              'OPM averaging time
Private mInterpolNum As Long            'Number of interpolated data
Private mSweepTime As Double            'Sweep time
Private mTSLPowerRange As Long          'TSL power monitor range
Private mNumofDataPoints As Long
Private mSamplingRate As Double
Private mNumofSPUdata As Long
Private mSamplStep As Double
Private mCentWav As Double


Public Function STS_Initialize(TSLaddress As Integer, OPMaddress As Integer, SPUaddress As Integer, _
                                TSL_ID As String, OPM_ID As String) As Long
'*****************************************************************
'   Initialize instruments
'
'Opens VISA sessions and initialize each instrument.
'Program is begun by this function.
'SPU-100Addressへ入力する数値は、NI-MAX等で接続されている
'DAQmxのデバイス番号を確認して入力してください。
'（例:デバイス番号が"Dev1"の場合、"1"を入力）
'
'　 input:  TSLaddress  GPIB address of TSL-510
'           OPMaddress  GPIB address of optical powermeter
'           SPUaddress  Device number of SPU-100
'   output: TSL_ID      ID text of TSL-510
'           OPM_ID      ID text of optical powermeter
'
'*****************************************************************

Dim TSLResourceName As String
Dim OPMResourceName As String
Dim SPUResourceName As String
Dim MinWav As Double            'Minimum tunable wavelength
Dim MaxWav As Double            'Maximum tunable wavelength
Dim TrigStep As Double          'Trigger output step
Dim SwpSpeed As Double          'Sweep speed
Dim Retstat As Long
Dim TSLResponseData() As Double 'TSL power monitor response data
Dim OPMResponseData() As Double 'OPM power monitor response data
Dim TSLResArrSize As Long
Dim OPMResArrSize() As Long

'1.Open TSL-510
    TSLResourceName = "GPIB0::" & format(TSLaddress) & "::instr"
    Retstat = TSL_Open_Session(TSLResourceName, TSLhandle, TSL_ID)
    DoEvents
'2.Configure initial settings of TSL-510
    Retstat = m2TSL_Init_settings(TSLhandle, MinWav, MaxWav, mTSLResponseData(), mTSLResArrSize)
    DoEvents
'3.Open OPM
    OPMResourceName = "GPIB0::" & format(OPMaddress) & "::instr"
    Retstat = OPM_Open_Session(OPMResourceName, OPMhandle, OPM_ID)
    DoEvents
'4.Configure initial settings of OPM
    Retstat = OPM_Init_settings(OPMhandle, OPM_ID, mOPMSlots, mOPMResponseData(), mOPMResArrSize())
    DoEvents
'5.Open SPU-100
    SPUResourceName = "Dev" & format(SPUaddress)
    Retstat = SPU_Open_Session(SPUResourceName, SPUHandle)
    DoEvents
    

End Function
Public Function STS_Calibration(PMCalDataArr() As Double) As Long
      '*****************************************************************
      'Calibration function for TSL-510 power monitor
      '
      'Executes power calibration of TSL-510 power monitor.
      'It should be performed when the combination of TSL-510 and SPU-100 is changed,
      'or operation environment is changed. It is preferable to execute once a day.
      '
      '   output:             PMCalDataArr(3) キャリブレーション係数の配列（要素数4）
      '   Module variable:    mPMCalDataArr(3) These data is used in Data process.
      '                       General value is as follows.
      '                       Element (0):    12 to 15
      '                       Element (1):    2.5 to 3.0
      '                       Element (2):    0.5 to 0.6
      '                       Element (3):    0.1 to 0.12
      '
      '*****************************************************************
      Dim setWav As Double
      Dim Cnt As Long
      Dim Att As Double
      Dim AttArr(4) As Double
      Dim StartAtt As Double
      Dim AttStep As Double
      Dim PwrRange As Long
      Dim PrePwrRange As Long
      Dim Range As Long
      Dim Pwr As Double
      Dim LogPwr As Double
      Dim SamplingRate As Double
      Dim SampleNum As Long
      Dim TSLPwr() As Double
      Dim MonitPwr As Double
      Dim SPUdatArr() As Double
      Dim dataArr() As Double
      Dim ArrSize As Long
      Dim slope As Double
      Dim intercept As Double


10        On Error GoTo ErrorHandler

      '1.Set to Max Pwr wavelength
20        TSLErrChk TSL510_get_MaxPwrWav_Q(TSLhandle, setWav)
30        TSLErrChk TSL510_set_Wavelength(TSLhandle, setWav)
40        TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
      '2.Set TSL power Unit to mW
50        TSLErrChk TSL510_set_PwrUnits(TSLhandle, TSL510_PwrUnit_mW)
60        TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
      '3.Set TSL Attenuation Manual
70        TSLErrChk TSL510_set_AttnAuto(TSLhandle, TSL510_Attn_MANUAL)
80        TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
      '4.Set TSL power monitor range Auto
90        TSLErrChk TSL510_set_Pwr_Range_Auto(TSLhandle, TSL510_Range_AUTO, 0)
100       TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)

      '5. Config SPU
110       SamplingRate = 100000 'Hz
120       SampleNum = 2
130       Call SPU_Config_task(SPUHandle, SamplingRate, SampleNum)
          
      'Measurement Loop 1
140       StartAtt = 30 'dB
150       AttStep = 1 'dB

160       For Cnt = 0 To 30 'set Attenuator for 31 steps
170           Att = StartAtt - AttStep * Cnt 'attenuation value, 30, 29, .. ,0
              
              '6.Set TSL Attenuation and Wait for OPC
180           TSLErrChk TSL510_set_PwrAttn(TSLhandle, Att)
190           TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
200           Sleep (200)
              '7.Read TSL power monitor range
210           TSLErrChk TSL510_get_Pwr_Range_Q(TSLhandle, PwrRange)
220           TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
230           If Cnt <> 0 Then
240               If PwrRange <> PrePwrRange Then
                      'Record current attenuation
250                   AttArr(PrePwrRange) = Att + AttStep
                      
260               End If
270           End If
              
280           PrePwrRange = PwrRange  'Record current power range
290           DoEvents
300       Next Cnt
310       AttArr(0) = 0 'Attenuation value for max power range

      'Measurement Loop 2
320       ArrSize = 20
330       ReDim TSLPwr(ArrSize)
340       ReDim SPUdatArr(ArrSize)
350       For Range = 0 To 3
              '8.Set TSL power monitor range HOLD
360           TSLErrChk TSL510_set_Pwr_Range_Auto(TSLhandle, TSL510_Range_HOLD, Range)
370           Sleep (200)
380           TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
              
390           slope = 0
400           For Cnt = 0 To 19   'set Attenuator for 20 steps
410               Pwr = 1 - 0.05 * Cnt    'decrese Att = 1, 0.95, 0.9 ... 0.05
420               LogPwr = 10 * Log10(Pwr) 'LogPwr=0, -0.2, -0.5 ... -13
430               Att = AttArr(Range) - LogPwr
440               If Att > 30 Then
450                   Att = 30
460               ElseIf Att < 0 Then
470                   Att = 0
480               End If
                  
                  '9. Set TSL Attenuation
490               TSLErrChk TSL510_set_PwrAttn(TSLhandle, Att)
500               TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
510               Sleep (200)
                  
                  '10. Read TSL power
520               TSLErrChk TSL510_get_ActualPwr_Q(TSLhandle, MonitPwr)
                  'Record to array
530               TSLPwr(Cnt) = MonitPwr
                  
                  '11. Read SPU
540               Call SPU_Read_Data(SPUHandle, SampleNum, dataArr())
                  'Record to array
550               SPUdatArr(Cnt) = dataArr(3)
                  
560               slope = slope + TSLPwr(Cnt) / SPUdatArr(Cnt)
570               DoEvents
580           Next Cnt
              
              '12.Calculate scale factor
590           ArrSize = Cnt
600           slope = slope / ArrSize
              'Record to array
610           PMCalDataArr(Range) = slope
620           mPMCalDataArr(Range) = slope    '
              
              '13. Set TSL power monitor range AUTO
630           TSLErrChk TSL510_set_Pwr_Range_Auto(TSLhandle, TSL510_Range_AUTO, 0)
640           TSLErrChk TSL_510_Wait_for_OPC(TSLhandle, 30)
          
650       Next Range
          
660       STS_Calibration = 0
          
670       Exit Function
          
ErrorHandler:
680       STS_Calibration = Err.Number
690       Call mMakeErrorLog(Err.Number, Err.description, "STS_Calibration()", Erl)
700       Err.Clear
End Function


Public Function STS_PrepareLambdaScan(SourcePower As Double, _
                                    StartWavelength As Double, _
                                    StopWavelength As Double, _
                                    stepSize As Double, _
                                    sweepSpeed As Double, _
                                    NumofDataPoints As Long, _
                                    slot As Long, sensor As Long) As Long
                              
      '*****************************************************************
      '   Measurement function
      'Sets the parameters for wavelength scan
      '   input:  SourcePower         Output power of TSL-510     [dBm]
      '           StartWaveelngth     Start wavelength of scan    [nm]
      '           StopWavelength      Stop wavelength of sca      [nm]
      '           stepSize            Sampling step               [nm]
      '   output: NumofdataPoints     Number of samples
      '           NumofSlots          Number of OPM slots
      '*****************************************************************


      Dim CentWav As Double       'center wavelength of scan range
      Dim sweepTime As Double     'scan time
      Dim SamplingRate As Double  'sampling rate of SPU-100
      Dim SampleNum As Long
      Dim OPMDataArr() As Double  '
      Dim SPUDataArr() As Double
      Dim NumofSPUdata As Long    'Number of data acquired by SPU-100
      Dim dataArr() As Double
'      Dim sweepSpeed As Double
      Dim monitorVol As Double

10    On Error GoTo ErrorHandler
          
      '1.Configure TSL-510
          
          'Sweep Speed
'20        sweepSpeed = stepSize * 10000   'eg:step 0.001nm->speed 10nm/s, step 0.004nm->speed 40nm/s
          'Limited within 5 to 40nm/s
30        If sweepSpeed > 40 Then
40            sweepSpeed = 40
50        ElseIf sweepSpeed < 5 Then
60            sweepSpeed = 5
70        End If
          
80        WaitSecond 0.1
          
90        STS_chk_Status TSL_Set_Parameters(TSLhandle, SourcePower, StartWavelength, StopWavelength, stepSize, _
                          sweepSpeed, mAvgTime, mInterpolNum, mTSLPowerRange)
          
100       WaitSecond 0.1
          ' Read power monitor
110       SamplingRate = 100000 'Hz
120       SampleNum = 2
          
130       DAQmxErrChk SPU_Config_task(SPUHandle, SamplingRate, SampleNum)
140       DAQmxErrChk SPU_Read_Data(SPUHandle, SampleNum, dataArr())
          
150       monitorVol = dataArr(3)
          'If power monitor output is more then 1.5V, set to the upper range.(range number is lower)
          'パワーモニタが1.5V以上のときは1つ上のレンジ(レンジ番号は1マイナス)へ設定する
160       If monitorVol > 1.5 Then
170           mTSLPowerRange = mTSLPowerRange - 1
180       End If

      '2,Configure OPM
190       CentWav = (StartWavelength + StopWavelength) / 2                            'Center wavelength
          
200       STS_chk_Status OPM_Set_Wavlength(OPMhandle, slot, sensor, CentWav)


      '3.Configure SPU-100
210       sweepTime = Abs(StartWavelength - StopWavelength) / sweepSpeed              'Time span for sweep
220       SamplingRate = 100000 'Hz   SPU-100 sampling rate
230       NumofSPUdata = (sweepTime + 0.5) * SamplingRate   'SPU-100 sampling time, 0.5 seconds longer than sweep time
          
240       STS_chk_Status SPU_Config_task(SPUHandle, SamplingRate, NumofSPUdata)

      '4.function return
          
250       NumofDataPoints = CLng(Abs(StartWavelength - StopWavelength) / stepSize) + 1    'Number of sample
260       mSlot = slot
          mSensor = sensor
          'save to module level variables モジュールレベル変数へ保存
270       mStartWav = StartWavelength
280       mStopWav = StopWavelength
290       mSwpSpeed = sweepSpeed
300       mSweepTime = sweepTime
310       mSamplingRate = SamplingRate
320       mNumofSPUdata = NumofSPUdata
330       mSamplStep = stepSize
340       mNumofDataPoints = NumofDataPoints
350       mCentWav = CentWav
          
360       STS_PrepareLambdaScan = 0   'function return
          
370       Exit Function
          
ErrorHandler:
          
380       STS_PrepareLambdaScan = Err.Number
          

End Function


Public Function STS_getLambdaScanParameters_Q(StartWavelength As Double, _
                                    StopWavelength As Double, _
                                    averagingTime As Double, _
                                    sweepSpeed As Double) As Long
'****************************************************
'Reads out the parameters for wavelength scan
'   input:  None
'   output: StartWavelength Start wavelength of scan[nm]
'           StopWavelength  Stop wavelength of scan [nm]
'           AveragingTime   OPM　Averaging Time     [sec]
'           SweepSpeed      Scan Speed              [nm/s]
'****************************************************
    
    'Reads from module level variables モジュールレベル変数から読出し
    StartWavelength = mStartWav
    StopWavelength = mStopWav
    averagingTime = mAvgTime
    sweepSpeed = mSwpSpeed
    
End Function

Public Function m2STS_executeLambdaScan(NumofScan As Long, _
                                    InitPowerRange As Double, _
                                    rangeDecrement As Double, _
                                    wavelengthArray() As Double, _
                                    powerArray1() As Double, _
                                    powerArray2() As Double, _
                                    powerArray3() As Double, _
                                    powerArray4() As Double, _
                                    powerArray5() As Double, _
                                    powerArray6() As Double, _
                                    powerArray7() As Double, _
                                    powerArray8() As Double) As Long
      '****************************************************
      'Starts swept scan
      'Execute wavelength scan and returns measurement data
      '   input:  NumofScan       Number of scan. 1 to 3.
      '           InitPowerRange  Power range of OPM for the first scan. 10 to -30[dBm]
      '           rangeDecrement  Decrement step of power range. 10 to 30[dB]
      '   output: wavelengthArray Wavelength data array in nm.
      '           powerArray1()    Power data of Ch1 normalized by source power.[dB]
      '           powerArray2()    Power data of Ch2 normalized by source power.[dB]
      '           powerArray3()    Power data of Ch3 normalized by source power.[dB]
      '           powerArray4()    Power data of Ch4 normalized by source power.[dB]
      '           powerArray5()    Power data of Ch5 normalized by source power.[dB]
      '           powerArray6()    Power data of Ch6 normalized by source power.[dB]
      '           powerArray7()    Power data of Ch7 normalized by source power.[dB]
      '           powerArray8()    Power data of Ch8 normalized by source power.[dB]
      '****************************************************


      Dim scani As Long           'scan counter
      Dim OPMpowerRange As Double 'power range of OPM
      Dim OPMDataArr() As Double  'array for OPM data read
      Dim SPUDataArr() As Double  'array for SPU data read
      Dim InputArr() As Double    'array for OPM data process input
      Dim OutputArr() As Double   'array for OPM data process output
      Dim Cnt As Long
      Dim slot As Long
      Dim NumOfData As Long
      Dim OPMSlots As Long
      Dim sweepTime As Double
      Dim averageTime As Double
      Dim TSLPowerRange As Long
      Dim NumofSPUdata As Long
      Dim NumofOPMdata As Long
      Dim SamplingRate As Double
      Dim InterpolNum As Long
      Dim speed As Double
      Dim samplStep As Double
      Dim startWav As Double
      Dim Retstat As Long
      Dim TimeArr() As Double
      Dim WavArr() As Double
      Dim PwrMoniArr() As Double
      Dim WavResArr() As Double
      Dim WavResArrSize As Long
      Dim lngChannels As Long 'Channels of power meter slot
      Dim iChannel As Long
      Dim lngArrayIdx As Long


10    On Error GoTo ErrorHandle
20        NumOfData = mNumofDataPoints
30        OPMSlots = mOPMSlots
40        sweepTime = mSweepTime
50        averageTime = mAvgTime
60        TSLPowerRange = mTSLPowerRange
70        NumofSPUdata = mNumofSPUdata
80        SamplingRate = mSamplingRate
90        InterpolNum = mInterpolNum
100       speed = mSwpSpeed
110       samplStep = mSamplStep
120       startWav = mStartWav
          

130       ReDim wavelengthArray(NumOfData)
140       ReDim powerArray1(NumOfData, 1)
150       ReDim powerArray2(NumOfData, 1)
160       ReDim powerArray3(NumOfData, 1)
170       ReDim powerArray4(NumOfData, 1)
180       ReDim powerArray5(NumOfData, 1)
190       ReDim powerArray6(NumOfData, 1)
200       ReDim powerArray7(NumOfData, 1)
210       ReDim powerArray8(NumOfData, 1)

220       For scani = 0 To NumofScan - 1
          
230           WaitSecond 0.1
              
240           OPMpowerRange = CDbl(InitPowerRange - rangeDecrement * scani)
              
              '1,Set Measurement function mode
250           Retstat = OPM_Start_Logging(OPMhandle, mSlot, mSensor, OPMpowerRange, sweepTime, _
                                          averageTime, NumofOPMdata)
                                          
260           WaitSecond 0.2
270           Retstat = TSL_Standby_Logging(TSLhandle, TSLPowerRange)
              
              
280           WaitSecond 0.2
              '2.Start Measurement
290           Retstat = SPU_Start_task(SPUHandle)
300           Retstat = TSL_Start_Logging(TSLhandle)
              
              '3.Wait for Data acq complete and read out data from OPM
              
310           ReDim OPMDataArr(OPMSlots, 2, NumofOPMdata) '(slots, channels, Num of OPM data)
              
320           WaitSecond 0.2
330           Retstat = OPM_Read_LogData(OPMhandle, mSlot, mSensor, NumofOPMdata, OPMDataArr())
              
340           WaitSecond 0.2
              '4,Read out data from SPU-100
350           Retstat = SPU_Read_Data(SPUHandle, NumofSPUdata, SPUDataArr())
              
360           WaitSecond 0.2
              '5.Stop measurement
370           Retstat = SPU_Stop_Task(SPUHandle)
              
380           WaitSecond 0.1
390           Retstat = TSL_Stop_Logging(TSLhandle)
              
400           WaitSecond 0.1
410           Retstat = OPM_Stop_Logging(OPMhandle, mSlot, mSensor)
              
              '6.Process SPU data
420           ReDim TimeArr(NumOfData)
430           ReDim WavArr(NumOfData)
440           ReDim PwrMoniArr(NumOfData)
              
450           WaitSecond 0.2
460           Call STS_Data_Process(SPUDataArr(0), 2 * NumofSPUdata, InterpolNum, _
                              SamplingRate, speed, samplStep, startWav, _
                              TimeArr(0), NumOfData, WavArr(0), PwrMoniArr(0))
                              
               mPMCalDataArr(0) = 17.4564991
               mPMCalDataArr(1) = 3.597617361
               mPMCalDataArr(2) = 0.718810558
               mPMCalDataArr(3) = 0.144120686

              '7.Compensate data
470           WaitSecond 0.2
480           Call CompMonitorData(WavArr(0), PwrMoniArr(0), NumOfData, TSLPowerRange, _
                              mPMCalDataArr(0), 4, mTSLResponseData(0), mTSLResArrSize, PwrMoniArr(0))
              
              '8.Process OPM data
490           ReDim InputArr(NumofOPMdata)
500           ReDim OutputArr(NumOfData)
              
              
510
              

                  

                      'データ処理用に1slotごとの1次元配列へ格納
600                   For Cnt = 0 To NumofOPMdata - 1
610                       InputArr(Cnt) = OPMDataArr(mSlot, mSensor, Cnt)
620                   Next Cnt
                      
                      'OPMの波長感度補正データを1slotごとの1次元配列へ格納
630                   WavResArrSize = mOPMResArrSize(mSlot, mSensor)
640                   ReDim WavResArr(WavResArrSize)
650                   For Cnt = 0 To WavResArrSize - 1
660                       WavResArr(Cnt) = mOPMResponseData(mSlot, mSensor, Cnt)
670                   Next Cnt
                      
680                   Call OPM_Data_Process(InputArr(0), NumofOPMdata, averageTime, _
                                      TimeArr(0), WavArr(0), NumOfData, WavResArr(0), WavResArrSize, _
                                      mCentWav, OutputArr(0))

                      'パワーデータを出力配列へ格納
                      'Set array index
            
                Select Case mSlot
                Case 1
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray1(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray1(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 2
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray2(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray2(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 3
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray3(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray3(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 4
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray4(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray4(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 5
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray5(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray5(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 6
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray6(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray6(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 7
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray7(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray7(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                Case 8
                    For Cnt = 0 To NumOfData - 1
                    If OutputArr(Cnt) <= OPMpowerRange + 3 Then '(レンジ+3dB)以下のデータのみ保存する
                        If mSensor = hp816x_CHAN_1 Then
                            powerArray8(Cnt, 0) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        Else
                            powerArray8(Cnt, 1) = OutputArr(Cnt) - PwrMoniArr(Cnt)
                        End If
                    End If
                    Next Cnt
                End Select
750
                      
                      
1070      Next scani
          
1080      For Cnt = 0 To NumOfData - 1
1090          wavelengthArray(Cnt) = WavArr(Cnt)  '波長データを出力配列へ格納
1100      Next Cnt
          
1110      Exit Function
ErrorHandle:
1120      m2STS_executeLambdaScan = Err.Number
1130      Err.Clear

End Function

Public Sub STS_Close()
'    Closes the VISA sessions 各計測器との通信を開放します。プログラム終了時に1回実行します。
    Call SPU_Close_Session(SPUHandle)
    Call TSL_Close_Session(TSLhandle)
    Call OPM_Close_Session(OPMhandle)

End Sub
Public Sub STS_chk_Status(status As Long)
'
'   Utility function to handle errors
'
    If status <> 0 Then
        Err.Raise status
    End If


End Sub


Public Function Log10(x)
   '常用対数の計算
   Log10 = Log(x) / Log(10#)
End Function

