Attribute VB_Name = "SweptTestSystem"
Option Explicit

'************************************************
'   Constants for Command set function
'************************************************
Global Const TSL510_Command_SANTEC = 0
Global Const TSL510_Command_SCPI = 1

'************************************************
'   Constants for Trigger Configuration function
'************************************************
Global Const TSL510_InpTrig_DISABLE = 0
Global Const TSL510_InpTrig_ENABLE = 1

Global Const TSL510_OutpTrig_NONE = 0
Global Const TSL510_OutpTrig_STOP = 1
Global Const TSL510_OutpTrig_START = 2
Global Const TSL510_OutpTrig_STEP = 3

'************************************************
'   Constants for Attenuator function
'************************************************
Global Const TSL510_Attn_MANUAL = 0
Global Const TSL510_Attn_AUTO = 1

'************************************************
'   Constants for Sweep function
'************************************************
Global Const TSL510_SwpSpeed_5NMPS = 5
Global Const TSL510_SwpSpeed_10NMPS = 10
Global Const TSL510_SwpSpeed_20NMPS = 20
Global Const TSL510_SwpSpeed_40NMPS = 40

'************************************************
'   Constants for Power Monitor Range function
'************************************************
Global Const TSL510_Range_HOLD = 0
Global Const TSL510_Range_AUTO = 1

'************************************************
'   Constants for Power unit function
'************************************************
Global Const TSL510_PwrUnit_dBm = 0
Global Const TSL510_PwrUnit_mW = 1

'************************************************
'   Constants for Laset state
'************************************************
Global Const TSL510_Laser_OFF = 0
Global Const TSL510_Laser_ON = 1


'************************************************
'   Functions for general
'************************************************
Declare Function TSL510_Init Lib "SweptSystemShared.dll" _
    (ByVal VISAResourceName As String, _
    ByVal IDQuery As Integer, _
    ByRef InstrHandle As Long, _
    ByVal InstrumentName As String, _
    ByVal strLen As Long) As Long

Declare Function TSL510_Close Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long) As Long

Declare Function TSL510_set_CommSet Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal CommandSetName As Long) As Long

Declare Function TSL510_Reset Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long) As Long

Declare Function TSL510_get_Error_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal errorMessage As String, _
    ByRef MessageLength As Long) As Long

Declare Function TSL_510_Wait_for_OPC Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal timeOut As Long) As Long
Declare Function MW2dBm Lib "SweptSystemShared.dll" _
    (ByVal PowerMW As Double, _
    ByRef PowerDBm As Double) As Long

Declare Sub CompMonitorData Lib "SweptSystemShared.dll" _
    (ByRef WavArr As Double, _
    ByRef PwrMoniArr As Double, _
    ByVal Numofdata As Long, _
    ByVal TSLPRange As Long, _
    ByRef PMoniCal As Double, _
    ByVal NumOfPMoniCal As Long, _
    ByRef WPCAL As Double, _
    ByVal NumOfWPCal As Long, _
    ByRef PwrMoniArrOut As Double)
'
Declare Sub STS_Data_Process Lib "SweptSystemShared.dll" _
    (ByRef SPUDataArr As Double, _
    ByVal NumofSPUdata As Long, _
    ByVal InterpolNum As Long, _
    ByVal SamplingRate As Double, _
    ByVal sweepSpeed As Double, _
    ByVal samplingStep As Double, _
    ByVal StartWavelength As Double, _
    ByRef TimeArr As Double, _
    ByRef Numofdata As Long, _
    ByRef WavArr As Double, _
    ByRef PwrMoniArr As Double)

'
Declare Sub OPM_Data_Process Lib "SweptSystemShared.dll" _
   (ByRef OPMDatArrIn As Double, _
    ByVal OPMArrSize As Long, _
    ByVal averagingTime As Double, _
    ByRef TimeArr As Double, _
    ByRef WavArr As Double, _
    ByVal Numofdata As Long, _
    ByRef WavResponseDatArr As Double, _
    ByVal WavResArrSize As Long, _
    ByVal CenterWL As Double, _
    ByRef PwrDatArrOut As Double)
    
Declare Function TSL510_Error_Message Lib "SweptSystemShared.dll" _
   (ByVal ErrorCode As Long, _
    ByVal errorMessage As String, _
    ByVal MessageLength As Long) As Long

'************************************************
'   Functions for Wavelength
'************************************************
Declare Function TSL510_set_Wavelength Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal wavelength As Double) As Long

Declare Function TSL510_get_WavMinMax_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef minimumWavelength As Double, _
    ByRef maximumWavelength As Double) As Long

Declare Function TSL510_get_MaxPwrWav_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef MaxPowerWavelength As Double) As Long

'************************************************
'   Functions for Optical Power
'************************************************
Declare Function TSL510_set_Power_ACC Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal TargetPower As Double, _
    ByRef ResultPower As Double) As Long

Declare Function TSL510_set_PwrUnits Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal powerUnit As Integer) As Long

Declare Function TSL510_set_PwrAttn Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal attenuation As Double) As Long

Declare Function TSL510_get_ActualPwr_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef power As Double) As Long

Declare Function TSL510_set_AttnAuto Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal AttenuatorMode As Integer) As Long
    
Declare Function TSL510_set_Pwr_Range_Auto Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal rangeMode As Long, _
    ByVal Range As Long) As Long
    
Declare Function TSL510_get_Pwr_Range_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef Range As Long) As Long

Declare Function TSL510_get_Pwr_Moni_Wav_Dep_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef Responsivity As Double, _
    ByRef Numofdata As Long, _
    ByRef Min As Double, _
    ByRef Max As Double, _
    ByRef step As Double) As Long
    
Declare Function TSL510_set_LD_OnOff Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal State As Long) As Long

'************************************************
'   Functions for Wavelength scan
'************************************************
Declare Function TSL510_get_SwpStart_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef StartWavelength As Double) As Long
    
Declare Function TSL510_get_SwpStop_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef StopWavelength As Double) As Long
    
Declare Function TSL510_get_SwpSpeed_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef sweepSpeed As Double) As Long
    
Declare Function TSL510_get_Sweep_config_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef StartWavelength As Double, _
    ByRef StopWavelength As Double, _
    ByRef TrigStep As Double, _
    ByRef sweepSpeed As Double) As Long

Declare Function TSL510_set_SwpMode Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal sweepMode As Long) As Long
    
Declare Function TSL510_set_SwpStart Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal StartWavelength As Double) As Long
    
Declare Function TSL510_set_SwpStop Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal StopWavelength As Double) As Long
    
Declare Function TSL510_set_SwpSpeed Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal sweepSpeed As Double) As Long
    
Declare Function TSL510_set_Sweep_Config Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal StartWavelength As Double, _
    ByVal StopWavelength As Double, _
    ByVal TrigStep As Double, _
    ByVal sweepSpeed As Double) As Long

Declare Function TSL510_chk_Sweep_Config Lib "SweptSystemShared.dll" _
    (ByRef StartWavelength As Double, _
    ByRef StopWavelength As Double, _
    ByVal samplingStep As Double, _
    ByVal sweepSpeed As Double, _
    ByRef InterpolateNumber As Long, _
    ByRef TriggerStep As Double, _
    ByRef averagingTime As Double) As Long

'************************************************
'   Functions for Trigger
'************************************************
Declare Function TSL510_set_TrigConfig Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal InputTriggerMode As Integer, _
    ByVal OutputTriggerMode As Integer) As Long
    
Declare Function TSL510_set_TrigOutpStep Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByVal TriggerStep As Double) As Long
    
Declare Function TSL510_get_TrigOutpStep_Q Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long, _
    ByRef TrigStep As Double) As Long
    
Declare Function TSL510_set_Trig_Stdby Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long) As Long

Declare Function TSL510_set_SoftTrigger Lib "SweptSystemShared.dll" _
    (ByRef InstrHandle As Long) As Long

Declare Function LVDLLStatus Lib "SweptSystemShared.dll" _
    (ByRef errStr As String, ByVal errStrLen As Long, ByRef module As Long) As Long

