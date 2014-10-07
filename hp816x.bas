Attribute VB_Name = "hp816x"
'  BASIC START
' ***************************************************************************
'   STANDARD SECTION
'   Constants and function prototypes for HP standard functions.
' ***************************************************************************
' ---------------------------------------------------------------------------
'  DEVELOPER: Remove what you don't need from this standard function
'               section, with the exception that VPP required functions
'               may not be removed.
'             Don't add to this section - add to the instrument specific
'               section below.
'             Don't change section - if you need to "change" the prototype
'               of a standard function, delete it from this section and
'               add a new function (named differently) in the instrument
'               specific section.
' ---------------------------------------------------------------------------
        ' *************************************************
        '   Standard constant error conditions returned
        '   by driver functions.
        '     HP Common Error numbers start at BFFC0D00
        '     The parameter errors extend the number of
        '       errors over the eight defined in VPP 3.4
        ' *************************************************
' Additional Parameter Errors
Global Const VI_ERROR_PARAMETER9 = &HBFFC0009
Global Const VI_ERROR_PARAMETER10 = &HBFFC000A
Global Const VI_ERROR_PARAMETER11 = &HBFFC000B
Global Const hp816x_INSTR_ERROR_NULL_PTR = &HBFFC0D02
Global Const hp816x_INSTR_ERROR_RESET_FAILED = &HBFFC0D03
Global Const hp816x_INSTR_ERROR_UNEXPECTED = &HBFFC0D04
Global Const hp816x_INSTR_ERROR_INV_SESSION = &HBFFC0D05
Global Const hp816x_INSTR_ERROR_LOOKUP = &HBFFC0D06
Global Const hp816x_INSTR_ERROR_DETECTED = &HBFFC0D07
Global Const hp816x_INSTR_NO_LAST_COMMA = &HBFFC0D08
Global Const hp816x_INSTR_NO_VALID_SLOT = &HBFFC0D09
Global Const hp816x_INSTR_NO_VALID_CHAN = &HBFFC0D0A
Global Const hp816x_DRIVER_LOCKED = &HBFFC0D0B
Global Const hp816x_MEM_ALLOC = &HBFFC0D0C
Global Const hp816x_INSTR_NO_VALID_SRC = &HBFFC0D0D
Global Const hp816x_INSTR_LOGGING_ACTIVE = &HBFFC0D0E
Global Const hp816x_INSTR_STAB_LOGG_ACTIVE = &HBFFC0D0F
Global Const hp816x_INSTR_MINMAX_ACTIVE = &HBFFC0D10
Global Const hp816x_NOVALID_SRCSEL = &HBFFC0D11
Global Const hp816x_INVALID_FUNC = &HBFFC0D12
Global Const hp816x_INVALID_IDN = &HBFFC0D13
Global Const hp816x_MODULE_NOT_PLUGGED = &HBFFC0D14
Global Const hp816x_NO_DUAL_SOURCE = &HBFFC0D15
Global Const hp816x_NO_VALID_MODULATION = &HBFFC0D16
Global Const hp816x_INVALID_PASSWORD = &HBFFC0D17
Global Const hp816x_MOD_SOURCE_NOT_SUPPORTED = &HBFFC0D18
Global Const hp816x_OPT_OUTPUT_NOT_SUPPORTED = &HBFFC0D19
Global Const hp816x_NO_POWERMETER = &HBFFC0D1A
Global Const hp816x_NO_LSPREPARE_CALL = &HBFFC0D1B
Global Const hp816x_PARAMETER_MISMATCH = &HBFFC0D1C
Global Const hp816x_SAME_SENSOR = &HBFFC0D1D
Global Const hp816x_NO_BUILTIN_SOURCE = &HBFFC0D1E
Global Const hp816x_ARRAYSIZE_TOSMALL = &HBFFC0D1F
Global Const hp816x_ZERODIVIDE = &HBFFC0D20
Global Const hp816x_NO_FRAME_REGISTERED = &HBFFC0D21
Global Const hp816x_NO_BL_TLS = &HBFFC0D22
Global Const hp816x_NO_DATA_ALLOC = &HBFFC0D23
Global Const hp816x_NO_TRIGGERS = &HBFFC0D24
Global Const hp816x_INVALID_VI = &HBFFC0D25
Global Const hp816x_FILEOPEN_ERROR = &HBFFC0D26
Global Const hp816x_TOO_FEW_TRIGGERS = &HBFFC0D27
Global Const hp816x_MODVERSION_ERROR = &HBFFC0D28
Global Const hp816x_LLOG_NUMBER = &HBFFC0D29
Global Const hp816x_CHAN_EXCLUDED = &HBFFC0D2A
Global Const hp816x_NOT_RELIABLE = &HBFFC0D2B
Global Const hp816x_INTERNAL_ERROR = &HBFFC0D2C
Global Const hp816x_SOFT_LOCKED = &HBFFC0D2D
Global Const hp816x_ERR_OFFSET_TBL = &HBFFC0D2E
Global Const hp816x_NO_BUILTIN_PM = &HBFFC0D2F
Global Const hp816x_INSTR_ERROR_PARAMETER9 = &HBFFC0D30
Global Const hp816x_INSTR_ERROR_PARAMETER10 = &HBFFC0D31
Global Const hp816x_INSTR_ERROR_PARAMETER11 = &HBFFC0D32
Global Const hp816x_INSTR_ERROR_PARAMETER12 = &HBFFC0D33
Global Const hp816x_INSTR_ERROR_PARAMETER13 = &HBFFC0D34
Global Const hp816x_INSTR_ERROR_PARAMETER14 = &HBFFC0D35
Global Const hp816x_INSTR_ERROR_PARAMETER15 = &HBFFC0D36
Global Const hp816x_INSTR_ERROR_PARAMETER16 = &HBFFC0D37
Global Const hp816x_INSTR_ERROR_PARAMETER17 = &HBFFC0D38
Global Const hp816x_INSTR_ERROR_PARAMETER18 = &HBFFC0D39
Global Const hp816x_WRONG_FCT = &HBFFC0D3A
Global Const hp816x_UNKNOWN_MODULE = &HBFFC0D3B
        ' *************************************************
        '   Constants used by system status functions
        '     These defines are bit numbers which define
        '     the operation and questionable registers.
        '     They are instrument specific.
        ' *************************************************
' ---------------------------------------------------------------------------
'  DEVELOPER: Modify these bit values to reflect the meanings of the
'             operation and questionable status registers for your
'               instrument.
' ---------------------------------------------------------------------------
Global Const hp816x_QUES_BIT0 = 1
Global Const hp816x_QUES_BIT1 = 2
Global Const hp816x_QUES_BIT2 = 4
Global Const hp816x_QUES_BIT3 = 8
Global Const hp816x_QUES_BIT4 = 16
Global Const hp816x_QUES_BIT5 = 32
Global Const hp816x_QUES_BIT6 = 64
Global Const hp816x_QUES_BIT7 = 128
Global Const hp816x_QUES_BIT8 = 256
Global Const hp816x_QUES_BIT9 = 512
Global Const hp816x_QUES_BIT10 = 1024
Global Const hp816x_QUES_BIT11 = 2048
Global Const hp816x_QUES_BIT12 = 4096
Global Const hp816x_QUES_BIT13 = 8192
Global Const hp816x_QUES_BIT14 = 16384
Global Const hp816x_QUES_BIT15 = 32768
Global Const hp816x_OPER_BIT0 = 1
Global Const hp816x_OPER_BIT1 = 2
Global Const hp816x_OPER_BIT2 = 4
Global Const hp816x_OPER_BIT3 = 8
Global Const hp816x_OPER_BIT4 = 16
Global Const hp816x_OPER_BIT5 = 32
Global Const hp816x_OPER_BIT6 = 64
Global Const hp816x_OPER_BIT7 = 128
Global Const hp816x_OPER_BIT8 = 256
Global Const hp816x_OPER_BIT9 = 512
Global Const hp816x_OPER_BIT10 = 1024
Global Const hp816x_OPER_BIT11 = 2048
Global Const hp816x_OPER_BIT12 = 4096
Global Const hp816x_OPER_BIT13 = 8192
Global Const hp816x_OPER_BIT14 = 16384
Global Const hp816x_OPER_BIT15 = 32768
        ' *************************************************
        '   Constants used by function hp816x_timeOut
        ' *************************************************
Global Const hp816x_TIMEOUT_MAX = 2147483647
Global Const hp816x_TIMEOUT_MIN = 0&
Global Const hp816x_CMDINT16ARR_Q_MIN = 1&
Global Const hp816x_CMDINT16ARR_Q_MAX = 2147483647
Global Const hp816x_CMDINT32ARR_Q_MIN = 1&
Global Const hp816x_CMDINT32ARR_Q_MAX = 2147483647
Global Const hp816x_CMDREAL64ARR_Q_MIN = 1&
Global Const hp816x_CMDREAL64ARR_Q_MAX = 2147483647
        ' *************************************************
        '   Required plug and play functions from VPP-3.1
        ' *************************************************
Declare Function hp816x_init Lib "hp816x_32.dll" (ByVal resourceName As String, ByVal IDQuery As Integer, ByVal resetDevice As Integer, IHandle As Long) As Long
Declare Function hp816x_getHandle Lib "hp816x_32.dll" (ByVal instrumentHandle As Long, IHandle As Long) As Long
Declare Function hp816x_close Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_reset Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_self_test Lib "hp816x_32.dll" (ByVal IHandle As Long, selfTestResult As Integer, ByVal selfTestMessage As String) As Long
Declare Function hp816x_mainframeSelftest Lib "hp816x_32.dll" (ByVal IHandle As Long, selfTestResult As Integer, ByVal selfTestMessage As String) As Long
Declare Function hp816x_error_query Lib "hp816x_32.dll" (ByVal IHandle As Long, errorCode As Long, ByVal errorMessage As String) As Long
Declare Function hp816x_error_message Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal statusCode As Long, ByVal message As String) As Long
Declare Function hp816x_revision_query Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal instrumentDriverRevision As String, ByVal firmwareRevision As String) As Long
Declare Function hp816x_revision_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal instrumentDriverRevision As String, ByVal firmwareRevision As String) As Long
        ' *************************************************
        '   HP standard utility functions
        ' *************************************************
Declare Function hp816x_timeOut Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal setTimeOut As Long) As Long
Declare Function hp816x_timeOut_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, timeOut As Long) As Long
Declare Function hp816x_errorQueryDetect Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal setErrorQueryDetect As Integer) As Long
Declare Function hp816x_errorQueryDetect_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, errorQueryDetect As Integer) As Long
Declare Function hp816x_dcl Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_opc_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, instrumentReady As Integer) As Long
Declare Function hp816x_WattToDBm Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal watt As Double, dBm As Double) As Long
Declare Function hp816x_dbmToWatt Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal dBm As Double, watt As Double) As Long
Declare Function hp816x_setDate Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal year As Long, ByVal month As Long, ByVal day As Long) As Long
Declare Function hp816x_getDate Lib "hp816x_32.dll" (ByVal IHandle As Long, year As Long, month As Long, day As Long) As Long
Declare Function hp816x_setTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal hour As Long, ByVal minute As Long, ByVal second As Long) As Long
Declare Function hp816x_getTime Lib "hp816x_32.dll" (ByVal IHandle As Long, hour As Long, minute As Long, second As Long) As Long
Declare Function hp816x_SystemError Lib "hp816x_32.dll" (ByVal IHandle As Long, errorNumber As Long, ByVal errorMessage As String) As Long
Declare Function hp816x_cls Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
        ' *************************************************
        '   HP standard status functions
        ' *************************************************
Declare Function hp816x_readStatusByte_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, statusByte As Integer) As Long
Declare Function hp816x_operEvent_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, operationEventRegister As Long) As Long
Declare Function hp816x_operCond_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, operationConditionRegister As Long) As Long
Declare Function hp816x_quesEvent_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, questionableEventRegister As Long) As Long
Declare Function hp816x_quesCond_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, questionableConditionRegister As Long) As Long
        ' *************************************************
        '   HP standard command passthrough functions
        ' *************************************************
Declare Function hp816x_cmd Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal sendStringCommand As String) As Long
Declare Function hp816x_cmdString_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryStringCommand As String, ByVal stringSize As Long, ByVal stringResult As String) As Long
Declare Function hp816x_cmdInt Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal sendIntegerCommand As String, ByVal sendInteger As Long) As Long
Declare Function hp816x_cmdInt16_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryI16Command As String, i16Result As Integer) As Long
Declare Function hp816x_cmdInt32_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryI32Command As String, i32Result As Long) As Long
Declare Function hp816x_cmdInt16Arr_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryI16ArrayCommand As String, ByVal i16ArraySize As Long, i16ArrayResult As Integer, i16ArrayCount As Long) As Long
Declare Function hp816x_cmdInt32Arr_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryI32ArrayCommand As String, ByVal i32ArraySize As Long, i32ArrayResult As Long, i32ArrayCount As Long) As Long
Declare Function hp816x_cmdReal Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal sendRealCommand As String, ByVal sendReal As Double) As Long
Declare Function hp816x_cmdReal64_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal queryRealCommand As String, realResult As Double) As Long
Declare Function hp816x_cmdReal64Arr_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal realArrayCommand As String, ByVal realArraySize As Long, realArrayResult As Double, realArrayCount As Long) As Long
'  End of HP standard declarations
' ---------------------------------------------------------------------------
' ***************************************************************************
'   INSTRUMENT SPECIFIC SECTION
'   Constants and function prototypes for instrument specific functions.
' ***************************************************************************
' ---------------------------------------------------------------------------
'  DEVELOPER: Add constants and function prototypes here.
'             As a metter of style, add the constant #define's first,
'               followed by function prototypes.
'             Remember that function prototypes must be consistent with
'               the driver's function panel prototypes.
' ---------------------------------------------------------------------------
        ' *************************************************
        '   Instrument specific constants
        ' *************************************************
' Maximum of supported instruments
Global Const hp816x_MAX_INSTR = 256
' bus type definitions
Global Const hp816x_INTF_GPIB = 0
Global Const hp816x_INTF_ASRL = 1
Global Const hp816x_INTF_ALL = 2
' list visa adresses
Global Const hp816x_SEL_816X = 1
Global Const hp816x_SEL_ALL = 0
'  channel and slot definitions
Global Const hp816x_SLOT_0 = 0
Global Const hp816x_SLOT_1 = 1
Global Const hp816x_SLOT_2 = 2
Global Const hp816x_SLOT_3 = 3
Global Const hp816x_SLOT_4 = 4
Global Const hp816x_SLOT_5 = 5
Global Const hp816x_SLOT_6 = 6
Global Const hp816x_SLOT_7 = 7
Global Const hp816x_SLOT_8 = 8
Global Const hp816x_SLOT_9 = 9
Global Const hp816x_SLOT_10 = 10
Global Const hp816x_SLOT_11 = 11
Global Const hp816x_SLOT_12 = 12
Global Const hp816x_SLOT_13 = 13
Global Const hp816x_SLOT_14 = 14
Global Const hp816x_SLOT_15 = 15
Global Const hp816x_SLOT_16 = 16
Global Const hp816x_SLOT_17 = 17
Global Const hp816x_CHAN_1 = 0
Global Const hp816x_CHAN_2 = 1
        ' **_***********************************************
        '   Constants for retrieving the type of modul plugged
        ' *************************************************
Global Const hp816x_UNDEF = 0&
Global Const hp816x_SINGLE_SENSOR = 1&
Global Const hp816x_DUAL_SENSOR = 2&
Global Const hp816x_FIXED_SINGLE_SOURCE = 3&
Global Const hp816x_FIXED_DUAL_SOURCE = 4&
Global Const hp816x_TUNABLE_SOURCE = 5&
Global Const hp816x_RETURN_LOSS = 6&
Global Const hp816x_ATTENUATOR = 7&
Global Const hp816x_SWITCH = 8&
        ' *************************************************
        '   Frame specific  constants
        ' *************************************************
Global Const hp816x_PWM_SLOT_MIN = 1
Global Const hp816x_MAX_SLOTS = 18
Global Const hp816x_NODE_A = 0
Global Const hp816x_NODE_B = 1
Global Const hp816x_TRIG_DISABLED = 0
Global Const hp816x_TRIG_DEFAULT = 1
Global Const hp816x_TRIG_PASSTHROUGH = 2
Global Const hp816x_TRIG_LOOPBACK = 3
Global Const hp816x_TRIG_CUSTOM = 4
        ' *************************************************
        '   Attenuator  specific  constants
        ' *************************************************
Global Const hp816x_SELECT_MIN = 0
Global Const hp816x_SELECT_MAX = 1
Global Const hp816x_SELECT_DEF = 2
Global Const hp816x_SELECT_MANUAL = 3
        ' *************************************************
        '   Powermeter  specific  constants
        ' *************************************************
' available channels
Global Const hp816x_PWM_CHAN_MIN = 0
Global Const hp816x_PWM_CHAN_MAX = 1
' power range also used for return loss
Global Const hp816x_PWM_RANGE_AUTO = 1
Global Const hp816x_PWM_RANGE_MANUAL = 0
' reference used internal or other modul
Global Const hp816x_PWM_TO_REF = 0
Global Const hp816x_PWM_TO_MOD = 1
Global Const hp816x_PWM_REF_ABSOLUTE = 0
Global Const hp816x_PWM_REF_RELATIV = 1
' trigger modes
Global Const hp816x_PWM_IMMEDIATE = 0
Global Const hp816x_PWM_CONTINUOUS = 1
' trigger force
Global Const hp816x_PWM_WAIT_TRIG = 0
Global Const hp816x_PWM_FORCE_TRIG = 1
' wave limits in nm
Global Const hp816x_PWM_WAVE_MIN = 0.0000003
Global Const hp816x_PWM_WAVE_MAX = 0.0000018
' averaging time limits in s
Global Const hp816x_PWM_ATIME_MIN = 0.000025
Global Const hp816x_PWM_ATIME_MAX = 900#
' correction limits in dbm
Global Const hp816x_PWM_CORR_MIN = -180#
Global Const hp816x_PWM_CORR_MAX = 200#
' reference limits in db, dBm
Global Const hp816x_PWM_REV_DB_MIN = -180#
Global Const hp816x_PWM_REV_DB_MAX = 200#
Global Const hp816x_PWM_REV_WATT_MIN = 1E-18
Global Const hp816x_PWM_REV_WATT_MAX = 1E+18
' power range in dBm
Global Const hp816x_PWM_RANGE_MIN = -110#
Global Const hp816x_PWM_RANGE_MAX = 30#
' power range mode
Global Const hp816x_PWM_RANGE_AUTO_OFF = 0
Global Const hp816x_PWM_RANGE_AUTO_ON = 1
' logging limits
Global Const hp816x_MAX_LOGNUMBER = 100001
' logging limits for old modules
Global Const hp816x_MAX_OLD_LOGNR = 4001
' includes 11 bytes header and 1 CR
Global Const hp816x_HEADER_SIZE = 12
Global Const hp816x_MAX_FUNCMODES = 4
Global Const hp816x_LOGSTAB_TTIME = 3600# * 24#
Global Const hp816x_LOGSTAB_DTIME = 3600# * 24#
' function types
Global Const hp816x_NONE = 0
Global Const hp816x_LOGGING = 1
Global Const hp816x_STABILITY = 2
Global Const hp816x_MINMAX = 3
' MINMAX modes
Global Const hp816x_MM_CONT = 0
Global Const hp816x_MM_WIN = 1
Global Const hp816x_MM_REFRESH = 2
' Trigger Configuration
' input Trigger
Global Const hp816x_PWM_TRIGIN_IGN = 0
' single measurement
Global Const hp816x_PWM_TRIGIN_SME = 1
' measurement completed
Global Const hp816x_PWM_TRIGIN_CME = 2
' output Trigger
' no  trigger
Global Const hp816x_PWM_TRIGOUT_NONE = 0
' end of averaging trigger
Global Const hp816x_PWM_TRIGOUT_AVG = 1
' begin of averaging trigger
Global Const hp816x_PWM_TRIGOUT_MEAS = 2
        ' *************************************************
        '   fixed laser source specific  constants
        ' *************************************************
' available channels for fixed laser
Global Const hp816x_FLS_CHAN_MIN = 1
Global Const hp816x_FLS_CHAN_MAX = 2
Global Const hp816x_FLS_SRC_SEL_MAX = 3
Global Const hp816x_FLS_ATTENUATION = 0
Global Const hp816x_FLS_POWER = 1
Global Const hp816x_CC_MODE_OFF = 0
Global Const hp816x_CC_MODE_ON = 1
Global Const hp816x_MODULATION_OFF = 0
Global Const hp816x_MODULATION_ON = 1
' defines for which laser sources a command is applied
Global Const hp816x_LOWER_SRC = 0
Global Const hp816x_UPPER_SRC = 1
Global Const hp816x_BOTH_SRC = 2
Global Const hp816x_EXTERN_SRC = 2
' attenuation limits
Global Const hp816x_ATT_MIN = 0#
Global Const hp816x_ATT_MAX = 10#
'  Modulation enabled
Global Const hp816x_MOD_DISABLED = 0
Global Const hp816x_MOD_ENABLED = 1
'  Modulation enabled
Global Const hp816x_SBS_CONTROL_DISABLED = 0
Global Const hp816x_SBS_CONTROL_ENABLED = 1
'  Modulation preselection
Global Const hp816x_AM_MIN = 0
Global Const hp816x_AM_MAX = 1
Global Const hp816x_AM_DEFAULT = 2
Global Const hp816x_AM_MANUAL = 3
'  Modulation sources
Global Const hp816x_AM_OFF = 0
Global Const hp816x_AM_INT = 1
Global Const hp816x_AM_CC = 2
Global Const hp816x_AM_LFCC = 3
Global Const hp816x_AM_BACKPLANE = 4
'  Attenuation mode
Global Const hp816x_ATT_MODE = 0
Global Const hp816x_ATT_POWER = 1
'  laser state
Global Const hp816x_LASER_OFF = 0
Global Const hp816x_LASER_ON = 1
'  power unit
Global Const hp816x_PU_DBM = 0
Global Const hp816x_PU_WATT = 1
' only for return loss
Global Const hp816x_PU_DB = 2
'  trigger state
Global Const hp816x_TRIG_DIS = 0
Global Const hp816x_TRIG_MOD = 1
        ' *************************************************
        '   return loss module specific  constants
        ' *************************************************
' trigger configuration
Global Const hp816x_RLM_TRIGIN_IGN = 0
Global Const hp816x_RLM_TRIGIN_SME = 1
Global Const hp816x_RLM_TRIGIN_CME = 2
Global Const hp816x_RLM_TRIGOUT_NONE = 0
Global Const hp816x_RLM_TRIGOUT_AVG = 1
'  calibration
Global Const hp816x_CAL_REFL = 0
Global Const hp816x_CAL_TERM = 1
Global Const hp816x_CAL_FACTORY = 2
' trigger modes
Global Const hp816x_RLM_IMMEDIATE = 0
Global Const hp816x_RLM_CONTINUOUS = 1
' power range
Global Const hp816x_RLM_RANGE_AUTO = 1
Global Const hp816x_RLM_RANGE_MANUAL = 0
' averaging time limits in s for return loss
Global Const hp816x_RLM_ATIME_MIN = 0.02
Global Const hp816x_RLM_ATIME_MAX = 10#
' attenuation limits
Global Const hp816x_RLM_ATT_MIN = 0#
Global Const hp816x_RLM_ATT_MAX = 3#
        ' *************************************************
        '   switch specific  constants
        ' *************************************************
' input port constants
Global Const hp816x_SWT_INP_A = 0
Global Const hp816x_SWT_INP_B = 1
' output port constants
Global Const hp816x_SWT_OUT_1 = 0
Global Const hp816x_SWT_OUT_2 = 1
Global Const hp816x_SWT_OUT_3 = 2
Global Const hp816x_SWT_OUT_4 = 3
Global Const hp816x_SWT_OUT_5 = 4
Global Const hp816x_SWT_OUT_6 = 5
Global Const hp816x_SWT_OUT_7 = 6
Global Const hp816x_SWT_OUT_8 = 7
        ' *************************************************
        '   tunable laser source specific  constants
        ' *************************************************
' wavelength units
Global Const hp816x_TLS_CHAN_MIN = 1
Global Const hp816x_TLS_CHAN_MAX = 2
'  Modulation output
Global Const hp816x_MOD_ALWAYS = 0
Global Const hp816x_MOD_LREADY = 1
' input type for wavelength and pow
Global Const hp816x_INP_MIN = 0
Global Const hp816x_INP_DEF = 1
Global Const hp816x_INP_MAX = 2
Global Const hp816x_INP_MAN = 3
' supported trigger in
Global Const hp816x_TLS_TRIGIN_IGN = 0
Global Const hp816x_TLS_TRIGIN_NEXTSTEP = 1
Global Const hp816x_TLS_TRIGIN_SWEEPSTART = 2
' supported trigger out
Global Const hp816x_TLS_TRIGOUT_DISABLED = 0
Global Const hp816x_TLS_TRIGOUT_MOD = 1
Global Const hp816x_TLS_TRIGOUT_STEPEND = 2
Global Const hp816x_TLS_TRIGOUT_SWSTART = 3
Global Const hp816x_TLS_TRIGOUT_SWEND = 4
Global Const hp816x_TLS_ATTENUATION = 0
Global Const hp816x_TLS_POWER = 1
' optical output modes
Global Const hp816x_HIGHPOW = 0
Global Const hp816x_LOWSSE = 1
Global Const hp816x_BHR = 2
Global Const hp816x_BLR = 3
' sweep repeat constants
Global Const hp816x_ONEWAY = 0
Global Const hp816x_TWOWAY = 1
' sweep parameter constants
Global Const hp816x_SWEEP_WAVE = 0
Global Const hp816x_SWEEP_POW = 1
Global Const hp816x_SWEEP_ATT = 2
' sweep state commands
Global Const hp816x_SW_CMD_STOP = 0
Global Const hp816x_SW_CMD_START = 1
Global Const hp816x_SW_CMD_PAUSE = 2
Global Const hp816x_SW_CMD_CONT = 3
' bnc output constants
Global Const hp816x_BNC_MOD = 0
Global Const hp816x_BNC_VPP = 1
Global Const hp816x_BNC_VPL = 2
' sweep mode constants
Global Const hp816x_SWEEP_STEP = 0
Global Const hp816x_SWEEP_MAN = 1
Global Const hp816x_SWEEP_CONT = 2
Global Const hp816x_SWEEP_FAST = 3
'  attenuation limits for tuneable lasersources
Global Const hp816x_TLS_ATT_MIN = 0
Global Const hp816x_TLS_ATT_MAX = 60
' sweep limits
Global Const hp816x_MIN_SWEEP_STEP = 0.0000000000001
Global Const hp816x_MIN_SWEEP_CYCLE = 1
Global Const hp816x_MAX_SWEEP_CYCLE = 999
Global Const hp816x_MIN_SWEEP_DWELL = 0.1
Global Const hp816x_MAX_SWEEP_DWELL = 1000000#
' sweep speeds
Global Const hp816x_SWEEP_SPEED_LOW = 0.0000000005
Global Const hp816x_SWEEP_SPEED_MEDIUM = 0.000000005
Global Const hp816x_SWEEP_SPEED_HIGH = 0.00000004
' rise time limits
Global Const hp816x_MIN_RISETIME = 0#
Global Const hp816x_MAX_RISETIME = 3#
' coherence control level range for DFB laser sources
Global Const hp816x_CCMIN_LEVEL = 1#
Global Const hp816x_CCMAX_LEVEL = 99.98
' SBS level range for Fireball laser sources
Global Const hp816x_SBSMIN_LEVEL = 0#
Global Const hp816x_SBSMAX_LEVEL = 100#
'  Modulation types
Global Const hp816x_MOD_INT = 0
Global Const hp816x_MOD_CC = 1
Global Const hp816x_MOD_AEXT = 2
Global Const hp816x_MOD_DEXT = 3
Global Const hp816x_MOD_VWLOCK = 4
Global Const hp816x_MOD_BACKPL = 5
' Lambda Scan defines
Global Const hp816x_NO_OF_SCANS_1 = 0
Global Const hp816x_NO_OF_SCANS_2 = 1
Global Const hp816x_NO_OF_SCANS_3 = 2
Global Const hp816x_PWM_CHANNEL_1 = 0
Global Const hp816x_PWM_CHANNEL_2 = 1
Global Const hp816x_PWM_CHANNEL_3 = 2
Global Const hp816x_PWM_CHANNEL_4 = 3
Global Const hp816x_PWM_CHANNEL_5 = 4
Global Const hp816x_PWM_CHANNEL_6 = 5
Global Const hp816x_PWM_CHANNEL_7 = 6
Global Const hp816x_PWM_CHANNEL_8 = 7
' sweep speeds
Global Const hp816x_SPEED_80NM = -1
Global Const hp816x_SPEED_40NM = 0
Global Const hp816x_SPEED_20NM = 1
Global Const hp816x_SPEED_10NM = 2
Global Const hp816x_SPEED_5NM = 3
Global Const hp816x_SPEED_05NM = 4
Global Const hp816x_SPEED_AUTO = 5
' baudrate definitions
Global Const hp816x_BAUD_9600 = 0
Global Const hp816x_BAUD_19200 = 1
Global Const hp816x_BAUD_38400 = 2
' number of channels a multiframe lambda scan can handle
Global Const MAX_MF_LAMBDASCAN_CHAN = 1000
' ****************************************************************************
' ***********************  Utility   Funktions ****************************
' ****************************************************************************
Declare Function hp816x_listVisa_Q Lib "hp816x_32.dll" (ByVal interf As Long, ByVal selection As Integer, noOfDef As Long, ByVal listofAdresses As String) As Long
Declare Function hp816x_getInstrumentId_Q Lib "hp816x_32.dll" (ByVal busAddress As String, ByVal IDNString As String) As Long
Declare Function hp816x_setBaudrate Lib "hp816x_32.dll" (ByVal interfaceIdentifier As String, ByVal baudrate As Long) As Long
' ****************************************************************************
' ***********************  Frame Specific Functions **************************
' ****************************************************************************
Declare Function hp816x_driverLogg Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal filename As String, ByVal logging As Integer, ByVal includeReplies As Integer) As Long
Declare Function hp816x_forceTransaction Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal forceTransaction As Integer) As Long
Declare Function hp816x_preset Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_moduleSelftest Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slottoTest As Long, result As Integer, ByVal selfTestMessage As String) As Long
Declare Function hp816x_lockUnlockInstument Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal softLock As Integer, ByVal password As String) As Long
Declare Function hp816x_getLockState Lib "hp816x_32.dll" (ByVal IHandle As Long, softLock As Integer, remoteInterlock As Integer) As Long
Declare Function hp816x_enableDisableDisplay Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal display As Integer) As Long
Declare Function hp816x_getSlotInformation_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal arraySize As Long, slotInfo As Long) As Long
Declare Function hp816x_getModuleStatus_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, statusSummary As Integer, moduleStatusArray As Long, maxMessageLength As Long) As Long
Declare Function hp816x_convertQuestionableStatus_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal questionableStatusInput As Long, ByVal message As String) As Long
Declare Function hp816x_standardTriggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal triggerConfiguration As Long, ByVal nodeAInputConfig As Long, ByVal nodeInputBConfig As Long, ByVal outputMatrixConfiguration As Long) As Long
Declare Function hp816x_standardTriggerConfiguration_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, triggerConfiguration As Long, nodeAInputConfig As Long, nodeInputBConfig As Long, outputMatrixConfiguration As Long) As Long
Declare Function hp816x_nodeInputConfiguration Lib "hp816x_32.dll" (ByVal connectionFunctionNodeA As Integer, ByVal BNCTriggerConnectorA As Integer, ByVal nodeBTriggerOutput As Integer, ByVal slotA0 As Integer, ByVal slotA1 As Integer, ByVal slotA2 As Integer, ByVal slotA3 As Integer, ByVal slotA4 As Integer, ByVal connectionfunctionNodeB As Integer, ByVal BNCTriggerConnectorB As Integer, ByVal nodeATriggerOutput As Integer, ByVal slotB0 As Integer, ByVal slotB1 As Integer, ByVal slotB2 As Integer, ByVal slotB3 As Integer, ByVal slotB4 As Integer, resultNodeA As Long, resultNodeB As Long) As Long
Declare Function hp816x_trigOutConfiguration Lib "hp816x_32.dll" (ByVal nodeswitchedtoBNCOutput As Integer, ByVal slot0 As Integer, ByVal slot1 As Integer, ByVal slot2 As Integer, ByVal slot3 As Integer, ByVal slot4 As Integer, result As Long) As Long
Declare Function hp816x_generateTrigger Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal triggerAt As Integer) As Long
' ****************************************************************************
' ***********************  Attenuator Specific Funktions **********************
' ****************************************************************************
Declare Function hp816x_set_ATT_attenuation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal attenuation As Double, ByVal waitforCompletion As Integer) As Long
Declare Function hp816x_get_ATT_attenuation_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_set_ATT_attenuationOffset Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal offset As Double) As Long
Declare Function hp816x_get_ATT_attenuationOffset_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_set_ATT_attenuatorSpeed Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal speed As Double) As Long
Declare Function hp816x_get_ATT_attenuatorSpeed_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_ATT_displayToOffset Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long) As Long
Declare Function hp816x_set_ATT_wavelength Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal wavelength As Double) As Long
Declare Function hp816x_get_ATT_wavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_set_ATT_powerUnit Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal unit As Integer) As Long
Declare Function hp816x_get_ATT_powerUnit_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, unit As Integer) As Long
Declare Function hp816x_set_ATT_absPowerMode Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal absolutePowerMode As Integer) As Long
Declare Function hp816x_get_ATT_absPowerMode_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, absolutePowerMode As Integer) As Long
Declare Function hp816x_set_ATT_power Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal powerControl As Integer, ByVal powerUnit As Integer, ByVal selection As Long, ByVal power As Double) As Long
Declare Function hp816x_get_ATT_power_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, powerControl As Integer, powerUnit As Integer, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_cp_ATT_refFromExtPM Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal slot As Long, ByVal channel As Long) As Long
Declare Function hp816x_set_ATT_powerReference Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal reference As Double) As Long
Declare Function hp816x_get_ATT_powerReference_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, def As Double, current As Double) As Long
Declare Function hp816x_set_ATT_shutterState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal shutterState As Integer) As Long
Declare Function hp816x_get_ATT_shutterState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, shutterState As Integer) As Long
Declare Function hp816x_set_ATT_shutterAtPowerOn Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal shutterState As Integer) As Long
Declare Function hp816x_get_ATT_shutterStateAtPowerOn_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, shutterState As Integer) As Long
Declare Function hp816x_get_ATT_operStatus Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, value As Long, ByVal conditionalInfo As String) As Long
Declare Function hp816x_read_ATT_outputPower Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, outputPower As Double) As Long
Declare Function hp816x_fetch_ATT_outputPower Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, outputPower As Double) As Long
Declare Function hp816x_set_ATT_wlOffsRefPowermeter Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal slot As Long, ByVal channel As Long) As Long
Declare Function hp816x_get_ATT_wlOffsRefPowermeter_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, slot As Long, channel As Long) As Long
Declare Function hp816x_set_ATT_wavelengthOffsetState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal offsetDependant As Integer) As Long
Declare Function hp816x_get_ATT_wavelengthOffsetState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, offsetDependancy As Integer) As Long
Declare Function hp816x_set_ATT_wavelengthOffset Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal wavelength As Double, ByVal offsetSource As Integer, ByVal offsetValue As Double) As Long
Declare Function hp816x_get_ATT_wavelengthOffsetIndex_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal tableIndex As Long, wavelength As Double, offset As Double) As Long
Declare Function hp816x_get_ATT_offsetFromWavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal wavelength As Double, offset As Double) As Long
Declare Function hp816x_delete_ATT_offsetTblEntries Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal noOfEntries As Integer, ByVal wavelengthOrIndex As Double) As Long
Declare Function hp816x_get_ATT_offsetTable_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, wavelengthArray As Double, offsetArray As Double) As Long
Declare Function hp816x_get_ATT_offsetTblSize_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, currentSize As Long, maximumSize As Long) As Long
Declare Function hp816x_set_ATT_powerOffset Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal selection As Long, ByVal offset As Double) As Long
Declare Function hp816x_get_ATT_powerOffset_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, minimum As Double, maximum As Double, defVal As Double, actual As Double) As Long
Declare Function hp816x_set_ATT_pwrOffsRefPM Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal slot As Long, ByVal channel As Long) As Long
Declare Function hp816x_set_ATT_offsByRefPM Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal slot As Long, ByVal channel As Long) As Long
Declare Function hp816x_set_ATT_controlLoopState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal controlLoop As Integer) As Long
Declare Function hp816x_get_ATT_controlLoopState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, controlLoopState As Integer) As Long
Declare Function hp816x_set_ATT_avTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal averagingTime As Double) As Long
Declare Function hp816x_get_ATT_avTime_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, averagingTime As Double) As Long
Declare Function hp816x_set_ATT_triggerConfig Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, ByVal triggerIn As Long) As Long
Declare Function hp816x_get_ATT_triggerConfig_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, triggerIn As Long) As Long
Declare Function hp816x_zero_ATT_powermeter Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long) As Long
Declare Function hp816x_get_ATT_zeroResult_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal ATTSlot As Long, lastZeroResult As Long) As Long
Declare Function hp816x_zero_ATT_all Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
' ****************************************************************************
' ***********************  Powermeter Specific Funktions **********************
' ****************************************************************************
Declare Function hp816x_PWM_slaveChannelCheck Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal slaveChannelCheck As Integer) As Long
Declare Function hp816x_set_PWM_parameters Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal rangeMode As Integer, ByVal powerUnit As Integer, ByVal internalTrigger As Integer, ByVal wavelength As Double, ByVal averagingTime As Double, ByVal powerRange As Double) As Long
Declare Function hp816x_get_PWM_parameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, rangeMode As Integer, powerUnit As Integer, internalTrigger As Integer, wavelength As Double, averagingTime As Double, powerRange As Double) As Long
Declare Function hp816x_set_PWM_referenceSource Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal measure As Long, ByVal referenceSource As Long, ByVal slotNumber As Long, ByVal channel As Long) As Long
Declare Function hp816x_get_PWM_referenceSource_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, measure As Long, referenceSource As Long, slotNumber As Long, channel As Long) As Long
Declare Function hp816x_set_PWM_internalTrigger Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal internalTrigger As Integer) As Long
Declare Function hp816x_set_PWM_averagingTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal averagingTime As Double) As Long
Declare Function hp816x_get_PWM_averagingTime_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, averagingTime As Double) As Long
Declare Function hp816x_set_PWM_wavelength Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal wavelength As Double) As Long
Declare Function hp816x_get_PWM_wavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, minWavelength As Double, maxWavelength As Double, currentWavelength As Double) As Long
Declare Function hp816x_set_PWM_calibration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal calibration As Double) As Long
Declare Function hp816x_get_PWM_calibration_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, calibration As Double) As Long
Declare Function hp816x_set_PWM_powerRange Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal rangeMode As Integer, ByVal powerRange As Double) As Long
Declare Function hp816x_get_PWM_powerRange_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, rangeMode As Integer, powerRange As Double) As Long
Declare Function hp816x_set_PWM_powerUnit Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal powerUnit As Long) As Long
Declare Function hp816x_get_PWM_powerUnit_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, powerUnit As Long) As Long
Declare Function hp816x_set_PWM_referenceValue Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal internalReference As Double, ByVal referenceChannel As Double) As Long
Declare Function hp816x_get_PWM_referenceValue_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, referenceMode As Long, internalReference As Double, referenceChannel As Double) As Long
Declare Function hp816x_set_PWM_triggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal triggerIn As Long, ByVal triggerOut As Long) As Long
Declare Function hp816x_get_PWM_triggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, trigIn As Long, trigOut As Long) As Long
Declare Function hp816x_PWM_displayToReference Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long) As Long
Declare Function hp816x_startp_PWM_internalTrigger Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long) As Long
Declare Function hp816x_PWM_readValue Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, measuredValue As Double) As Long
Declare Function hp816x_PWM_fetchValue Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, measuredValue As Double) As Long
Declare Function hp816x_PWM_readAll Lib "hp816x_32.dll" (ByVal IHandle As Long, numberofChannels As Long, slots As Long, channels As Long, values As Double) As Long
Declare Function hp816x_PWM_zeroing Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, zeroingResult As Long) As Long
Declare Function hp816x_PWM_zeroingAll Lib "hp816x_32.dll" (ByVal IHandle As Long, zeroingResult As Long) As Long
Declare Function hp816x_PWM_ignoreError Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal channelNumber As Long, ByVal ignoreError As Integer, ByVal instrumentErrorNumber As Long) As Long
Declare Function hp816x_set_PWM_logging Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal averagingTime As Double, ByVal numberofValues As Long, estimatedTimeout As Long) As Long
Declare Function hp816x_get_PWM_loggingResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal waitforCompletion As Integer, ByVal resultUnit As Integer, loggingStatus As Integer, loggingResult As Double) As Long
Declare Function hp816x_set_PWM_stability Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal averagingTime As Double, ByVal periodTime As Double, ByVal totalTime As Double, estimatedResults As Long) As Long
Declare Function hp816x_get_PWM_stabilityResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal waitforCompletion As Integer, ByVal resultUnit As Integer, stabilityStatus As Integer, stabilityResult As Double) As Long
Declare Function hp816x_set_PWM_minMax Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, ByVal minMaxMode As Long, ByVal dataPoints As Long, estimatedTimeout As Long) As Long
Declare Function hp816x_get_PWM_minMaxResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long, minimum As Double, maximum As Double, current As Double) As Long
Declare Function hp816x_PWM_functionStop Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMSlot As Long, ByVal channelNumber As Long) As Long
' ****************************************************************************
' ***********************  Fixed Laser Source Specific Funktions *************
' ****************************************************************************
Declare Function hp816x_set_FLS_parameters Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal wavelengthSource As Long, ByVal turnLaser As Integer, ByVal modulationLowerSource As Long, ByVal modulationUpperSource As Long, ByVal modulationFreqLowerSource As Double, ByVal modulationFreqUpperSource As Double, ByVal attenuationLowerSource As Double, ByVal attenuationUpperSource As Double) As Long
Declare Function hp816x_get_FLS_parameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, wavelengthSource As Long, turnLaser As Integer, modulationLowerSource As Long, modulationUpperSource As Long, modulationFreqLowerSource As Double, modulationFreqUpperSource As Double, attenuationLowerSource As Double, attenuationUpperSource As Double) As Long
Declare Function hp816x_set_FLS_modulation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal wavelengthSource As Long, ByVal modulationFrequency As Long, ByVal modulationSource As Long, ByVal manualFrequency As Double) As Long
Declare Function hp816x_get_FLS_modulationSettings_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal wavelengthSource As Long, modulationState As Integer, modSource As Long, minimumFrequency As Double, maximumFrequency As Double, defValFrequency As Double, currentFrequency As Double) As Long
Declare Function hp816x_set_FLS_triggerState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal outputTrigger As Integer) As Long
Declare Function hp816x_get_FLS_triggerState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, outputTrigger As Integer) As Long
Declare Function hp816x_set_FLS_laserSource Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal laserSource As Long) As Long
Declare Function hp816x_get_FLS_laserSource_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, outputWavelength As Long) As Long
Declare Function hp816x_get_FLS_wavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, wavelengthLowerSource As Double, wavelengthUpperSource As Double) As Long
Declare Function hp816x_set_FLS_attenuation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal attenuationLowerSource As Double, ByVal attenuationUpperSource As Double) As Long
Declare Function hp816x_get_FLS_attenuation_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, attenuationLowerSource As Double, attenuationUpperSource As Double) As Long
Declare Function hp816x_get_FLS_power Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal laserSource As Long, laserPower As Double) As Long
Declare Function hp816x_set_FLS_laserState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, ByVal laserState As Integer) As Long
Declare Function hp816x_get_FLS_laserState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal FLSSlot As Long, laserState As Integer) As Long
' ****************************************************************************
' ***********************  Switch Module Specific Funktions   ***********
' ****************************************************************************
Declare Function hp816x_get_SWT_type Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal SWTSlot As Long, switchType As Long, ByVal switchDescription As String) As Long
Declare Function hp816x_set_SWT_route Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal SWTSlot As Long, ByVal inpt As Long, ByVal outpt As Long) As Long
Declare Function hp816x_get_SWT_route Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal SWTSlot As Long, ByVal inpt As Long, outpt As Long) As Long
Declare Function hp816x_get_SWT_routeTable Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal SWTSlot As Long, ByVal RouteTable As String) As Long
' ****************************************************************************
' ***********************  Return Loss Module Specific Funktions   ***********
' ****************************************************************************
Declare Function hp816x_set_RLM_parameters Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal internalTrigger As Integer, ByVal wavelength As Double, ByVal averagingTime As Double, ByVal laserSource As Long, ByVal laserState As Integer) As Long
Declare Function hp816x_get_RLM_parameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, internalTrigger As Integer, wavelength As Double, averagingTime As Double, laserSource As Long, laserState As Integer) As Long
Declare Function hp816x_set_RLM_internalTrigger Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal internalTrigger As Integer) As Long
Declare Function hp816x_set_RLM_averagingTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal averagingTime As Double) As Long
Declare Function hp816x_get_RLM_averagingTime_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, averagingTime As Double) As Long
Declare Function hp816x_set_RLM_wavelength Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal wavelength As Double) As Long
Declare Function hp816x_get_RLM_wavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, minWavelength As Double, maxWavelength As Double, defWavelength As Double, currentWavelength As Double) As Long
Declare Function hp816x_calibrate_RLM Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal calibrate As Long) As Long
Declare Function hp816x_set_RLM_triggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal triggerIn As Long, ByVal triggerOut As Long) As Long
Declare Function hp816x_startp_RLM_internalTrigger Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long) As Long
Declare Function hp816x_RLM_readReturnLoss Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, returnLoss As Double) As Long
Declare Function hp816x_RLM_fetchReturnLoss Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, returnLoss As Double) As Long
Declare Function hp816x_RLM_readValue Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal monitorDiode As Integer, powerValue As Double) As Long
Declare Function hp816x_RLM_fetchValue Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal monitorDiode As Integer, powerValue As Double) As Long
Declare Function hp816x_RLM_zeroing Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, zeroingResult As Long) As Long
Declare Function hp816x_RLM_zeroingAll Lib "hp816x_32.dll" (ByVal IHandle As Long, summaryofZeroingAll As Long) As Long
Declare Function hp816x_set_RLM_powerRange Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal rangeMode As Integer, ByVal powerRange As Double, ByVal powerRangeSecondSensor As Double) As Long
Declare Function hp816x_get_RLM_powerRange_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, rangeMode As Integer, powerRange As Double, powerRangeSecondSensor As Double) As Long
Declare Function hp816x_set_RLM_rlReference Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, ByVal returnLossReference As Double) As Long
Declare Function hp816x_get_RLM_rlReference_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, returnLossReference As Double) As Long
Declare Function hp816x_set_RLM_FPDelta Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, ByVal frontPanelDelta As Double) As Long
Declare Function hp816x_get_RLM_FPDelta_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, frontPanelDelta As Double) As Long
Declare Function hp816x_calculate_RL Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal mref As Double, ByVal mpara As Double, ByVal pref As Double, ByVal ppara As Double, ByVal mdut As Double, ByVal pdut As Double, ByVal FPDelta As Double, returnLoss As Double) As Long
Declare Function hp816x_get_RLM_reflectanceValues_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, mref As Double, pref As Double) As Long
Declare Function hp816x_get_RLM_terminationValues_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, mpara As Double, ppara As Double) As Long
Declare Function hp816x_get_RLM_dutValues_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, mdut As Double, pdut As Double) As Long
Declare Function hp816x_get_RLM_srcWavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, wavelengthLowerSource As Double, wavelengthUpperSource As Double) As Long
Declare Function hp816x_set_RLM_laserSourceParameters Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, ByVal turnLaser As Integer) As Long
Declare Function hp816x_get_RLM_laserSourceParameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, laserSource As Long, laserState As Integer) As Long
Declare Function hp816x_set_RLM_laserState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserState As Integer) As Long
Declare Function hp816x_get_RLM_modulationState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, lowFrequencyControl As Integer) As Long
Declare Function hp816x_set_RLM_modulationState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal laserSource As Long, ByVal lowFrequencyControl As Integer) As Long
Declare Function hp816x_get_RLM_laserState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, laserState As Integer) As Long
Declare Function hp816x_enable_RLM_sweep Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal enableRLMLambdaSweep As Integer) As Long
Declare Function hp816x_set_RLM_logging Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal averagingTime As Double, ByVal dataPoints As Long, estimatedTimeout As Long) As Long
Declare Function hp816x_get_RLM_loggingResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal waitforCompletion As Integer, ByVal resultCalculatedas As Integer, ByVal monitorDiode As Integer, loggingStatus As Integer, loggingResult As Double) As Long
Declare Function hp816x_set_RLM_stability Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal averagingTime As Double, ByVal delayTime As Double, ByVal totalTime As Double, estimatedResults As Long) As Long
Declare Function hp816x_get_RLM_stabilityResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal waitforCompletion As Integer, ByVal resultUnit As Integer, stabilityStatus As Integer, stabilityResult As Double) As Long
Declare Function hp816x_set_RLM_minMax Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, ByVal minMaxMode As Long, ByVal dataPoints As Long, estimatedTime As Long) As Long
Declare Function hp816x_get_RLM_minMaxResults_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long, minimum As Double, maximum As Double, current As Double) As Long
Declare Function hp816x_RLM_functionStop Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal RLMSlot As Long) As Long
' ****************************************************************************
' ***********************  Tunable Laser Source Specific Funktions ***********
' ****************************************************************************
Declare Function hp816x_WaitForOPC Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal waitforOperationComplete As Integer) As Long
Declare Function hp816x_set_TLS_parameters Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal powerUnit As Long, ByVal opticalOutput As Long, ByVal turnLaser As Integer, ByVal power As Double, ByVal attenuation As Double, ByVal wavelength As Double) As Long
Declare Function hp816x_get_TLS_parameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, powerUnit As Long, laserState As Integer, opticalOutput As Long, power As Double, attenuation As Double, wavelength As Double) As Long
Declare Function hp816x_set_TLS_wavelength Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal wavelengthSelection As Long, ByVal wavelength As Double) As Long
Declare Function hp816x_get_TLS_wavelength_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, minimumWavelength As Double, defValWavelength As Double, maximumWavelength As Double, currentWavelength As Double) As Long
Declare Function hp816x_set_TLS_power Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal unit As Long, ByVal powerSelection As Long, ByVal manualPower As Double) As Long
Declare Function hp816x_get_TLS_power_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, powerUnits As Long, minimumPower As Double, defValPower As Double, maximumPower As Double, currentPower As Double) As Long
Declare Function hp816x_set_TLS_opticalOutput Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal setOpticalOutput As Long) As Long
Declare Function hp816x_get_TLS_opticalOutput_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, opticalOutput As Long) As Long
Declare Function hp816x_get_TLS_powerMaxInRange_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal startpofRange As Double, ByVal endofRange As Double, maximumPower As Double) As Long
Declare Function hp816x_set_TLS_laserState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal laserState As Integer) As Long
Declare Function hp816x_get_TLS_laserState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, laserState As Integer) As Long
Declare Function hp816x_set_TLS_laserRiseTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal laserRiseTime As Double) As Long
Declare Function hp816x_get_TLS_laserRiseTime Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, laserRiseTime As Double) As Long
Declare Function hp816x_set_TLS_triggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal triggerIn As Long, ByVal triggerOut As Long) As Long
Declare Function hp816x_get_TLS_triggerConfiguration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, triggerIn As Long, triggerOut As Long) As Long
Declare Function hp816x_set_TLS_attenuation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal powerMode As Integer, ByVal darkenLaser As Integer, ByVal attenuation As Double) As Long
Declare Function hp816x_get_TLS_attenuationSettings_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, powerMode As Integer, dark As Integer, attenuation As Double) As Long
Declare Function hp816x_set_TLS_dark Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal darkenLaser As Integer) As Long
Declare Function hp816x_get_TLS_darkState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, dark As Integer) As Long
Declare Function hp816x_get_TLS_temperatures Lib "hp816x_32.dll" (ByVal IHandle As Long, actualTemperature As Double, temperatureDifference As Double, temperatureLastZero As Double) As Long
Declare Function hp816x_get_TLS_temperaturesEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, actualTemperature As Double, temperatureDifference As Double, temperatureLastZero As Double) As Long
Declare Function hp816x_set_TLS_autoCalibration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal autocalibration As Integer) As Long
Declare Function hp816x_get_TLS_autoCalState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, autoCalibrationState As Integer) As Long
Declare Function hp816x_get_TLS_accClass Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, accuracyClass As Long) As Long
Declare Function hp816x_TLS_zeroing Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_TLS_zeroingAll Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_TLS_zeroingEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long) As Long
Declare Function hp816x_displayToLambdaZero Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long) As Long
Declare Function hp816x_get_TLS_lambdaZero_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, lambda0 As Double) As Long
Declare Function hp816x_set_TLS_frequencyOffset Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal offset As Double) As Long
Declare Function hp816x_get_TLS_frequencyOffset_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, offset As Double) As Long
Declare Function hp816x_TLS_configureBNC Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal BNCOutput As Long) As Long
Declare Function hp816x_get_TLS_BNC_config_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, BNCOutput As Long) As Long
Declare Function hp816x_set_TLS_sweep Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal sweepMode As Long, ByVal repeatMode As Long, ByVal cycles As Long, ByVal dwellTime As Double, ByVal startpWavelength As Double, ByVal stoppWavelength As Double, ByVal stepSize As Double, ByVal sweepSpeed As Double) As Long
Declare Function hp816x_TLS_sweepControl Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal action As Long) As Long
Declare Function hp816x_get_TLS_sweepState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, sweepState As Long) As Long
Declare Function hp816x_TLS_sweepNextStep Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long) As Long
Declare Function hp816x_TLS_sweepPreviousStep Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long) As Long
Declare Function hp816x_TLS_sweepWait Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long) As Long
Declare Function hp816x_set_TLS_modulation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal modulationSource As Long, ByVal modulationOutput As Integer, ByVal modulation As Integer, ByVal modulationFrequency As Double) As Long
Declare Function hp816x_get_TLS_modulationSettings_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, modulationSource As Long, modulationOutput As Integer, modulationState As Integer, frequency As Double) As Long
Declare Function hp816x_set_TLS_SBS_control Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal mod_state As Integer, ByVal modulationFrequency As Double) As Long
Declare Function hp816x_get_TLS_SBS_control_q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, mod_state As Integer, frequency As Double) As Long
Declare Function hp816x_set_TLS_ccLevel Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal CCLevel As Double) As Long
Declare Function hp816x_get_TLS_ccLevel_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, CCLevel As Double) As Long
Declare Function hp816x_set_TLS_SBSLevel Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal SBSLevel As Double) As Long
Declare Function hp816x_get_TLS_SBSLevel_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, SBSLevel As Double) As Long
Declare Function hp816x_set_TLS_lambdaLoggingState Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal lambdaLoggingState As Integer) As Long
Declare Function hp816x_set_TLS_lambdaLoggingStateEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal lambdaLoggingState As Integer) As Long
Declare Function hp816x_get_TLS_lambdaLoggingState_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, lambdaLoggingState As Integer) As Long
Declare Function hp816x_get_TLS_lambdaLoggingStateEx_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, lambdaLoggingState As Integer) As Long
Declare Function hp816x_get_TLS_wavelengthData_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal Array_Size As Long, wavelengthData As Double) As Long
Declare Function hp816x_get_TLS_wavelengthDataEx_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal Array_Size As Long, wavelengthData As Double) As Long
Declare Function hp816x_get_TLS_wavelengthPoints_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, numberofWavelengthPoints As Long) As Long
Declare Function hp816x_get_TLS_wavelengthPointsEx_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, numberofWavelengthPoints As Long) As Long
Declare Function hp816x_get_TLS_powerPoints_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, numberofPowerValues As Long) As Long
Declare Function hp816x_get_TLS_powerData_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal TLSSlot As Long, ByVal arraySize As Long, wavelengthData As Double, powerData As Double) As Long
' ****************************************************************************
' ***********************  Applications               ***********
' ****************************************************************************
Declare Function hp816x_enableHighSweepSpeed Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal highSweepSpeed As Integer) As Long
Declare Function hp816x_returnEquidistantData Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal equallySpacedDatapoints As Integer) As Long
Declare Function hp816x_set_LambdaScan_wavelength Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal powermeterWavelength As Double) As Long
Declare Function hp816x_prepareLambdaScan Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal unit As Long, ByVal power As Double, ByVal opticalOutput As Long, ByVal numberofScans As Long, ByVal PWMChannels As Long, ByVal startpWavelength As Double, ByVal stoppWavelength As Double, ByVal stepSize As Double, numberofDatapoints As Long, numberofValueArrays As Long) As Long
Declare Function hp816x_getLambdaScanParameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, startpWavelength As Double, stoppWavelength As Double, averagingTime As Double, sweepSpeed As Double) As Long
Declare Function hp816x_executeLambdaScan Lib "hp816x_32.dll" (ByVal IHandle As Long, wavelengthArray As Double, powerArray1 As Double, powerArray2 As Double, powerArray3 As Double, powerArray4 As Double, powerArray5 As Double, powerArray6 As Double, powerArray7 As Double, powerArray8 As Double) As Long
' ****************************************************************************
' ***********************  Multiple Mainframe Support              ***********
' ****************************************************************************
Declare Function hp816x_registerMainframe Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_unregisterMainframe Lib "hp816x_32.dll" (ByVal IHandle As Long) As Long
Declare Function hp816x_setSweepSpeed Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal sweepSpeed As Long) As Long
Declare Function hp816x_prepareMfLambdaScan Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal powerUnit As Long, ByVal power As Double, ByVal opticalOutput As Long, ByVal numberofScans As Long, ByVal PWMChannels As Long, ByVal startpWavelength As Double, ByVal stoppWavelength As Double, ByVal stepSize As Double, numberofDatapoints As Long, numberofChannels As Long) As Long
Declare Function hp816x_executeMfLambdaScan Lib "hp816x_32.dll" (ByVal IHandle As Long, wavelengthArray As Double) As Long
Declare Function hp816x_getMFLambdaScanParameters_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, startpWavelength As Double, stoppWavelength As Double, averagingTime As Double, sweepSpeed As Double) As Long
Declare Function hp816x_getLambdaScanResult Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMChannel As Long, ByVal clippUnderRange As Integer, ByVal clippingLimit As Double, powerArray As Double, lambdaArray As Double) As Long
Declare Function hp816x_getNoOfRegPWMChannels_Q Lib "hp816x_32.dll" (ByVal IHandle As Long, numberofPWMChannels As Long) As Long
Declare Function hp816x_getChannelLocation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMChannel As Long, mainframeNumber As Long, slotNumber As Long, channelNumber As Long) As Long
Declare Function hp816x_excludeChannel Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMChannel As Long, ByVal excludeChannel As Integer) As Long
Declare Function hp816x_setInitialRangeParams Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal PWMChannel As Long, ByVal resettoDefault As Integer, ByVal initialRange As Double, ByVal rangeDecrement As Double) As Long
Declare Function hp816x_setScanAttenuation Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal scanAttenuation As Double) As Long
Declare Function hp816x_getWlRespTblSize Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, size As Long, CSVSize As Long) As Long
Declare Function hp816x_getWlRespTblSizeEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal chan As Long, size As Long, CSVSize As Long) As Long
Declare Function hp816x_readWlRespTable Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, wavelength As Double, responseFactor As Double) As Long
Declare Function hp816x_readWlRespTableEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal chan As Long, wavelength As Double, responseFactor As Double) As Long
Declare Function hp816x_readWlRepTblCSV Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal CSVList As String) As Long
Declare Function hp816x_readWlRepTblCSV_Ex Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal chan As Long, ByVal CSVList As String) As Long
Declare Function hp816x_spectralCalibration Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal size_of_Spectrum As Long, wavelength As Double, power As Double, wavelengthResult As Double, ByVal error_Diagnose As String) As Long
Declare Function hp816x_spectralCalibrationEx Lib "hp816x_32.dll" (ByVal IHandle As Long, ByVal slot As Long, ByVal chan As Long, ByVal size_of_Spectrum As Long, wavelength As Double, power As Double, wavelengthResult As Double, ByVal error_Diagnose As String) As Long
