VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "santec Swept Test System sample software"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   13395
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   600
      TabIndex        =   35
      Top             =   8160
      Width           =   5895
   End
   Begin VB.CommandButton cmdCalibration 
      Caption         =   "Calibration"
      Height          =   735
      Left            =   600
      TabIndex        =   20
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   5535
      Left            =   4320
      ScaleHeight     =   5475
      ScaleWidth      =   7035
      TabIndex        =   19
      Top             =   2040
      Width           =   7095
   End
   Begin VB.CommandButton cmdStartMeas 
      Caption         =   "Measurement"
      Height          =   735
      Left            =   2160
      TabIndex        =   18
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Measurement Settings"
      Height          =   3615
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   2895
      Begin VB.TextBox txtScanNum 
         Height          =   270
         Left            =   1560
         TabIndex        =   33
         Text            =   "2"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtOPMRangeStep 
         Height          =   270
         Left            =   1560
         TabIndex        =   30
         Text            =   "30"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtOPMRangeInit 
         Height          =   270
         Left            =   1560
         TabIndex        =   21
         Text            =   "0"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtStep 
         Height          =   270
         Left            =   1560
         TabIndex        =   13
         Text            =   "0.005"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtStop 
         Height          =   270
         Left            =   1560
         TabIndex        =   12
         Text            =   "1610"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtStart 
         Height          =   270
         Left            =   1560
         TabIndex        =   11
         Text            =   "1510"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Scan #"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Step (dB)"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Initial  (dBm)"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Averaging Time(s)"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblAvgTime 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.001"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblSpeed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "OPM Range"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Speed (nm/s)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Step (nm)"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Stop (nm)"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Start (nm)"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txtSPUAddress 
      Height          =   270
      Left            =   2760
      TabIndex        =   5
      Text            =   "17"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtTSLAddress 
      Height          =   270
      Left            =   2760
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtOPMAddress 
      Height          =   270
      Left            =   2760
      TabIndex        =   1
      Text            =   "20"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "connect"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   26
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   25
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   24
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   23
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Input instrments address"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblSPUAddress 
      Caption         =   "SPU-100 device number"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblOPMAddress 
      Caption         =   "OPM GPIB Address"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblTSLAddress 
      Caption         =   "TSL-510 GPIB Address"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblTSL_ID 
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblOPM_ID 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************
'   santec Swept Test System sample software
'
'   This program is a sample program to evaluate DLL functions
'   for Swept Test System.
'
'   2011.01.13 santec corporation
'*****************************************************************



Private Sub cmdCalibration_Click()
'*****************************************************************
'   Calibration function for TSL-510 power monitor
'*****************************************************************
Dim Cnt As Long                 'Loop counter
Dim PMCalDataArr(3) As Double   'TSL power monitor calibration data
Dim status As Long

'Display
    List1.AddItem "Calibration of power monitor."
    'Disables Calibration button
    cmdCalibration.Enabled = False
    cmdStartMeas.Enabled = False
    Form1.MousePointer = vbHourglass

'Device control
    'Power Monitor Calibration
    status = STS_Calibration(PMCalDataArr())

'Display
    'Displays cal data
    If status = 0 Then
        List1.AddItem "Finished."
        For Cnt = 0 To 3
            List1.AddItem "Range" & Cnt & ": " & Format(PMCalDataArr(Cnt), "0.000")
        Next Cnt
    Else
            List1.AddItem "Calibration error!"
    End If
    
    'Enables Start measurement button
    cmdCalibration.Enabled = True
    cmdStartMeas.Enabled = True
    Form1.MousePointer = vbDefault
    
    DoEvents
    
    
End Sub

Private Sub cmdInit_Click()
'*****************************************************************
'   Initialize instruments
'*****************************************************************

Dim TSLResourceName As String
Dim TSL_ID As String
Dim OPMResourceName As String
Dim OPM_ID As String
Dim SPUResourceName As String
Dim startWav As Double  'Sweep start wavelength
Dim stopWav As Double   'Sweep stop wavelength
Dim TrigStep As Double  'Trigger output step
Dim SwpSpeed As Double  'Sweep speed
    
    cmdInit.Enabled = False
    Form1.MousePointer = vbHourglass


    Call STS_Initialize(Val(txtTSLAddress.Text), Val(txtOPMAddress.Text), Val(txtSPUAddress.Text), _
                        TSL_ID, OPM_ID)
    Form1.Show
    lblTSL_ID.Caption = Format(TSL_ID)
    DoEvents
    lblOPM_ID.Caption = Format(OPM_ID)
    DoEvents
    
    
    'Enables Calibration button
    cmdCalibration.Enabled = True
    Form1.MousePointer = vbDefault

End Sub

Private Sub cmdStartMeas_Click()
'*****************************************************************
'   Measurement function
'*****************************************************************

Dim startWav As Double
Dim stopWav As Double
Dim samplStep As Double
Dim sweepSpeed As Double
Dim AvgTime As Double
Dim NumOfData As Long
Dim SourcePower As Double
Dim NumofScan As Long
Dim InitPowerRange As Double
Dim RangeDecriment As Double
Dim NumofChannels As Long
Dim wavelengthArray() As Double
Dim powerArray1() As Double
Dim powerArray2() As Double
Dim powerArray3() As Double
Dim powerArray4() As Double
Dim powerArray5() As Double
Dim powerArray6() As Double
Dim powerArray7() As Double
Dim powerArray8() As Double
Dim slot As Long

    'Disables Start measurement button
    cmdCalibration.Enabled = False
    cmdStartMeas.Enabled = False
    Form1.MousePointer = vbHourglass

    'Input data from controls
    startWav = CDbl(txtStart.Text)
    stopWav = CDbl(txtStop.Text)
    samplStep = CDbl(txtStep.Text)
    NumofScan = CLng(txtScanNum)
    InitPowerRange = CDbl(txtOPMRangeInit.Text)
    RangeDecriment = CDbl(txtOPMRangeStep.Text)
    SourcePower = InitPowerRange

    
    Call STS_PrepareLambdaScan(SourcePower, startWav, stopWav, samplStep, _
                                NumOfData, 1, 1)
    
    Call STS_getLambdaScanParameters_Q(startWav, stopWav, AvgTime, sweepSpeed)
   
    txtStart.Text = Format(startWav)
    txtStop.Text = Format(stopWav)
    lblSpeed.Caption = Format(sweepSpeed)
    lblAvgTime.Caption = Format(AvgTime)
    DoEvents
    
    
    Call m2STS_executeLambdaScan(NumofScan, _
                                InitPowerRange, _
                                RangeDecriment, _
                                wavelengthArray(), _
                                powerArray1(), _
                                powerArray2(), _
                                powerArray3(), _
                                powerArray4(), _
                                powerArray5(), _
                                powerArray6(), _
                                powerArray7(), _
                                powerArray8())
                                
   


'Display data

    Call InitPlotArea(1500, 1630, -60, 20)
    
        Call PlotData(wavelengthArray(), powerArray1(), NumOfData, vbYellow)
        Call PlotData(wavelengthArray(), powerArray2(), NumOfData, vbRed)
    
    'Enables Start measurement button
    cmdCalibration.Enabled = True
    cmdStartMeas.Enabled = True
    Form1.MousePointer = vbDefault

DoEvents
End Sub





Private Sub Form_Load()
Dim errStr As String * 255
Dim strLen As Long
Dim module As Long
Dim dBm As Double

    Form1.Show
    
    strLen = 255
    List1.AddItem "Loading SweptSystemShared.DLL."
    DoEvents
    Call MW2dBm(1, dBm)
    List1.AddItem "Done." & dBm
    
    'Enables Initialize buttons
    cmdInit.Enabled = True
    'Disables other buttons
    cmdCalibration.Enabled = False
    cmdStartMeas.Enabled = False
    
End Sub

Private Sub Form_Terminate()

    Call STS_Close
    
    'Enables all buttons
    cmdInit.Enabled = True
    cmdCalibration.Enabled = True
    cmdStartMeas.Enabled = True
    
End Sub

Private Sub InitPlotArea(X1 As Double, X2 As Double, Y1 As Double, Y2 As Double)
'*****************************************************************
'   Clear and scale plot area
'*****************************************************************
    
    Picture1.Cls
    Picture1.Scale (X1, Y2)-(X2, Y1)
    Picture1.DrawWidth = 1
    lblX(0).Caption = X1
    lblX(1).Caption = X2
    lblY(0).Caption = Y1
    lblY(1).Caption = Y2
    
    DoEvents
End Sub
    
Private Sub PlotData(ArrX() As Double, ArrY() As Double, NumOfData As Long, Plotcolor As Long)
'*****************************************************************
'   Plot data points
'*****************************************************************

Dim Cnt As Long
Dim X0 As Double
Dim Y0 As Double
Dim X1 As Double
Dim Y1 As Double

    Picture1.ForeColor = Plotcolor
    X0 = ArrX(0)
    Y0 = ArrY(0)
    For Cnt = 1 To NumOfData - 1
        X1 = ArrX(Cnt)
        Y1 = ArrY(Cnt)
        
        Picture1.Line (X0, Y0)-(X1, Y1)
        DoEvents
        X0 = X1
        Y0 = Y1
    Next Cnt
End Sub


