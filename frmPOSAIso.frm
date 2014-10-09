VERSION 5.00
Begin VB.Form frmPOSAIso 
   Caption         =   "POSA Iso Testing Ver. 3.141010"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15855
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10110
   ScaleWidth      =   15855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCH1 
      Caption         =   "CH 1"
      Height          =   735
      Left            =   4800
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCH2 
      Caption         =   "CH 2"
      Height          =   735
      Left            =   4800
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCH4 
      Caption         =   "CH 4"
      Height          =   735
      Left            =   4800
      TabIndex        =   37
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCH3 
      Caption         =   "CH 3"
      Height          =   735
      Left            =   4800
      TabIndex        =   36
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   8535
      Left            =   7000
      ScaleHeight     =   8475
      ScaleWidth      =   8475
      TabIndex        =   31
      Top             =   480
      Width           =   8535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Testing Result"
      Height          =   4095
      Left            =   120
      TabIndex        =   19
      Top             =   5040
      Width           =   6615
      Begin VB.TextBox txtILMax 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtWL_L 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtBD_L 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtWLDelta_L 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtIso_L 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtCWL 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtBW 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtWL_R 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtBD_R 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtWLDelta_R 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtIso_R 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtIL 
         Height          =   435
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblILMax 
         Caption         =   "IL_Max :"
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "IL_ITU :"
         Height          =   495
         Left            =   3360
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Iso_Right :"
         Height          =   495
         Left            =   3360
         TabIndex        =   29
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "WL_Delta_R :"
         Height          =   495
         Left            =   3360
         TabIndex        =   28
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Boundary_R :"
         Height          =   495
         Left            =   3360
         TabIndex        =   27
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "0.5dBWL_R :"
         Height          =   495
         Left            =   3360
         TabIndex        =   26
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Bandwidth :"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "CWL_ITU :"
         Height          =   495
         Left            =   3360
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Iso_Left :"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "WL_Delta_L :"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Boundary_L :"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "0.5dBWL_L :"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Testing Info"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.ComboBox cboMode 
         Height          =   435
         ItemData        =   "frmPOSAIso.frx":0000
         Left            =   1680
         List            =   "frmPOSAIso.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080FF80&
         Caption         =   "Save 保存"
         Height          =   735
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "Calc"
         Height          =   735
         Left            =   600
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtStopWav 
         Height          =   435
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "1314"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtStartWav 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "1302"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtIsoSpec 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2805
         Width           =   2535
      End
      Begin VB.TextBox txtTestPoint 
         Height          =   435
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2250
         Width           =   2535
      End
      Begin VB.TextBox txtWO 
         Height          =   435
         Left            =   1680
         TabIndex        =   6
         Top             =   1710
         Width           =   2535
      End
      Begin VB.TextBox txtFilePath 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   5
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox txtSN 
         Height          =   435
         Left            =   1680
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cboCH 
         Height          =   435
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Mode :"
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   615
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Iso/IL Spec :"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1485
      End
      Begin VB.Label Label9 
         Caption         =   "Iso TPoint :"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "PN :"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "FilePath :"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "SN :"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "CH :"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lblPBup 
      Height          =   375
      Left            =   7000
      TabIndex        =   50
      Top             =   9400
      Width           =   8400
   End
   Begin VB.Label lblPBdown 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   7000
      TabIndex        =   49
      Top             =   9400
      Width           =   8400
   End
   Begin VB.Label lblY 
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   35
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblY 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   34
      Top             =   9000
      Width           =   375
   End
   Begin VB.Label lblX 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   14760
      TabIndex        =   33
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label lblX 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   32
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   -120
      TabIndex        =   17
      Top             =   9360
      Width           =   6975
   End
End
Attribute VB_Name = "frmPOSAIso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim dblStartWav() As Double
'Dim dblStopWav() As Double
'Dim strLambda() As String
'Dim strDelta() As String
Private Type tTestSpec

    dblStartWav As Double
    dblStopWav As Double
    strIsoTestPoint() As String
    strIsoSpec() As String
    strILTestSeg() As String
    strILDelta() As String
    strILSpec() As String

End Type

Private Type tTestData

    StrTestName As String
    dblCentWL_ITU As Double
    dblIL_ITU As Double
    dblBandWidth As Double
    
    dblWL_dBDown(1) As Double
    dblIL_dBDown(1) As Double
    dblIsoDelta(1) As Double
    dblWavDelta(1) As Double

End Type


Dim dblWLData() As Double
Dim dblILData() As Double
Dim strStatus As String

Dim ChannelNum As Integer
Dim flgRepeat As Boolean
Dim CHSpec() As tTestSpec
Dim CHData() As tTestData

Private Sub InitPlotArea(x1 As Double, x2 As Double, y1 As Double, y2 As Double)
'*****************************************************************
'   Clear and scale plot area
'*****************************************************************
    
    Picture1.Cls
    Picture1.Scale (x1, y2)-(x2, y1)
'    Picture1.AutoRedraw = True
'    Picture1.Scale (X1, Y1)-(X2, Y2)
    Picture1.DrawWidth = 1
    lblX(0).Caption = x1
    lblX(1).Caption = x2
    lblY(0).Caption = y1
    lblY(1).Caption = y2
    
    DoEvents
End Sub
'Dim SantecSetting As tInstrSetting
Private Sub PlotData(ArrX() As Double, ArrY() As Double, NumOfData As Long, Plotcolor As Long)
'*****************************************************************
'   Plot data points
'*****************************************************************

Dim Cnt As Long
Dim X0 As Double
Dim Y0 As Double
Dim x1 As Double
Dim y1 As Double

    Picture1.ForeColor = Plotcolor
    X0 = ArrX(0)
    Y0 = ArrY(0)
    For Cnt = 1 To NumOfData - 1
        x1 = ArrX(Cnt)
        y1 = ArrY(Cnt)
        
        Picture1.Line (X0, Y0)-(x1, y1)
        DoEvents
        X0 = x1
        Y0 = y1
    Next Cnt
End Sub
Private Sub DisplaySpec(ByVal pCH As Integer)

    cboCH.Text = pCH
    txtStartWav.Text = CHSpec(pCH).dblStartWav
    txtStopWav.Text = CHSpec(pCH).dblStopWav
    txtTestPoint.Text = Join(CHSpec(pCH).strIsoTestPoint, ";")
    txtIsoSpec.Text = Join(CHSpec(pCH).strIsoSpec, ";") & " | " & Join(CHSpec(pCH).strILSpec, ";") & " | " & Join(CHSpec(pCH).strILDelta, ";")
    

End Sub

Private Sub LoadSpec(ByVal pPN As String)

On Error GoTo ErrorHandle

    Dim I As Integer

    'Testing Spec
    ChannelNum = Val(INIRead(pPN, "ChannelNum", App.Path & "\\Settings"))
    
    'Clear cboCH then add item
    cboCH.Clear
    For I = 0 To ChannelNum
        cboCH.AddItem I
    Next I
    
    ReDim CHSpec(ChannelNum) As tTestSpec
    ReDim CHData(ChannelNum) As tTestData
    For I = 0 To ChannelNum
        CHSpec(I).dblStartWav = INIRead(pPN, "StartWav" & I, App.Path & "\\Settings")
        CHSpec(I).dblStopWav = INIRead(pPN, "StopWav" & I, App.Path & "\\Settings")
        CHSpec(I).strIsoTestPoint = Split(INIRead(pPN, "IsoTestPoint" & I, App.Path & "\\Settings"), ";")
        CHSpec(I).strIsoSpec = Split(INIRead(pPN, "IsoSpec" & I, App.Path & "\\Settings"), ";")
        CHSpec(I).strILTestSeg = Split(INIRead(pPN, "ILTestSeg" & I, App.Path & "\\Settings"), ";")
        CHSpec(I).strILSpec = Split(INIRead(pPN, "ILSpec" & I, App.Path & "\\Settings"), ";")
        CHSpec(I).strILDelta = Split(INIRead(pPN, "ILDelta" & I, App.Path & "\\Settings"), ";")
    Next I
    
    
    
    Exit Sub
ErrorHandle:
    MsgBox " Spec 文件设置错误，请检查相应设置！！"
    Call mMakeErrorLog(Err.Number, Err.Description, "SaveCalData", Erl)
    Err.Clear
End Sub
Private Function LoadPresetFile(ByVal filepath As String, ByVal pfileName As String) As Boolean
    Dim iFile As Integer
    Dim strPathFile As String
    Dim index As Long
    Dim strTemp2() As String
    Dim strTemp As String
On Error GoTo ErrorHandle
    LoadPresetFile = True
    strPathFile = filepath & "\\" & pfileName & ".csv"
    
    If Dir(strPathFile, vbDirectory) = "" Then
           MsgBox "Not such file found.. pls check if SN is correct..." & vbCrLf & vbCrLf & "没有找到该SN的文件，请检查SN号是否正确。。。"
           LoadPresetFile = False
           Exit Function
    End If
    
    iFile = FreeFile

    Open strPathFile For Input As #iFile
        
        'Get data-------------------------------------------
        index = -13
        
        Do While Not EOF(iFile) 'Input data lines

            Line Input #iFile, strTemp
            
                If index > -1 Then
                ReDim Preserve dblWLData(0 To index)
                ReDim Preserve dblILData(0 To index)
                strTemp2 = Split(strTemp, ",")
                dblWLData(index) = Val(strTemp2(0))
                dblILData(index) = Val(strTemp2(1))
                End If
            index = index + 1
            
        Loop
    Close #iFile
    Exit Function
ErrorHandle:
    LoadPresetFile = False
    MsgBox "Loading testing data error, pls check the file.." & vbCrLf & vbCrLf & "加载测试数据文件错误，请检查该SN的cvs文件。。"
    Call mMakeErrorLog(Err.Number, Err.Description, "LoadPresetFile", Erl)
    Err.Clear
End Function



Private Function SaveCalData(ByVal strData As String, ByVal filepath As String) As Boolean
    Dim DefaultFileName As String
    Dim I As Long
    Dim iFile As Integer
    Dim fileTitle As String

On Error GoTo ErrorHandle
    SaveCalData = False
    
    DefaultFileName = filepath & "\\" & Date & ".csv"
    
    If Dir(DefaultFileName, vbNormal) = "" Then
        iFile = FreeFile
        fileTitle = "SN,CH,P/F,CWL_ITU,IL_ITU,IL_Max,BW_ITU,0.5dBWL_Left,Boundary_Left,WLDelta_Left,0.5dBWL_Right,Boundary_Right,WLDelta_Right,Iso_Left,Iso_Right,Date"
        
        Open DefaultFileName For Output As #iFile
            Print #iFile, fileTitle
        Close #iFile
        
    End If
    
    'If file not exist, then create one------------------
    If Dir(DefaultFileName, vbNormal) <> "" Then
    
        iFile = FreeFile
            
        Open DefaultFileName For Append As #iFile   '这种写法为追加
        
                Print #iFile, strData
                
        Close #iFile
        
    End If
    '----------------------------------------------------
    
    
    
    lblStatus.Caption = "Save result data OK!!"
    SaveCalData = True
    
    Exit Function
ErrorHandle:
    Call mMakeErrorLog(Err.Number, Err.Description, "SaveCalData", Erl)
    Err.Clear
End Function




Private Sub cboCH_Click()

    DisplaySpec (Val(Trim(cboCH.Text)))

End Sub

Private Sub cboMode_Click()
    If cboMode.Text = "Analyze & Save File" Then
        txtFilePath.Locked = True
        txtFilePath.Text = "D:\\POSA Result\\Raw Data"
        txtFilePath.BackColor = mButtonGray
    ElseIf cboMode.Text = "Analyze File Only" Then
        txtFilePath.Locked = False
        txtFilePath.BackColor = &H80000005
        txtFilePath.Text = ""
    End If
End Sub

Private Sub cmdCalc_Click()

If cboMode.Text = "Analyze & Save File" Then
    If Dir(txtFilePath.Text, vbDirectory) = "" Then
        MkDir (txtFilePath.Text)
    End If
    
    Call LoadPresetFile("C:\POSA\", "temp")

ElseIf cboMode.Text = "Analyze File Only" Then
    If txtFilePath.Text = "" Then
        MsgBox "File path can not be empty." & vbCrLf & vbCrLf & "文件路径不能为空."
        Exit Sub
    End If
    If Dir(txtFilePath.Text, vbDirectory) = "" Then
        MsgBox "Pls check your archive directory." & vbCrLf & vbCrLf & "请检查你的存档路径是否存在."
        Exit Sub
    End If
    Call LoadPresetFile(txtFilePath.Text, txtSN.Text & "-" & cboCH.Text)

Else

    MsgBox "Pls select a mode !"
    Exit Sub
End If
    Call CalcRawData


End Sub

Private Function CalcRawData() As Boolean

    Dim I As Long
    Dim intMaxIndex As Long
    Dim maxILIndex As Long
    Dim minILIndex As Long
    Dim minILIndex2 As Long
    Dim maxIL As Double
    Dim minIL As Double
    Dim minIL2 As Double
    Dim CHNum As Integer
    Dim tempIndex(3) As Integer
    Dim tempILIndex(1) As Integer
    Dim dblSampleStep As Double
    On Error GoTo ErrorHandle
        
    CalcRawData = False
    
    intMaxIndex = UBound(dblWLData())
    maxIL = -999.99
    minIL = 999.99
    minIL2 = 999.99
    
    CHNum = Val(cboCH.Text)
    dblSampleStep = 0.001
    strStatus = "Passed"
    tempIndex(0) = CInt((Val(CHSpec(CHNum).strIsoTestPoint(0)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    tempIndex(1) = CInt((Val(CHSpec(CHNum).strIsoTestPoint(1)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    tempIndex(2) = CInt((Val(CHSpec(CHNum).strIsoTestPoint(2)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    tempIndex(3) = CInt((Val(CHSpec(CHNum).strIsoTestPoint(3)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    
    tempILIndex(0) = CInt((Val(CHSpec(CHNum).strILTestSeg(0)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    tempILIndex(1) = CInt((Val(CHSpec(CHNum).strILTestSeg(1)) - CHSpec(CHNum).dblStartWav) / dblSampleStep)
    
    '0 Clear the textboxs' colors
    txtBW.BackColor = &H80000005
    txtCWL.BackColor = &H80000005
    txtILMax.BackColor = &H80000005
    txtIL.BackColor = &H80000005
    txtWL_L.BackColor = &H80000005
    txtWL_R.BackColor = &H80000005
    txtBD_L.BackColor = &H80000005
    txtBD_R.BackColor = &H80000005
    txtWLDelta_L.BackColor = &H80000005
    txtWLDelta_R.BackColor = &H80000005
    txtIso_L.BackColor = &H80000005
    txtIso_R.BackColor = &H80000005
    '------------------------------
    txtBW.Text = ""
    txtCWL.Text = ""
    txtILMax.Text = ""
    txtIL.Text = ""
    txtWL_L.Text = ""
    txtWL_R.Text = ""
    txtBD_L.Text = ""
    txtBD_R.Text = ""
    txtWLDelta_L.Text = ""
    txtWLDelta_R.Text = ""
    txtIso_L.Text = ""
    txtIso_R.Text = ""
    
    If dblWLData(0) <> Val(txtStartWav.Text) Or dblWLData(intMaxIndex) <> Val(txtStopWav.Text) Then
        MsgBox "没有找到相关波长数据，请检查CH是否正确或检查数据文件是否保存正确。。"
        Exit Function
    End If
    
    '1. find the top IL point..(maxIL),ITU_IL(minIL)
    For I = 0 To intMaxIndex

        If dblILData(I) > maxIL Then
            maxIL = dblILData(I)
            maxILIndex = I
        End If

    Next I
    


    'minIL in IL bandwith
    For I = tempILIndex(0) To tempILIndex(1)
    
        If dblILData(I) < minIL Then
            minIL = dblILData(I)
            minILIndex = I
        End If
    
    Next I
    CHData(CHNum).dblIL_ITU = minIL
    
    'minIl in Iso Bandwith
    
    For I = tempIndex(1) To tempIndex(2)
        If dblILData(I) < minIL2 Then
            minIL2 = dblILData(I)
            minILIndex2 = I
        End If
    Next I
    
        '=============================从中间往两边找=====================================
    '2. 0.5dB down of minIL point
    For I = minILIndex To 1 Step -1
        If dblILData(I) + 0.5 - minIL > 0# And dblILData(I - 1) + 0.5 - minIL < 0# Then
            CHData(CHNum).dblWL_dBDown(0) = dblWLData(I - 1)
            CHData(CHNum).dblIL_dBDown(0) = dblILData(I - 1)
            Exit For
        End If
    Next I

    For I = minILIndex + 1 To intMaxIndex - 1

        If dblILData(I) + 0.5 - minIL > 0# And dblILData(I + 1) + 0.5 - minIL < 0# Then

            CHData(CHNum).dblWL_dBDown(1) = dblWLData(I + 1)
            CHData(CHNum).dblIL_dBDown(1) = dblILData(I + 1)
            Exit For
        End If

    Next I
    
    
'    '=============================从中间往两边找=====================================
'    '2. 0.5dB down of MaxIL point
'    For I = maxILIndex To 1 Step -1
'        If dblILData(I) + 0.5 - maxIL > 0# And dblILData(I - 1) + 0.5 - maxIL < 0# Then
'            CHData(CHNum).dblWL_dBDown(0) = dblWLData(I - 1)
'            CHData(CHNum).dblIL_dBDown(0) = dblILData(I - 1)
'            Exit For
'        End If
'    Next I
'
'    For I = maxILIndex + 1 To intMaxIndex - 1
'
'        If dblILData(I) + 0.5 - maxIL > 0# And dblILData(I + 1) + 0.5 - maxIL < 0# Then
'
'            CHData(CHNum).dblWL_dBDown(1) = dblWLData(I + 1)
'            CHData(CHNum).dblIL_dBDown(1) = dblILData(I + 1)
'            Exit For
'        End If
'
'    Next I
    '===============================================================================
    '*******************************************************************************
    '=============================从两边往中间找=====================================
    ''2. 0.5dB down of MaxIL point
    'For I = 1 To intMaxIndex - 1
    '    If dblILData(I - 1) + 0.5 - maxIL < 0# And dblILData(I) + 0.5 - maxIL > 0# Then
    '
    '        CHData(CHNum).dblWL_dBDown(0) = dblWLData(I - 1)
    '        CHData(CHNum).dblIL_dBDown(0) = dblILData(I - 1)

    '    End If
    '
    '    If dblILData(intMaxIndex - I) + 0.5 - maxIL < 0# And dblILData(intMaxIndex - I - 1) + 0.5 - maxIL > 0# Then
    '
    '        CHData(CHNum).dblWL_dBDown(1) = dblWLData(intMaxIndex - I)
    '        CHData(CHNum).dblIL_dBDown(1) = dblILData(intMaxIndex - I)

    '    End If
    '
    'Next I
    '================================================================================================
    '3. bandwidth,CWL
    CHData(CHNum).dblBandWidth = CHData(CHNum).dblWL_dBDown(1) - CHData(CHNum).dblWL_dBDown(0)
    CHData(CHNum).dblCentWL_ITU = (CHData(CHNum).dblWL_dBDown(0) + CHData(CHNum).dblWL_dBDown(1)) / 2#

    'wavelengthDelta,isodelta
    CHData(CHNum).dblWavDelta(0) = CHData(CHNum).dblWL_dBDown(0) - Val(CHSpec(CHNum).strIsoTestPoint(1))
    CHData(CHNum).dblWavDelta(1) = CHData(CHNum).dblWL_dBDown(1) - Val(CHSpec(CHNum).strIsoTestPoint(2))
'    CHData(CHNum).dblIsoDelta(0) = Abs(dblILData(tempIndex(1)) - dblILData(tempIndex(0)))
'    CHData(CHNum).dblIsoDelta(1) = Abs(dblILData(tempIndex(2)) - dblILData(tempIndex(3)))
    CHData(CHNum).dblIsoDelta(0) = Abs(dblILData(minILIndex2) - dblILData(tempIndex(0)))
    CHData(CHNum).dblIsoDelta(1) = Abs(dblILData(minILIndex2) - dblILData(tempIndex(3)))
    
    '5.display data
    txtBW.Text = Format(CHData(CHNum).dblBandWidth, "#.###")
    txtCWL.Text = Format(CHData(CHNum).dblCentWL_ITU, "#.###")
    txtILMax.Text = Format(maxIL, "#.###")
    txtIL.Text = Format(CHData(CHNum).dblIL_ITU, "#.###")
    txtWL_L.Text = Format(CHData(CHNum).dblWL_dBDown(0), "#.###")
    txtWL_R.Text = Format(CHData(CHNum).dblWL_dBDown(1), "#.###")
    txtBD_L.Text = Format(CHSpec(CHNum).strIsoTestPoint(1), "#.###")
    txtBD_R.Text = Format(CHSpec(CHNum).strIsoTestPoint(2), "#.###")
    txtWLDelta_L.Text = Format(CHData(CHNum).dblWavDelta(0), "#.###")
    txtWLDelta_R.Text = Format(CHData(CHNum).dblWavDelta(1), "#.###")
    txtIso_L.Text = Format(CHData(CHNum).dblIsoDelta(0), "#.###")
    txtIso_R.Text = Format(CHData(CHNum).dblIsoDelta(1), "#.###")
    '6. display color
    '=======Isodelta========
    If CHData(CHNum).dblIsoDelta(0) < CHSpec(CHNum).strIsoSpec(0) Then
        txtIso_L.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtIso_L.BackColor = mLightGreen
    End If
    
    If CHData(CHNum).dblIsoDelta(1) < CHSpec(CHNum).strIsoSpec(1) Then
        txtIso_R.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtIso_R.BackColor = mLightGreen
    End If
    '========WavlengthDelta=========
    If CHData(CHNum).dblWavDelta(0) > 0 Then
        txtWLDelta_L.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtWLDelta_L.BackColor = mLightGreen
    End If
    
    If CHData(CHNum).dblWavDelta(1) < 0 Then
        txtWLDelta_R.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtWLDelta_R.BackColor = mLightGreen
    End If
'    'update 2014 Jun.17th
'    'add 如果WLDelta_Left（J列）或WLDelta_Right （M列）的绝对值小于0.1nm，
'    '则必须保证0.5dB带宽（H列和K列之差）〉= 2.9nm
'
'    If CHData(CHNum).dblWavDelta(0) < 0 And Abs(CHData(CHNum).dblWavDelta(0)) < 0.1 Then
'        txtWLDelta_L.BackColor = mLightYellow
'        If CHData(CHNum).dblBandWidth < 2.9 Then
'            txtBW.BackColor = mLightRed
'            strStatus = "Failed"
'        End If
'    End If
'
'    If CHData(CHNum).dblWavDelta(1) > 0 And Abs(CHData(CHNum).dblWavDelta(1)) < 0.1 Then
'        txtWLDelta_R.BackColor = mLightYellow
'        If CHData(CHNum).dblBandWidth < 2.9 Then
'            txtBW.BackColor = mLightRed
'            strStatus = "Failed"
'        End If
'    End If
    '============IL===============
    If CHData(CHNum).dblIL_ITU > Val(CHSpec(CHNum).strILSpec(1)) Or CHData(CHNum).dblIL_ITU < Val(CHSpec(CHNum).strILSpec(0)) Then
        txtIL.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtIL.BackColor = mLightGreen
    End If
    
    If maxIL > Val(CHSpec(CHNum).strILSpec(1)) Or maxIL < Val(CHSpec(CHNum).strILSpec(0)) Then
        txtILMax.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtILMax.BackColor = mLightGreen
    End If
    
    '=============IL Boundry================
    If Abs(CHData(CHNum).dblWL_dBDown(0) - Val(CHSpec(CHNum).strILTestSeg(0))) < Val(CHSpec(CHNum).strILDelta(0)) Then
        txtWL_L.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtWL_L.BackColor = mLightGreen
    
    End If
    
    If Abs(CHData(CHNum).dblWL_dBDown(1) - Val(CHSpec(CHNum).strILTestSeg(1))) < Val(CHSpec(CHNum).strILDelta(1)) Then
        txtWL_R.BackColor = mLightRed
        strStatus = "Failed"
    Else
        txtWL_R.BackColor = mLightGreen
    
    End If
    
    'Display data
    
    Call InitPlotArea(txtStartWav.Text, txtStopWav.Text, -60, 5)
    'Picture1.Line (startWav, 0)-(startWav, 0), vbRed
    Call PlotData(dblWLData(), dblILData(), intMaxIndex, vbYellow)
    '        Call PlotData(wavelengthArray(), powerArray2(), NumOfData, vbRed)
    
    CalcRawData = True

Exit Function
ErrorHandle:
    CalcRawData = False
    strStatus = "Failed"
    Call mMakeErrorLog(Err.Number, Err.Description, "CalcRawData", Erl)
    Err.Clear
End Function

Private Sub cmdCH1_Click()
cboCH.Text = "1"
Call cboCH_Click
Call cmdCalc_Click

End Sub

Private Sub cmdCH2_Click()
cboCH.Text = "2"
Call cboCH_Click
Call cmdCalc_Click

End Sub

Private Sub cmdCH3_Click()
cboCH.Text = "3"
Call cboCH_Click
Call cmdCalc_Click

End Sub

Private Sub cmdCH4_Click()
cboCH.Text = "4"
Call cboCH_Click
Call cmdCalc_Click

End Sub

Private Sub cmdSave_Click()
'7 save calculation data
Dim calDataPath As String
Dim rawDataPath As String
Dim calResult As String
On Error GoTo ErrorHandle

'1 general check.
If txtSN.Text = "" Then
    MsgBox "Pls input a correct SN." & vbCrLf & vbCrLf & "请输入正确的SN."
    txtSN.SetFocus
    Exit Sub
End If

If txtWO.Text = "" Then
    MsgBox "Pls input a correct WO." & vbCrLf & vbCrLf & "请输入正确的WO."
    txtWO.SetFocus
    Exit Sub
End If

If txtFilePath.Text = "" Then

    MsgBox "Pls input a correct path." & vbCrLf & vbCrLf & "请输入正确的保存路径."
    txtFilePath.SetFocus
    Exit Sub
End If

'2. Save data
    
txtSN.Text = Replace(Trim(txtSN.Text), " ", "")
txtWO.Text = Replace(Trim(txtWO.Text), " ", "")

If CalcRawData = True Then
    calDataPath = "D:\\POSA Result\\Cal Data"
    calResult = txtSN.Text & "," & cboCH.Text & "," & strStatus & "," & txtCWL.Text & "," & txtIL.Text & "," & txtILMax.Text & "," _
                & txtBW.Text & "," & txtWL_L.Text & "," & txtBD_L.Text & "," & txtWLDelta_L.Text & "," _
                & txtWL_R.Text & "," & txtBD_R.Text & "," & txtWLDelta_R.Text & "," & txtIso_L.Text & "," & txtIso_R.Text & "," & Now
    If cboCH.Text <> "0" Then
        Call SaveCalData(calResult, calDataPath)
    End If
    
    If cboMode.Text = "Analyze & Save File" Then
        rawDataPath = txtFilePath.Text & "\\" & txtWO.Text
        
        If Dir(rawDataPath, vbDirectory) = "" Then
            MkDir (rawDataPath)
        End If
        
        If Dir(rawDataPath & "\\" & Date, vbDirectory) = "" Then
            MkDir (rawDataPath & "\\" & Date)
        End If
        
        FileCopy "C:\POSA\temp.csv", rawDataPath & "\\" & Date & "\\" & txtSN.Text & "-" & cboCH.Text & ".csv"
    End If
    
Else

    MsgBox "Pls prefer which channel to analyze first.." & "请选择要测试的CH。。"

End If

txtSN.Text = ""
txtSN.SetFocus
Exit Sub
ErrorHandle:

    Call mMakeErrorLog(Err.Number, Err.Description, "cmdSave_Click", Erl)
    Err.Clear
    txtSN.Text = ""
    txtSN.SetFocus
    
End Sub

Private Sub Form_Load()

Call LoadSpec("Default WO")
cboMode.ListIndex = 0
End Sub




