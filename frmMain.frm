VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Dimmers BIN Monitor"
   ClientHeight    =   12885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   FillStyle       =   0  '단색
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12885
   ScaleWidth      =   13590
   Begin prjDimmersBINmon.ucBINdps ucBINdps1 
      Height          =   7815
      Index           =   0
      Left            =   2040
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   13785
   End
   Begin VB.TextBox txtAVRcnt 
      Alignment       =   2  '가운데 맞춤
      Enabled         =   0   'False
      Height          =   270
      Left            =   10440
      TabIndex        =   11
      Text            =   "0/0"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdDmon 
      Caption         =   "dMon"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboIDX 
      Height          =   300
      Left            =   6360
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSD1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   6
      Top             =   11160
      Width           =   10575
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   13275
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.CommandButton cmdCFG 
         BackColor       =   &H00008000&
         Caption         =   "설 정"
         Height          =   375
         Left            =   9720
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
      Begin VB.Timer tmrAoDo 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6960
         Top             =   360
      End
      Begin VB.CommandButton cmdRunStop 
         BackColor       =   &H00008000&
         Caption         =   "RUN/STOP"
         Height          =   375
         Left            =   8160
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Timer tmrINIT 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   7440
         Top             =   360
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
         Caption         =   "종 료"
         Height          =   375
         Left            =   12240
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H00808080&
         Caption         =   "화면감추기"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin MSWinsockLib.Winsock wsPLC2 
         Left            =   6240
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock wsPLC1 
         Left            =   5760
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.Label lbRelDate 
         BackStyle       =   0  '투명
         Caption         =   "Release date"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label lbRelVersion 
         BackStyle       =   0  '투명
         Caption         =   "Release version"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbTitle 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "[조광조] BIN LEVEL MONITORING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   705
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   9195
      End
      Begin VB.Image imgLogo1 
         BorderStyle     =   1  '단일 고정
         Height          =   495
         Left            =   120
         Picture         =   "frmMain.frx":16AC2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label lbTeam 
         BackColor       =   &H00808080&
         Caption         =   "DASAN-InfoTEK"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Label lbTimeNow 
      BackStyle       =   0  '투명
      Caption         =   "RunTime"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "누적횟수:"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   1010
      Width           =   975
   End
   Begin VB.Label lbUpTime 
      BackStyle       =   0  '투명
      Caption         =   "Up_Time"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lbVS1 
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===========================================================================================
'
'                       2D LEVEL Monitoring System
'                       for BIN5 with SICK LMS-211
'
'                                   BIN5mon V1.00
'
'===========================================================================================


Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" (hpvDest As Any, _
                                           hpvSource As Any, _
                                           ByVal cbCopy As Long)

Private Const relVersion = "v1.00.01"
Private Const relDate = "2021-05-21"

Const NUM_OF_BIN = 6

Dim d1 As Single


Public PLCDataRangeMax As Integer

Public chkUsePLC As Integer

Dim ipAddr(20) As String
Dim ipPort(20) As String

Dim AOdata(33) As Integer
Dim AOdata2(33) As Integer

Dim AOdeep(20, 100) As Integer
Dim Hdeep(20, 100) As Integer
Public AOdeepCNT As Integer
Public AOdeepMAX As Integer        ''<=MAX:99
Public AOdeepFull As Boolean

Dim BinWidth As Integer
''
Dim BinMaxH(33) As Integer
Dim BinMinH(33) As Integer
''
Dim BinTYPE(33) As Integer


Private Sub cmdCFG_Click()

''    frmCFG.txtMaxHH = frmMain.txtMaxHH
''    frmCFG.txtBaseHH = frmMain.txtBaseHH

    If frmCFG.Visible = True Then
        frmCFG.Show
    Else
        frmCFG.Visible = True
    End If
    
    frmCFG.tmrCFG_update

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ret1
    
    ret1 = MsgBox("종료하면 모든 기능이 정지됩니다." & vbCrLf & "정말 종료 하시겠습니까?", vbYesNo)

    If ret1 <> vbYes Then
        Cancel = 1
        Exit Sub
    End If

    Unload frmSettings
    Unload frmCFG
End Sub

Private Sub Form_Click()
    If frmCFG.Visible = True Then
        frmCFG.tmrCFG_update
    End If
End Sub

Private Sub Form_DblClick()
    If frmCFG.Visible = True Then
        frmCFG.tmrCFG_update
    End If
End Sub

''Private Sub cmdDmon_Click()
''    ''' txtSD1 = ucBINmon1(cboIDX.ListIndex).ret_SDXY
''End Sub

Private Sub cmdExit_Click()
    Dim ret1
    
    ret1 = MsgBox("종료하면 모든 기능이 정지됩니다." & vbCrLf & "정말 종료 하시겠습니까?", vbYesNo)

    If ret1 <> vbYes Then
        Exit Sub
    End If

    End
End Sub



Private Sub cmdHide_Click()

    ''frmMain.Visible = False
    frmMain.Hide
    
    
End Sub

Private Sub cmdRunStop_Click()

    ''&H00008000& ''G
    ''&H00000080& ''R
    ''QBColor
  Dim i As Integer
  
''    If cmdRunStop.BackColor = &H8000& Then  ''run
''        For i = 0 To 10
''            ucBINmon1(i).scan_STOP
''        Next i
''        cmdRunStop.BackColor = &H80&        ''stop
''        txtMaxHH.Enabled = True
''    Else  ''stop
''        For i = 0 To 10
''            ucBINmon1(i).set_maxHH CLng(txtMaxHH)
''            ucBINmon1(i).scan_RUN
''        Next i
''        cmdRunStop.BackColor = &H8000&        ''run
''        txtMaxHH.Enabled = False
''    End If
''
        
End Sub

Private Sub Form_GotFocus()
'''
''Dim i
''    For i = 0 To 10
''        ucBINmon1(i).picCON_Cir1
''    Next i
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행되었습니다."
       Unload Me
       End
    End If
    
    lbUpTime.Caption = "Up_Time: " & Format(Now, "YYYY-MM-DD h:m:s")
    lbTimeNow.Caption = "RunTime: " & Format(Now, "YYYY-MM-DD h:m:s")
    
    frmMain.AutoRedraw = True

'    Me.Width = Screen.Width * (1280 / 1400)
'    Me.Height = Screen.Height * (1024 / 1050)

'    Me.Left = Screen.Width - Width
'    Me.Top = 0
'    frmMain.Move Screen.Width - Width, 0
    
    frmMain.Move 0, 0, Screen.Width, Screen.Height
    
    AOdeepMAX = GetSetting(App.Title, "Settings", "DeepMax", 60)
    If AOdeepMAX < 10 Then AOdeepMAX = 10
    If AOdeepMAX > 99 Then AOdeepMAX = 99
    AOdeepFull = False
    AOdeepCNT = 0
    For i = 0 To NUM_OF_BIN - 1
        For j = 0 To 99  ''AOdeepMAX
            AOdeep(i, j) = 0
            Hdeep(i, j) = 0
        Next j
    Next i
    
    chkUsePLC = GetSetting(App.Title, "Settings", "UsePLC", 0)
    If chkUsePLC < 0 Or chkUsePLC > 1 Then chkUsePLC = 0
    
    picTop.Left = 100
    picTop.Top = 100
    picTop.Height = 800   '''Height * 0.05 + 100
    picTop.Width = Width - 200
    ''''
        imgLogo1.Left = 100
        imgLogo1.Top = 100 ''100
        lbTitle.Left = (Width * 0.27)    ''+ 200  ''frTop.Width * 0.3
        lbTitle.Top = 50
        lbTitle.Height = 600
        lbTitle.Width = (Width * 0.5) - 500
        ''
        cmdExit.Top = 200
        cmdExit.Left = picTop.Width - 1200
        cmdHide.Top = 200
        cmdHide.Left = picTop.Width - 2600
        cmdRunStop.Top = 200
        cmdRunStop.Left = picTop.Width - 4000
        
        cmdCFG.Top = 200
        cmdCFG.Left = picTop.Width - 5000
        
        ''lbRelVersion.Top = 200
        ''lbRelVersion.Left = picTop.Width - 6050
        lbRelVersion = relVersion
        ''lbRelDate.Top = 400
        ''lbRelDate.Left = picTop.Width - 6050
        lbRelDate = relDate
        
    For i = 0 To 32
        AOdata(i) = 0
        AOdata2(i) = 0
    Next i
    
    BinWidth = 1800  '''2000
    ''''''''
    
    For i = 1 To NUM_OF_BIN - 1
        Load ucBINdps1(i)
    Next i
    ''''''
    For i = 0 To NUM_OF_BIN - 1
    
        ucBINdps1(i).Top = 1600

        ucBINdps1(i).Width = BinWidth  '''Width / 11 - 50
        ucBINdps1(i).Left = (i * (Width / NUM_OF_BIN)) + (Width / NUM_OF_BIN - BinWidth) / 2
        ucBINdps1(i).Height = 6100  '''12200

        ucBINdps1(i).Visible = True

        DoEvents

    Next i
'''''''''''
    
    txtSD1.Left = 100
    txtSD1.Top = Height - 1800
    txtSD1.Width = Width - 300
    txtSD1.Height = 1300

    ''' Set default IP addr/port for Dimmers
    ipAddr(0) = "192.168.0.21"
    ipPort(0) = "7001"
    ipAddr(1) = "192.168.0.21"
    ipPort(1) = "7002"
    ipAddr(2) = "192.168.0.21"
    ipPort(2) = "7003"
    ipAddr(3) = "192.168.0.21"
    ipPort(3) = "7004"
    '''
    ipAddr(4) = "192.168.0.22"
    ipPort(4) = "7001"
    ipAddr(5) = "192.168.0.22"
    ipPort(5) = "7002"
    ''ipAddr(6) = "192.168.0.22"
    ''ipPort(6) = "7003"
    ''ipAddr(7) = "192.168.0.22"
    ''ipPort(7) = "7004"
    ''ipAddr(8) = "192.168.0.22"
    ''ipPort(8) = "7005"
    ''ipAddr(9) = "192.168.0.22"
    ''ipPort(9) = "7006"
    
    Dim ipAddr_tmp As String
    Dim ipPort1_tmp As String
    Dim ipPort2_tmp As String
    For i = 0 To NUM_OF_BIN - 1
        ipAddr_tmp = GetSetting(App.Title, "Settings", "BinIPAddr_" & i, "Fail")
        ipPort1_tmp = GetSetting(App.Title, "Settings", "BinIPPort_" & i, "Fail")
        If IsValidIPAddress(ipAddr_tmp) = False Then
            ipAddr_tmp = ipAddr(i)
            ''SaveSetting App.Title, "Settings", "BinIPAddr_" & i, ipAddr_tmp
        End If
        If IsValidIPPort(ipPort1_tmp) = False Then
            ipPort1_tmp = ipPort(i)
            ''SaveSetting App.Title, "Settings", "BinIPPort_" & i, ipPort1_tmp
        End If
        ucBINdps1(i).setIDX i, ipAddr_tmp, ipPort1_tmp
    Next i
    
    ipAddr_tmp = GetSetting(App.Title, "Settings", "PLCIPAddr", "Fail")
    ipPort1_tmp = GetSetting(App.Title, "Settings", "PLCIPPort1", "Fail")
    ipPort2_tmp = GetSetting(App.Title, "Settings", "PLCIPPort2", "Fail")
    If IsValidIPAddress(ipAddr_tmp) = False Then
        ipAddr_tmp = "192.168.0.2"
    End If
    If IsValidIPPort(ipPort1_tmp) = False Then
        ipPort1_tmp = "12001"
    End If
    If IsValidIPPort(ipPort2_tmp) = False Then
        ipPort2_tmp = "12002"
    End If
    
    With wsPLC1
        .Close
        .RemoteHost = ipAddr_tmp
        .RemotePort = ipPort1_tmp
        .LocalPort = ipPort1_tmp
        
        .Bind .LocalPort
    End With
    
    With wsPLC2
        .Close
        .RemoteHost = ipAddr_tmp
        .RemotePort = ipPort2_tmp
        .LocalPort = ipPort2_tmp
        
        .Bind .LocalPort
    End With
    
    Dim plcDataRangeMax_tmp As String
    plcDataRangeMax_tmp = GetSetting(App.Title, "Settings", "PLCDataRangeMax", "Fail")
    If IsValidValue(plcDataRangeMax_tmp, 100, 32767) = False Then
        plcDataRangeMax_tmp = "2047"
    End If
    PLCDataRangeMax = plcDataRangeMax_tmp
    
    For i = 0 To NUM_OF_BIN - 1
        ucBINdps1(i).setOptionD "0", "0.6", "0.5"
    Next i
    
    
    For i = 0 To NUM_OF_BIN - 1
        ucBINdps1(i).setBinID
        ''ucBINdps1(i).picCON_Cir1
        ''
        cboIDX.AddItem i + 1
    Next i

    
    For i = 0 To NUM_OF_BIN - 1
        BinTYPE(i) = GetSetting(App.Title, "Settings", "BINtype_" & Trim(i), 2590)
        '''''''
        '''SaveSetting App.Title, "Settings", "BINtype_" & Trim(i), BinTYPE(i)
    Next i
    '''
    For i = 0 To NUM_OF_BIN - 1
        ucBINdps1(i).setScanTYPE BinTYPE(i) ''' 211  '''LMS-211  '''LD-LRS-3100,, DPS-2590
    Next i

    ''ucBINdps1(0).setScanTYPE 2590  '''''LD-LRS-3100,, DPS-2590
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   

    For i = 0 To NUM_OF_BIN - 1
        BinMaxH(i) = GetSetting(App.Title, "Settings", "MaxH_" & Trim(i), 1200)
        BinMinH(i) = GetSetting(App.Title, "Settings", "MinH_" & Trim(i), 0)
    Next i
    
    For i = 0 To NUM_OF_BIN - 1
        ucBINdps1(i).set_maxHHLH CLng(BinMaxH(i)), CLng(BinMinH(i))  '''CLng(txtMaxHH)
        '''''''''''''''''''''''
        ucBINdps1(i).rxMode = 0  ''7
        '''''''''''''''''''''''
        ''ucBINmon1(i).runCONN
    Next i

    Dim BinAngleTmp$, SensorAngleTmp$

    For i = 0 To NUM_OF_BIN - 1
        BinAngleTmp = _
            GetSetting(App.Title, "Settings", "BinAngle_" & i, "Fail")
        SensorAngleTmp = _
            GetSetting(App.Title, "Settings", "SensorAngle_" & i, "Fail")
        If IsNumeric(BinAngleTmp) = False _
            Or CSng(CInt(Val(BinAngleTmp))) <> CSng(Val(BinAngleTmp)) _
            Or CInt(Val(BinAngleTmp)) > 10! Or CInt(Val(BinAngleTmp)) < -10! _
            Then
            BinAngleTmp = "0"
            SaveSetting App.Title, "Settings", "BinAngle_" & i, BinAngleTmp
        End If
        If IsNumeric(SensorAngleTmp) = False _
            Or CSng(CInt(Val(SensorAngleTmp))) <> CSng(Val(SensorAngleTmp)) _
            Or CInt(Val(SensorAngleTmp)) > 48! Or CInt(Val(SensorAngleTmp)) < -48! _
            Then
            SensorAngleTmp = "0"
            SaveSetting App.Title, "Settings", "SensorAngle_" & i, SensorAngleTmp
        End If
        ucBINdps1(i).setBinSettings CInt(BinAngleTmp), CInt(SensorAngleTmp)
    Next i
'''''''''''



    
'''    '''''''[TEST]''''LD-LRS-3100,, DPS-2590
'''    i = 20
'''    Load ucBINdps1(i)
'''    ''''''
'''        ucBINdps1(i).Top = 7600 ''4000  ''7600
'''        ''
'''        ucBINdps1(i).Width = BinWidth  '''Width / 11 - 50
'''        ucBINdps1(i).Left = ((i - 10) * (Width / 11)) + 20 ''+ 1720
'''        ucBINdps1(i).Height = 6100  '''12200
'''        ''
'''        ucBINdps1(i).Visible = True
'''        ''
'''        DoEvents
'''
'''        ucBINdps1(i).setScanTYPE 2590  '''''LD-LRS-3100,, DPS-2590
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''        ucBINdps1(i).setIDX (i), "192.168.0.11", "8282"  ''"192.168.0.21", "7001"
'''
'''        ucBINdps1(i).setOptionD "0", "0.6", "0.5"
'''
'''        BinMaxH(i) = GetSetting(App.Title, "Settings", "MaxH_" & Trim(i), 1850)
'''        SaveSetting App.Title, "Settings", "MaxH_" & Trim(i), BinMaxH(i)
'''        ''''
'''        ucBINdps1(i).set_maxHH CLng(BinMaxH(i))  '''CLng(txtMaxHH)
'''        ucBINdps1(i).setBinID
'''        ''
'''        ucBINdps1(i).rxMode = 0
        
        
    
    
    cboIDX.ListIndex = 0
    cboIDX.Refresh
    

    lbVS1.Caption = Screen.Width & "x" & Screen.Height
'' _
''                    & ", " & ucBINmon1(0).Width & "x" & ucBINmon1(0).Height _
''                    & ", " & ucBINmon1(0).picGET_width & "x" & ucBINmon1(0).picGET_height


    BINLog vbCrLf & vbCrLf & Format(Now, "YYYYMMDD-hh:mm:ss") & " ====[DIMMERS BIN-LEVEL START]===" & vbCrLf, "조광조"



    tmrINIT.Interval = 5000
    tmrINIT.Enabled = True
    
    ''txtMaxHH.Enabled = False
    

''        cmdRunStop.BackColor = &H80&    ''stop
''        cmdRunStop_Click                ''<<RUN>>''


End Sub

Private Sub Form_Terminate()
    ''Return
End Sub

Private Sub tmrAoDo_Timer()

Dim i As Integer
Dim j As Integer
Dim ioD(33) As Integer
Dim str1 As String
Dim str2 As String


Dim aaD(33) As Integer

Dim avrD(20) As Integer
Dim avrDsum(20) As Long

Dim aaH(33) As Integer

Dim avrH(20) As Integer
Dim avrHsum(20) As Long

Dim UDPiV_1(29) As Integer  '''[16bit-word] to PLC : now-Use-10/30word!
Dim UDPiV_2(29) As Integer  '''[16bit-word] to PLC : now-Use-10/30word!
    
    lbTimeNow.Caption = "RunTime: " & Format(Now, "YYYY-MM-DD h:m:s")

    For i = 0 To NUM_OF_BIN - 1
        aaD(i) = ucBINdps1(i).ret_AOd   '''' (1~32767)
        '''''''''''''''''''''''''''''
        aaH(i) = ucBINdps1(i).ret_Height
        ''''''''''''''''''''''''''''''''
    Next i
    
    ''Get--First!!
    For i = 0 To NUM_OF_BIN - 1
        If (aaD(i) <= 0) Or (aaD(i) >= 32768) Then
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<AVR)
    For i = 0 To NUM_OF_BIN - 1
        AOdeep(i, AOdeepCNT) = aaD(i)
        Hdeep(i, AOdeepCNT) = aaH(i)
    Next i
    ''
    AOdeepCNT = AOdeepCNT + 1
    ''
    If AOdeepCNT >= AOdeepMAX Then  ''99
        If AOdeepFull = False Then
            txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
            AOdeepFull = True
        End If
        AOdeepCNT = 0       ''''Loop!
    End If


    For i = 0 To NUM_OF_BIN - 1
        avrDsum(i) = 0
        avrHsum(i) = 0
    Next i
    
    ''//??????????
    If AOdeepFull = True Then
    ''
        For i = 0 To NUM_OF_BIN - 1
            For j = 0 To AOdeepMAX - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
                avrHsum(i) = avrHsum(i) + Hdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepMAX)
            avrH(i) = CInt(avrHsum(i) / AOdeepMAX)
        Next i
    ''
    ElseIf AOdeepCNT > 1 Then
    ''
        txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
        For i = 0 To NUM_OF_BIN - 1
            For j = 0 To AOdeepCNT - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
                avrHsum(i) = avrHsum(i) + Hdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepCNT)
            avrH(i) = CInt(avrHsum(i) / AOdeepCNT)
        Next i
    ''
    Else
        txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
        For i = 0 To NUM_OF_BIN - 1
            avrD(i) = aaD(i)
            avrH(i) = aaH(i)
        Next i
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>AVR)

    ''set_avrHH for View
    For i = 0 To NUM_OF_BIN - 1
        ucBINdps1(i).avrAOd = avrD(i)
        ucBINdps1(i).avrHeight = avrH(i)
    Next i

    
    ''Replace!!
    For i = 0 To NUM_OF_BIN - 1
        aaD(i) = avrD(i)
    Next i

    ''Ready for PLC direct
    For i = 0 To 29
         UDPiV_1(i) = 0
    Next i
    '''
    ' Convert 1 ~ 327671 to 1 ~ PLCDataRangeMax
    For i = 0 To NUM_OF_BIN - 1
        If (aaD(i) = 0) Then
            UDPiV_1(i) = 0
        Else
            UDPiV_1(i) = CLng(aaD(i) - 1) * (PLCDataRangeMax - 1) / (32767 - 1) + 1
        End If
    Next i
    ''
    ''SAVE--Replace!!
    For i = 0 To NUM_OF_BIN - 1
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 32767)  ''0
        End If
    Next i


        ioD(0) = aaD(0)  ''ucBINdps1(0).ret_AOd
        ioD(1) = aaD(1)  ''ucBINdps1(1).ret_AOd
        ioD(2) = aaD(2)  ''ucBINdps1(2).ret_AOd
        ioD(3) = 1 ''0

        ioD(4) = aaD(3)  ''ucBINdps1(3).ret_AOd
        ioD(5) = aaD(4)  ''ucBINdps1(4).ret_AOd
        ioD(6) = aaD(5)  ''ucBINdps1(5).ret_AOd
        ioD(7) = 1 ''0

        ioD(8) = 0 ''aaD(6)  ''ucBINdps1(6).ret_AOd
        ioD(9) = 0 ''aaD(7)  ''ucBINdps1(7).ret_AOd
        ioD(10) = 0 ''aaD(8)  ''ucBINdps1(8).ret_AOd
        ioD(11) = 1 ''0

        ioD(12) = 0 ''aaD(9)  ''ucBINdps1(9).ret_AOd
        ''''''''''''''''''''''''''''''''''''
        ioD(13) = ioD(0)
        ioD(14) = ioD(1)
        ioD(15) = 1

        ioD(16) = ioD(2)
        ioD(17) = ioD(4)
        ioD(18) = ioD(5)
        ioD(19) = 1

        ioD(20) = ioD(6)
        ioD(21) = ioD(8)
        ioD(22) = ioD(9)
        ioD(23) = 1

        ioD(24) = ioD(10)
        ioD(25) = ioD(12)
        ioD(26) = 1
        ioD(27) = 1
''''

        For i = 0 To 27 ''31
    ''--------------------------------------------------------(Temp)
    ''        If (ioD(i) > 0) And (ioD(i) <= 32767) Then
    ''            AOdata(i) = ioD(i)
    ''        Else
    ''            Exit Sub
    ''            ''=========>> Cancle for Next~~ /(protect_Zero_send)
    ''        End If
    ''--------------------------------------------------------(Temp)

            AOdata(i) = ioD(i)
            ''''''''''''''''''
        Next i


        If Len(txtSD1) > 6000 Then
            txtSD1 = Mid(txtSD1, 3000)
        End If
        txtSD1 = txtSD1 & vbCrLf & vbCrLf

        str1 = " <1> "
        For i = 0 To 12  ''31
            ''str1 = str1 & " [1-" & Format((i + 1), "00") & "]" & Format(AOdata(i), "00000")
            str1 = str1 & " [1-" & Format((i + 1), "00") & "]" & Format(CLng(AOdata(i)) * 100 / 32768, "00.0")
        Next i
        txtSD1 = txtSD1 & Format(Now, "YYYYMMDD-hh:mm:ss") & str1 '' & vbCrLf
        txtSD1.SelStart = Len(txtSD1)
        

        BINLog str1, "조광조"
        
    If (chkUsePLC = 1) Then
        Dim buffer(59) As Byte
        
        CopyMemory buffer(0), UDPiV_1(0), 30 * 2
        wsPLC1.SendData buffer
    End If

''''
End Sub


Private Sub swap(b1 As Byte, b2 As Byte)
 
  b1 = b1 Xor b2
  b2 = b1 Xor b2
  b1 = b1 Xor b2
 
End Sub


Private Sub tmrINIT_Timer()
    tmrINIT.Enabled = False
    
    Dim i
    
''    For i = 0 To 10
''        ucBINmon1(i).picCON_Cir1
''    Next i
'''''''''''''''
    ''ucBINmon1(0).picCON_Cir1
    

    tmrAoDo.Interval = 2000  '''3000  '''1000
    tmrAoDo.Enabled = True
    
End Sub


''Private Sub ucBINmon1_upDXY(Index As Integer)
''    ''ucBINmon1(Index).ret_SDXY
''End Sub


