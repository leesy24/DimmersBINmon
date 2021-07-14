VERSION 5.00
Begin VB.Form frmCFG 
   BorderStyle     =   1  '단일 고정
   Caption         =   "설 정"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "센서 종류 설정"
      Height          =   2415
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   11175
      Begin VB.TextBox txtCtypes 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   310
         Width           =   615
      End
      Begin VB.CommandButton cmdSetTYPE 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lbBinNO2 
         Caption         =   "1) 조광조BIN-01"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "시스템 설정"
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   11175
      Begin VB.CheckBox chkUsePLC 
         Caption         =   "PLC 이용"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtPLCDataRangeMax 
         Enabled         =   0   'False
         Height          =   270
         Left            =   6840
         TabIndex        =   3
         Text            =   "32767"
         Top             =   675
         Width           =   615
      End
      Begin VB.TextBox txtPLCIPPort2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3240
         TabIndex        =   2
         Text            =   "99999"
         Top             =   1030
         Width           =   615
      End
      Begin VB.CommandButton cmdSetSYSTEM 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtPLCIPAddr 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3240
         TabIndex        =   0
         Text            =   "255.255.255.255"
         Top             =   310
         Width           =   1455
      End
      Begin VB.TextBox txtPLCIPPort1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3240
         TabIndex        =   1
         Text            =   "99999"
         Top             =   670
         Width           =   615
      End
      Begin VB.Label lbPLCDataRangeMax_ 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "(100~32767)"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbPLCDataRangeMax 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC Data range max."
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lbPLCIPAddr 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP addr."
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbPLCIPPort1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP port 1"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbPLCIPPort2 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP port 2"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCFGexit 
      Caption         =   "닫 기"
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Timer tmrCFG 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const NUM_OF_BIN = 6

Const TIMEOUT = 60000 ' 60secs

Dim isError_cmdSetSYSTEM As Boolean
Dim isError_cmdSetTYPE As Boolean

Private Sub chkUsePLC_KeyPress(KeyAscii As Integer)
'
    tmrCFG_update
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetSYSTEM_Update
    End If
'
End Sub

Private Sub chkUsePLC_Click()
'
    If (frmCFG.Visible = False) Then
        Exit Sub
    End If
'
    tmrCFG_update
'
    If (chkUsePLC.Value = 1) Then
        txtPLCIPAddr.Enabled = True
        txtPLCIPPort1.Enabled = True
        txtPLCIPPort2.Enabled = True
        txtPLCDataRangeMax.Enabled = True
    Else
        txtPLCIPAddr.Enabled = False
        txtPLCIPPort1.Enabled = False
        txtPLCIPPort2.Enabled = False
        txtPLCDataRangeMax.Enabled = False
    End If
'
End Sub

Private Sub cmdCFGexit_Click()
'
    tmrCFG.Enabled = False
'
    frmSettings.Visible = False
    frmCFG.Visible = False
'
    Unload frmSettings
    Unload Me
'
End Sub

Private Sub cmdSetSYSTEM_Click()
    Dim IsValid As Boolean
    Dim i
    
    tmrCFG_update
    
    isError_cmdSetSYSTEM = False
   
    'If (Val(txtAVRcnt) <> frmMain.AOdeepMAX) Then
    '    IsValid = True
    '
    '    If (Val(txtAVRcnt) < 10) Or (Val(txtAVRcnt) > 99) Then
    '        MsgBox lbAVRcnt & "는 10 이상 99 이하 이어야 합니다.", vbOKOnly
    '        IsValid = False
    '        isError_cmdSetSYSTEM = True
    '    End If
    '
    '    If (IsValid = True) Then
    '        SaveSetting App.Title, "Settings", "DeepMax", Val(txtAVRcnt)
    '        frmMain.AOdeepFull = False
    '        frmMain.AOdeepCNT = 0
    '        frmMain.AOdeepMAX = Val(txtAVRcnt)
    '    End If
    'End If
    
    If (chkUsePLC.Value <> frmMain.chkUsePLC) Then
        SaveSetting App.Title, "Settings", "UsePLC", chkUsePLC.Value
        frmMain.chkUsePLC = chkUsePLC.Value
    End If
    
    If (txtPLCIPAddr <> frmMain.wsPLC1.RemoteHost) Or _
       (txtPLCIPPort1 <> frmMain.wsPLC1.RemotePort) Or _
       (txtPLCIPPort2 <> frmMain.wsPLC2.RemotePort) Then
        IsValid = True
        
        If IsValidIPAddress(txtPLCIPAddr) = False Then
            MsgBox lbPLCIPAddr & "는 192.168.0.1 형태의 값 이어야 합니다.", vbOKOnly
            IsValid = False
            isError_cmdSetSYSTEM = True
        End If
        
        If IsValidIPPort(txtPLCIPPort1) = False Then
            MsgBox lbPLCIPPort1 & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
            IsValid = False
            isError_cmdSetSYSTEM = True
        End If
        
        If IsValidIPPort(txtPLCIPPort2) = False Then
            MsgBox lbPLCIPPort2 & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
            IsValid = False
            isError_cmdSetSYSTEM = True
        End If
        
        If txtPLCIPPort1 = txtPLCIPPort2 Then
            MsgBox lbPLCIPPort1 & "와 " & lbPLCIPPort2 & "는 서로 다른 값 이어야 합니다.", vbOKOnly
            IsValid = False
            isError_cmdSetSYSTEM = True
        End If
        
        If (IsValid = True) Then
            SaveSetting App.Title, "Settings", "PLCIPAddr", txtPLCIPAddr
            SaveSetting App.Title, "Settings", "PLCIPPort1", txtPLCIPPort1
            SaveSetting App.Title, "Settings", "PLCIPPort2", txtPLCIPPort2
        
            With frmMain.wsPLC1
                .Close
                .RemoteHost = txtPLCIPAddr
                .RemotePort = txtPLCIPPort1
                .LocalPort = txtPLCIPPort1
                .Bind .LocalPort
            End With
            With frmMain.wsPLC2
                .Close
                .RemoteHost = txtPLCIPAddr
                .RemotePort = txtPLCIPPort2
                .LocalPort = txtPLCIPPort2
                .Bind .LocalPort
            End With
        End If
    End If

    If (txtPLCDataRangeMax <> frmMain.PLCDataRangeMax) Then
        IsValid = True
        
        If IsValidValue(txtPLCDataRangeMax, 100, 32767) = False Then
            MsgBox lbPLCDataRangeMax & "는 100 ~ 32767 사이의 정수 값 이어야 합니다.", vbOKOnly
            IsValid = False
            isError_cmdSetSYSTEM = True
        End If
        
        If (IsValid = True) Then
            SaveSetting App.Title, "Settings", "PLCDataRangeMax", txtPLCDataRangeMax
            frmMain.PLCDataRangeMax = txtPLCDataRangeMax.Text
        End If
    End If
End Sub

Private Sub cmdSetTYPE_Click()
    Dim i
    
    tmrCFG_update
    
    isError_cmdSetTYPE = False
    
    For i = 0 To NUM_OF_BIN - 1
        If (CInt(txtCtypes(i).Text) <> frmMain.ucBINdps1(i).getScanTYPE) Then
            frmMain.ucBINdps1(i).setScanTYPE CInt(txtCtypes(i).Text)
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim TapIndex_base As Integer
    Dim iLeft As Long
    Dim iTop As Long

    'txtAVRcnt = frmMain.AOdeepMAX
    
    chkUsePLC.Value = frmMain.chkUsePLC
    If (chkUsePLC.Value = 1) Then
        txtPLCIPAddr.Enabled = True
        txtPLCIPPort1.Enabled = True
        txtPLCIPPort2.Enabled = True
        txtPLCDataRangeMax.Enabled = True
    End If
    
    txtPLCIPAddr.Text = frmMain.wsPLC1.RemoteHost
    txtPLCIPPort1.Text = frmMain.wsPLC1.RemotePort
    txtPLCIPPort2.Text = frmMain.wsPLC2.RemotePort
    
    txtPLCDataRangeMax.Text = frmMain.PLCDataRangeMax
    
    For i = 0 To NUM_OF_BIN - 1
        If i <> 0 Then
            Load lbBinNO2(i)
            Load txtCtypes(i)
            
            iLeft = lbBinNO2(0).Left + ((i) Mod 5) * (lbBinNO2(0).Width + txtCtypes(i).Width + 100)
            iTop = lbBinNO2(0).Top + ((i) \ 5) * 350
            
            lbBinNO2(i).Left = iLeft
            lbBinNO2(i).Top = iTop
            
            txtCtypes(i).Left = iLeft + lbBinNO2(0).Width
            txtCtypes(i).Top = iTop - 50
            txtCtypes(i).TabIndex = TapIndex_base + i
        Else
            TapIndex_base = txtCtypes(i).TabIndex
        End If
        
        lbBinNO2(i).Caption = frmMain.ucBINdps1(i).getBinCaption
        
        txtCtypes(i).Text = frmMain.ucBINdps1(i).getScanTYPE
    Next i

    For i = 0 To NUM_OF_BIN - 1
        lbBinNO2(i).Visible = True
        txtCtypes(i).Visible = True
    Next i
    
    tmrCFG_update
End Sub

Private Sub lbAVRcnt_Click()
'
    tmrCFG_update
'
End Sub

Private Sub lbBinNO2_Click(Index As Integer)
'
    tmrCFG_update
'
    If frmSettings.Visible = True Then
        frmSettings.Show
    End If
'
    frmSettings.Init _
        Index _
        , lbBinNO2(Index).Caption _
        , frmMain.ucBINdps1(Index).BinName _
        , frmMain.ucBINdps1(Index).ipAddr _
        , frmMain.ucBINdps1(Index).ipPort _
        , frmMain.ucBINdps1(Index).BinAngle _
        , frmMain.ucBINdps1(Index).SensorAngle _
        , frmMain.ucBINdps1(Index).maxHH _
        , frmMain.ucBINdps1(Index).minLH _
        , frmMain.ucBINdps1(Index).AOdeepMAX
'
    frmSettings.Visible = True
'
End Sub

Private Sub lbPLCDataRangeMax__Click()
'
    tmrCFG_update
'
End Sub

Private Sub lbPLCDataRangeMax_Click()
'
    tmrCFG_update
'
End Sub

Private Sub lbPLCIPAddr_Click()
'
    tmrCFG_update
'
End Sub

Private Sub lbPLCIPPort1_Click()
'
    tmrCFG_update
'
End Sub

Private Sub lbPLCIPPort2_Click()
'
    tmrCFG_update
'
End Sub

Private Sub tmrCFG_Timer()

    tmrCFG.Enabled = False
    
    frmSettings.Visible = False
    frmCFG.Visible = False
    
    Unload frmSettings
    Unload Me
    
End Sub

'Private Sub txtAVRcnt_KeyPress(KeyAscii As Integer)
''
'    If KeyAscii = 13 Then  ' The ENTER key.
'        cmdSetSYSTEM_Update
'    End If
''
'End Sub

'Private Sub txtAVRcnt_GotFocus()
''
'    tmrCFG_update
''
'End Sub

Private Sub txtCtypes_KeyPress(Index As Integer, KeyAscii As Integer)
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetTYPE_Update
    End If
'
End Sub

Private Sub txtCtypes_GotFocus(Index As Integer)
'
    tmrCFG_update
'
End Sub

Private Sub txtPLCDataRangeMax_KeyPress(KeyAscii As Integer)
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetSYSTEM_Update
    End If
'
End Sub

Private Sub txtPLCDataRangeMax_GotFocus()
'
    tmrCFG_update
'
End Sub

Private Sub txtPLCIPAddr_KeyPress(KeyAscii As Integer)
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetSYSTEM_Update
    End If
'
End Sub

Private Sub txtPLCIPAddr_GotFocus()
'
    tmrCFG_update
'
End Sub

Private Sub txtPLCIPPort1_KeyPress(KeyAscii As Integer)
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetSYSTEM_Update
    End If
'
End Sub

Private Sub txtPLCIPPort1_GotFocus()
'
    tmrCFG_update
'
End Sub

Private Sub txtPLCIPPort2_KeyPress(KeyAscii As Integer)
'
    If KeyAscii = 13 Then  ' The ENTER key.
        cmdSetSYSTEM_Update
    End If
'
End Sub

Private Sub txtPLCIPPort2_GotFocus()
'
    tmrCFG_update
'
End Sub

Public Sub tmrCFG_update()
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub

Private Sub cmdSetSYSTEM_Update()
    cmdSetSYSTEM_Click
    If (isError_cmdSetSYSTEM = False) Then
        SendKeys "{tab}"    ' Set the focus to the next control.
    End If
End Sub

Private Sub cmdSetTYPE_Update()
    cmdSetTYPE_Click
    If (isError_cmdSetTYPE = False) Then
        SendKeys "{tab}"    ' Set the focus to the next control.
    End If
End Sub


