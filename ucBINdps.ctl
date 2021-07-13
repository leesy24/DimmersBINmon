VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ucBINdps 
   Appearance      =   0  '평면
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  '투명
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   FillStyle       =   0  '단색
   ScaleHeight     =   9555
   ScaleWidth      =   1890
   Begin VB.Timer tmrHmax 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   600
   End
   Begin VB.TextBox txtHmin 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   0
      TabIndex        =   25
      Text            =   "500"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtTypes 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtHmax 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1260
      TabIndex        =   23
      Text            =   "2000"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdHmax 
      BackColor       =   &H00C0C000&
      Caption         =   "SET"
      Height          =   255
      Left            =   1080
      Style           =   1  '그래픽
      TabIndex        =   22
      Top             =   720
      Width           =   615
   End
   Begin VB.Timer tmrWDT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   480
      Top             =   4440
   End
   Begin VB.TextBox txtAOd 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      TabIndex        =   21
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtVV 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   450
      TabIndex        =   17
      Top             =   5160
      Width           =   525
   End
   Begin VB.TextBox txtAsum 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   0
      TabIndex        =   15
      Top             =   5160
      Width           =   465
   End
   Begin VB.TextBox txtAcnt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   450
      TabIndex        =   14
      Top             =   4920
      Width           =   525
   End
   Begin VB.TextBox txtRDmon 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrWS 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   4200
   End
   Begin VB.TextBox txtOpBot 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1200
      TabIndex        =   12
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtOpMid 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   720
      TabIndex        =   11
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtOpX 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox picXbar 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      FillColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      ScaleHeight     =   345
      ScaleWidth      =   1545
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAVRheight 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   4920
      Width           =   465
   End
   Begin VB.Timer tmrScan1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   4200
   End
   Begin VB.TextBox txtTime1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Text            =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtMode 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   1800
      Left            =   1080
      Top             =   4440
   End
   Begin VB.TextBox txtRXn 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   1080
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtRnn 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   600
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdCONN 
      BackColor       =   &H0000FF00&
      Caption         =   "BIN1"
      Height          =   255
      Left            =   120
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock wsock1 
      Left            =   0
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picCON 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      FillColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1545
      ScaleWidth      =   1545
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox picMON 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   1545
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox picGET 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      FillColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtBinAngle 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      TabIndex        =   26
      Text            =   "-10"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtSensorAngle 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1320
      TabIndex        =   27
      Text            =   "-48"
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lbHH 
      Appearance      =   0  '평면
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label lbVVV 
      Appearance      =   0  '평면
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      TabIndex        =   19
      Top             =   5760
      Width           =   885
   End
   Begin VB.Label lbAO 
      Appearance      =   0  '평면
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5760
      Width           =   825
   End
   Begin VB.Label lbHP 
      Appearance      =   0  '평면
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "V:0.0%"
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      TabIndex        =   16
      Top             =   5520
      Width           =   885
   End
End
Attribute VB_Name = "ucBINdps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const MAX_DEEPS_OF_BIN = 1200

Public Event Resize()

Public Event upDXY()


Private UCindex As Integer
Public BinName As String

Private swN As Integer

Private wsACT As Boolean

Private wsPause As Boolean


Public ipAddr As String
Public ipPort As String


Private Type POINTAPI
    x As Long
    y As Long
End Type
''''''''
Private pnt(200) As POINTAPI


Private handle As Long
Private ret1 As Long


Private inBUF(2000) As Byte
Private inCNT As Long

''Private rxMode As Integer
Public rxMode As Integer

Const PI = 3.14159265359   '''3.14159265358979  ''3.1415926535897932384626433832795

Private tmrRunWDTcnt As Integer


Private scanD(101) As Long
Private scanDfilt(101) As Long
''
Private scanD_backup(101) As Long  '''20170616


Private scanDX(101) As Long
Private scanDY(101) As Long
Private scanDXmin As Long
Private scanDXmax As Long
Private scanDYmin As Long


Public avrSUM As Double
Public avrCNT As Integer
Public avrMAX As Double
Public avrMIN As Double
Public avrAVR, avrMID As Double

Public maxHH As Long
Public minLH As Long
Public avrAOd As Integer
Public avrHeight As Integer

Private setAngle As Integer


Private startString As String   'Stores the string used to activate the sensor
Private stopString As String    'Store the string used to deactivate the sensor



Private INITed As Integer
Private picPASScnt As Integer

'''
Private ScanTYPE As Integer  '''DPS-2590 LMS-211 //  LD-LRS-3100,,
Public BinAngle%, SensorAngle%

Private inBUF2590 As String   '''inBUF2590(100000) As Byte

Private centerXsum As Double
Private centerXcnt As Integer

Private p_maxY As Double

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long


Public Function getScanTYPE() As Integer
    getScanTYPE = ScanTYPE
End Function


Public Sub setScanTYPE(iScan As Integer)  '''LD-LRS-3100,, DPS-2590
    ScanTYPE = iScan
    
    txtTypes = ScanTYPE  ''iScan
    
    SaveSetting App.Title, "Settings", "BINtype_" & Trim(UCindex), ScanTYPE
    ''SaveSetting

    If (txtTypes = 2590) Then
        txtTypes.BackColor = &HC0FFC0
    Else
        txtTypes.BackColor = &HFFFFC0
    End If

    
''''    If ScanTYPE = 2590 Then
''''            txtRxS.Visible = True
''''    Else
''''            txtRxS.Visible = False
''''    End If
    
End Sub


Public Sub setBinSettings(BinAngle_I%, SensorAngle_I%)
'
    BinAngle = BinAngle_I
    txtBinAngle.Text = BinAngle
'
    SensorAngle = SensorAngle_I
    txtSensorAngle.Text = SensorAngle
'
End Sub


Private Sub cmdHmax_Click()

    If cmdHmax.BackColor = vbGreen Then   ''&H00C0C000&
        cmdHmax.BackColor = &HC0C000
        txtHmax.Enabled = False
        txtHmin.Enabled = False
        txtTypes.Enabled = False
        
        tmrHmax.Enabled = False
        
        If (txtHmax <> maxHH) And ((txtHmax < MAX_DEEPS_OF_BIN * 0.8) Or (txtHmax > MAX_DEEPS_OF_BIN)) Then
            txtHmax = maxHH
        End If
        
        If (txtHmin <> minLH) And ((txtHmin < 0) Or (txtHmin > MAX_DEEPS_OF_BIN * 0.2)) Then
            txtHmin = minLH
        End If
        
        If (txtHmax <> maxHH) Or (txtHmin <> minLH) Then
            set_maxHHLH txtHmax, txtHmin
        End If
        
        If (txtTypes = 211) Or (txtTypes = 2590) Then
            ScanTYPE = txtTypes
            setScanTYPE ScanTYPE
        Else
            txtTypes = ScanTYPE
        End If
        
    Else
        cmdHmax.BackColor = vbGreen
        txtHmax.Enabled = True
        txtHmin.Enabled = True
        txtTypes.Enabled = True

        tmrHmax.Enabled = False
        tmrHmax.Interval = 60000 '' 60secs
        tmrHmax.Enabled = True

    End If
    
    If (txtTypes = 2590) Then
        txtTypes.BackColor = &HC0FFC0
    Else
        txtTypes.BackColor = &HFFFFC0
    End If
    
End Sub

Private Sub picXbar_Click()

'''    Dim maxyrange As Double                     'Sets max y range of Scan
'''    Dim minyrange As Double                     'Sets min y range of Scan
'''    Dim maxxrange As Double                     'sets max x range of scan
'''    Dim minxrange As Double                     'sets min x range of scan
'''    Dim angle(1 To 2000) As Double              'angle data
'''    Dim r(0 To 2000) As Double                  'radius data
'''    Dim x(0 To 2000) As Double                  'x - cartesian coordinate
'''    Dim y(0 To 2000) As Double                  'y - cartesian coordinate
'''    Dim n As Integer                            'number of data values
'''
'''    Dim minXL As Double
'''    Dim minXR As Double
    
        
    'Clears the previous scan plot
    picXbar.Cls


''    If picGET.Visible = True Then
''        picGET.Visible = False
''        picMON.Visible = True
''    Else
''        picGET.Visible = True
''        picMON.Visible = False
''    End If

End Sub


Private Sub picGET_Click()

    Dim maxyrange As Double                     'Sets max y range of Scan
    Dim minyrange As Double                     'Sets min y range of Scan
    Dim maxxrange As Double                     'sets max x range of scan
    Dim minxrange As Double                     'sets min x range of scan
    Dim angle(1 To 2000) As Double              'angle data
    Dim r(0 To 2000) As Double                  'radius data
    Dim x(0 To 2000) As Double                  'x - cartesian coordinate
    Dim y(0 To 2000) As Double                  'y - cartesian coordinate
    Dim n As Integer                            'number of data values
    
    Dim minXL As Double
    Dim minXR As Double
    
    'Clears the previous scan plot
    picGET.Cls
    'Set the scale for the plot in mm (starting upper left - lower right)
    maxyrange = 12000
    minyrange = 0
    maxxrange = 4000  ''' 7000 / 2 + 500
    minxrange = -4000  '''
    picGET.Scale (minxrange, maxyrange)-(maxxrange, minyrange)


    picGET.ForeColor = vbBlack  ''vbBlue
    ''
    
    picGET.Line (minxrange + 500, maxyrange - 500)-(minxrange + 500, maxyrange * Val(txtOpMid))
    picGET.Line (maxxrange - 500, maxyrange - 500)-(maxxrange - 500, maxyrange * Val(txtOpMid))
    ''
    picGET.Line (minxrange + 500, maxyrange * Val(txtOpMid))-(-600, minyrange - 100)
    picGET.Line (maxxrange - 500, maxyrange * Val(txtOpMid))-(600, minyrange - 100)
    ''picGET.Line (minxrange + 500, maxyrange * Val(txtOpMid))-((picGET.ScaleWidth * txtOpBot) - 3000 - 600, minyrange - 100)
    ''picGET.Line (maxxrange - 500, maxyrange * Val(txtOpMid))-((picGET.ScaleWidth * txtOpBot) - 3000 + 600, minyrange - 100)


    
    picGET.ForeColor = &HE0E0E0     ''vbCyan
        
    'Creates the axis lines and tic marks for the plot
    picGET.Line (0, 0)-(0, maxyrange)
    
    Dim t
    Dim k
    For t = 1 To 6
        picGET.Line ((1 / 50) * minxrange, (t / 6) * maxyrange)-((1 / 50) * maxxrange, (t / 6) * maxyrange)
        ''
        'Labeling the yaxis
        picGET.CurrentX = 0
        picGET.CurrentY = (t / 6) * maxyrange
        picGET.Print (t / 6) * (maxyrange / 1000) ''maxyrange
    Next t

    For t = -4 To 4  '''1 To 6
        picGET.Line (1000 * t, (1 / 50) * maxyrange)-(1000 * t, (1 / 50) * minyrange)
    Next t


    n = 100  ''361
    'Initialize starting values as zero
    x(1) = 0
    y(1) = 0
    minXL = 0
    minXR = 0
    
    For k = 0 To n
        x(k) = scanD(k) * Cos(((k) + 40 + BinAngle) * (PI / 180))  ''180
        ''x(k) = -x(k)
        x(k) = x(k) + Val(txtOpX.Text)
        
        ''y(k) = r(k) * Sin((angle(k) + 40 + BinAngle) * (PI / 180)) ''180
        y(k) = maxyrange - (scanD(k) * Sin(((k) + 40 + BinAngle) * (PI / 180)))  ''180
        
        If (x(k) > minxrange) And (x(k) < minXL) Then
            minXL = x(k)
        End If
        If (x(k) < maxxrange) And (x(k) > minXR) Then
            minXR = x(k)
        End If
    Next k
    
    centerXcnt = centerXcnt + 1
    If centerXcnt > frmMain.AOdeepMAX Then
        centerXsum = centerXsum - (centerXsum / centerXcnt)
        centerXcnt = frmMain.AOdeepMAX
    End If
    centerXsum = centerXsum + (minXR + minXL) / 2
    
    For k = 0 To n
        x(k) = x(k) - (centerXsum / centerXcnt)
    Next k

    For k = 0 To n
        'Draw lines between data points
        If k > 0 Then
            picGET.ForeColor = vbBlue  ''vbRed  ''vbCyan  ''vbBlack
            picGET.Line (x(k - 1), y(k - 1))-(x(k), y(k))
        End If
        
        'Plot the data points as circles
        If k < 2 Then
            picGET.ForeColor = vbRed
            picGET.Circle (x(k), y(k)), 150
        Else
            If k < 10 Or k > 90 Then
                picGET.ForeColor = vbYellow  ''vbMagenta  ''vbBlack
            Else
                picGET.ForeColor = vbCyan  ''vbYellow  ''vbMagenta  ''vbBlack
            End If
            picGET.Circle (x(k), y(k)), 100
        End If
    Next k

    '''''''''''''''''''''''''''''''''''''''
    For k = 0 To 100
        scanD_backup(k) = scanD(k)
    Next k
    '''''''''''''''''''''''''''''''''''''''
    
    Dim len1 As Long
    
    ''len1 = Abs(x(2)) + Abs(x(n))
    len1 = Abs(minXL) + Abs(minXR)
    ''
    ''picXbar.Cls  '''picXbar_Click
    '''''''''''''
    picXbar.ForeColor = vbWhite  ''vbCyan
    '''''''''''''
    picXbar.CurrentX = picXbar.Width / 2
    picXbar.CurrentY = 50
    picXbar.Print len1
    '''
    len1 = picXbar.Width * (len1 / (Abs(minxrange) + Abs(maxxrange)))
    minXL = (picXbar.Width - len1) / 2
    minXR = len1 + minXL
    picXbar.ForeColor = vbBlue  ''vbWhite  ''vbCyan
    picXbar.Line (minXL, picXbar.Height * 0.8)-(minXR, picXbar.Height * 0.8)
    '''
    
End Sub


Private Sub picMON_Click()
'''
    If (rxMode <> 7) Or (tmrRun.Enabled = False) Or (wsock1.State <> sckConnected) Then
    
        Exit Sub
        
    End If


    Dim maxyrange As Double                     'Sets max y range of Scan
    Dim minyrange As Double                     'Sets min y range of Scan
    Dim maxxrange As Double                     'sets max x range of scan
    Dim minxrange As Double                     'sets min x range of scan
'    Dim angle(1 To 2000) As Double              'angle data
'    Dim r(1 To 2000) As Double                  'radius data
    Dim x(0 To 200) As Double                  'x - cartesian coordinate
    Dim y(0 To 200) As Double                  'y - cartesian coordinate
     

    
    'Set the scale for the plot in mm (starting upper left - lower right)
    maxyrange = 20000
    minyrange = 0
    maxxrange = 10000
    minxrange = 0
    
    ''picMON.Scale (minxrange, maxyrange)-(maxxrange, minyrange)
    picMON.Scale (minxrange, maxyrange)-(maxxrange, minyrange)
    

Dim n
Dim k
''
Dim angle(200) As Double              'sets min x ran
Dim r(200) As Double                  'radius data

Dim minXL As Double
Dim minXR As Double
Dim maxY As Double

Dim X1, Y1, X2, Y2 As Double
    

''    n = 100  ''361
''    For k = 0 To n
''        Input #2, angle(k), r(k)
''
''        pnt(k).x = k + 10
''        pnt(k).y = (r(k) * Sin((angle(k) + 40 + BinAngle) * (PI / 180))) * 245 / 17850 + 50
''    Next k

    n = 100  ''361
    'Initialize starting values as zero
    x(1) = 0
    y(1) = 0
    minXL = 0
    minXR = 0
    maxY = 0

    For k = 0 To n
        x(k) = scanD(k) * Cos(((k) + 40 + BinAngle) * (PI / 180))  ''180
        ''x(k) = -x(k)
        x(k) = x(k) + Val(txtOpX.Text)
        
        ''y(k) = r(k) * Sin((angle(k) + 40 + BinAngle) * (PI / 180)) ''180
        ''y(k) = maxyrange - (scanD(k) * Sin(((k) + 40 + BinAngle) * (PI / 180)))  ''180
        y(k) = (scanD(k) * Sin(((k) + 40 + BinAngle) * (PI / 180)))   ''180
        
        x(k) = x(k) + 5000   '''BIN5::~3000
        ''''''''''''''''''
        
        pnt(k).x = x(k) / 100  ''about~(100 / 10000 = 1/100 = 0.01)
        pnt(k).y = y(k) / 100  ''about~(200 / 20000 = 2/200 = 0.01)
        
        If (x(k) < minXL) Then
            minXL = x(k)
        End If
        If (x(k) > minXR) Then
            minXR = x(k)
        End If
        If (k >= 10 And k <= 90 And y(k) > maxY) Then
            maxY = y(k)
        End If
    Next k
    If maxY = 0 Then
        maxY = p_maxY
    End If
    
    For k = 0 To n
        scanDX(k) = pnt(k).x  ''x(k)
        scanDY(k) = pnt(k).y  ''y(k)
    Next k
    minXL = minXL / 100
    minXR = minXR / 100
    scanDXmin = minXL
    scanDXmax = minXR

''===============================================================<<Fix>>!!
    pnt(101).x = 8  ''minXL ''- 1 ''6
    pnt(101).y = 15

    pnt(102).x = 8  ''minXL ''- 1 ''6
    pnt(102).y = 216 - (216 * txtOpMid)  ''100

    pnt(103).x = (maxxrange * 0.011) * txtOpBot - 5 ''49
    pnt(103).y = 216

    pnt(104).x = (maxxrange * 0.011) * txtOpBot + 5 ''59
    pnt(104).y = 216

    pnt(105).x = 100  ''minXR ''+ 1 ''102
    pnt(105).y = 216 - (216 * txtOpMid)  ''100
    ''
    pnt(106).x = 100  ''minXR ''+ 1 ''102
    pnt(106).y = 15
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''If (pnt(0).x < 90) And (txtAVRheight <> "") Then
    If (txtAVRheight <> "") Then
        If (txtAVRheight > 100) And (txtAVRheight < MAX_DEEPS_OF_BIN * 0.9) Then
            pnt(106).y = (txtAVRheight / MAX_DEEPS_OF_BIN) * 216
        End If
        If pnt(106).y > 100 Then
            pnt(106).y = 100
        End If
    End If
    ''
    pnt(107).x = pnt(0).x
    pnt(107).y = pnt(0).y
''===============================================================<<Fix>>!!


    'Clears the previous scan plot
    picMON.Cls
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    picMON.ForeColor = vbGreen  ''vbBlue  ''vbGreen ''&HE0E0E0  ''vbBlue  ''vbYellow  ''&HE0E0E0
    ''
    picMON.Line (minxrange + 500, maxyrange - 500)-(minxrange + 500, maxyrange * Val(txtOpMid))
    picMON.Line (maxxrange - 500, maxyrange - 500)-(maxxrange - 500, maxyrange * Val(txtOpMid))
    ''
    picMON.Line (minxrange + 500, maxyrange * Val(txtOpMid))-((picMON.ScaleWidth * txtOpBot) - 600, minyrange - 100)
    picMON.Line (maxxrange - 500, maxyrange * Val(txtOpMid))-((picMON.ScaleWidth * txtOpBot) + 600, minyrange - 100)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    picMON.FillStyle = vbSolid
    picMON.FillColor = &H80&  ''vbRed  ''&H404040     ''vbCyan
    picMON.ForeColor = &H40C0&      ''&HFF00FF   ''vbRed  ''vbGreen  ''vbRed  ''vbBlue  ''vbYellow  ''&HE0E0E0
    handle = picMON.hdc
    
    ''ret1 = Polygon(handle, pnt(0), 107)
    ''ret1 = Polygon(handle, pnt(0), 110)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''    ret1 = Polygon(handle, pnt(0), 108)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''
    ''''''''DoEvents
    ''''''''
    

''===============================================================
    
    picMON.ForeColor = vbCyan
    
''''=====================================================1
    avrSUM = 0
    avrCNT = 0
    avrMAX = 0
    avrMIN = maxY
    ''Dim d1, d2 As Double
    ''Dim s1, s2 As Double
    ''d1 = 0
    ''d2 = 0
    ''s1 = 0
    ''s2 = 0
    For k = 10 To 90
        ''d1 = Sqr((x(k - 1) - x(k)) ^ 2 + (y(k - 1) - y(k)) ^ 2)
        ''d2 = Sqr((x(k) - x(k + 1)) ^ 2 + (y(k) - y(k + 1)) ^ 2)
        ''s1 = Abs((y(k - 1) - y(k)) / (x(k - 1) - x(k)))
        ''s2 = Abs((y(k) - y(k + 1)) / (x(k) - x(k + 1)))
        If (Sqr((x(k - 1) - x(k)) ^ 2 + (y(k - 1) - y(k)) ^ 2) < 1000 And _
            Sqr((x(k) - x(k + 1)) ^ 2 + (y(k) - y(k + 1)) ^ 2) < 1000 And _
            Abs((y(k - 1) - y(k)) / (x(k - 1) - x(k))) <= 1.3 And _
            Abs((y(k) - y(k + 1)) / (x(k) - x(k + 1))) <= 1.3 _
            ) Then
            avrSUM = avrSUM + y(k)
            avrCNT = avrCNT + 1
            If (y(k) > avrMAX) Then
                avrMAX = y(k)
            End If
            If (y(k) < avrMIN) Then
                avrMIN = y(k)
            End If
        End If
    Next k
    txtAcnt.Text = avrCNT
    ''''''
    Dim avr1 As Double
    Dim avr12H As Double
    avr1 = 0
    ''''''
    avr12H = maxY
    If avrCNT >= 3 Then
        avrAVR = avrSUM / avrCNT
        avrMID = (avrMAX + avrMIN) / 2
        If ((avrMID / avrAVR) < 0.9) Then '' 0.85
        '' skip 10% of data on min
            avr1 = avrMIN + (avrMAX - avrMIN) * 10 / 100
        Else
            avr1 = avrMIN
        End If
        avrSUM = 0
        avrCNT = 0
        For k = 10 To 90
            If ((y(k) >= avr1) And _
                Sqr((x(k - 1) - x(k)) ^ 2 + (y(k - 1) - y(k)) ^ 2) < 1000 And _
                Sqr((x(k) - x(k + 1)) ^ 2 + (y(k) - y(k + 1)) ^ 2) < 1000 And _
                Abs((y(k - 1) - y(k)) / (x(k - 1) - x(k))) <= 1.3 And _
                Abs((y(k) - y(k + 1)) / (x(k) - x(k + 1))) <= 1.3 _
                ) Then
                avrSUM = avrSUM + y(k)
                avrCNT = avrCNT + 1
            End If
        Next k
        If avrCNT > 0 Then
            avr12H = avrSUM / avrCNT
        End If
    Else
        avrCNT = 0
    End If
    
    p_maxY = avr12H
    
    
    txtAcnt.Text = txtAcnt.Text & "," & avrCNT
    txtAVRheight.Text = CLng(avr12H / 10)
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    txtAsum.Text = CLng(MAX_DEEPS_OF_BIN - txtAVRheight.Text)   ''CLng(avrSUM)
    If txtAsum >= maxHH Then
        txtVV.Text = maxHH - minLH '' 100%
    ElseIf txtAsum <= minLH Then
        txtVV.Text = 0 '' 0%
    Else
        txtVV.Text = txtAsum - minLH
    End If
    
    txtAOd = CLng((txtVV / (maxHH - minLH)) * 32767)
    If txtAOd < 1 Then
        txtAOd = 1          '''v044~
    End If
    If txtAOd > 32767 Then
        txtAOd = 32767
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ''' H is the height of bin
    lbHH.Caption = "H:" & Format((avrHeight / 100), "#0.00")
    ''' I is the current in amperes
    lbAO.Caption = "I:" & Format(((avrAOd / 32768) * 16 + 4), "#0.00")  ''32768)
    ''' V is the volumn of bin
    If avrAOd >= 32767 Then
        lbHP.Caption = "V:" & "100%"
    ElseIf avrAOd <= 0 Then
        lbHP.Caption = "V:" & "0%"
    Else
        lbHP.Caption = "V:" & Format(avrAOd / 32767 * 100, "#0.0") & "%"
    End If
    ''' W is the weight of bin
    If txtOpMid >= 0.5 Then
        ''체적,중량:1~12:: 400[m*m*m]--520Ton
        lbVVV.Caption = "W:" & Format(avrAOd / 32767 * 300, "###0")   '''BIN5::400
    Else
        ''체적,중량: 8,9:: 150[m*m*m]--195Ton
        lbVVV.Caption = "W:" & Format(avrAOd / 32767 * 200, "###0")  '''BIN5::150
    End If


    RaiseEvent upDXY
    ''''''''''''''''

End Sub



Public Function ret_AOd() As Integer
    ret_AOd = Val(txtAOd)
End Function

Public Function ret_Height() As Integer
    '''ret_Height = Val(txtAsum)
    ret_Height = Val(txtVV)
End Function

Public Function ret_Act() As Integer
    If (wsACT = True) And (rxMode >= 7) Then
        ret_Act = 1
    Else
        ret_Act = 0
    End If
End Function

Public Function ret_HH() As Integer
    If lbHH <> "" Then
        ret_HH = CInt(Val(Mid(lbHH, 3) * 1000))
    Else
        ret_HH = 0
    End If
End Function

Public Function ret_VV() As Integer
    If lbVVV <> "" Then
        ret_VV = CInt(Val(Mid(lbVVV, 3)))
    Else
        ret_VV = 0
    End If
End Function

Public Function GETscanD(ang As Integer) As Long
    GETscanD = CLng(scanDfilt(ang))   '' / 10)
End Function

Public Sub set_maxHHLH(hh As Long, lh As Long)
    maxHH = hh
    minLH = lh
    ''''''''''
    txtHmax = hh
    txtHmin = lh
    
    
        SaveSetting App.Title, "Settings", "MaxH_" & Trim(UCindex), CInt(maxHH)
        SaveSetting App.Title, "Settings", "MinH_" & Trim(UCindex), CInt(minLH)
        ''SaveSetting
    
    
End Sub


Public Function ret_SDXY() As String

  Dim k As Integer
  Dim str1 As String
  
    ''str1 = ""
    str1 = "Xmin=" & scanDXmin & "  " & "Xmax=" & scanDXmax
    
    For k = 0 To 100
        If (k Mod 10) = 0 Then
            str1 = str1 & vbCrLf
        End If
        str1 = str1 & " [" & (k) & "]" & Str(scanDX(k)) & Str(scanDY(k))
    Next k
    
    ret_SDXY = str1
    
End Function



Private Sub tmrRun_Timer()

    Dim strA As String
    Dim data() As Byte
    
    txtTime1.Text = Format(Now, "ss")  ''' "hh:mm:ss")  ''' "YYYYMMDD h:m:s")

    If wsock1.State = sckConnected Then
        wsACT = True
        cmdCONN.BackColor = vbGreen
    Else
        wsACT = False
        cmdCONN.BackColor = vbRed
        centerXsum = 0
        centerXcnt = 0
    End If
    
    If txtMode = 7 Then
        txtMode.BackColor = &HFF8080
    Else
        txtMode.BackColor = vbRed
    End If
    


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''WDT
    tmrRunWDTcnt = tmrRunWDTcnt + 1
    If tmrRunWDTcnt > 10 Then  ''5
        wsock1.Close  '''(20170708)~
        ''''''''''''
        rxMode = 0
        tmrWS.Enabled = False   '''===>@tmrRun :: Restart!!!!!!!!!!
        tmrRunWDTcnt = 0
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''WDT
    
    If (rxMode = 0) And (tmrWS.Enabled = False) Then
        tmrWS.Interval = 1000
        tmrWS.Enabled = True
        
    ElseIf (rxMode = 1) And (wsock1.State = sckConnected) Then   ''re_1sec~
        If ScanTYPE = 211 Then
            wsock1.SendData stopString  ''scan_STOP
        End If
        
        If ScanTYPE = 2590 Then
            inCNT = 0
            inBUF2590 = ""
            rxMode = 7  '''==>RUN!!! DPS-2590
            Exit Sub  ''===============>>
        End If
    End If
    

    '''picCON_Cir1
    
    ''If (rxMode >= 7) And (wsPause = True) Then
    If (rxMode >= 7) And (INITed > 0) Then  ''<--201706

        If ScanTYPE = 211 Then
            ''''''''''''''''''''''''''''''''''''
            DoProcPIC   ''<===DPS-2590 !!!
            ''''''''''''''''''''''''''''''''''''
        End If
        
        If (ScanTYPE = 2590) And (wsock1.State = sckConnected) Then
            ''''''''''''''''''''''''''''''''''''
            strA = "SetAngle[" & SensorAngle & "]"
            
            data = StrConv(strA, vbFromUnicode)
            ''
            wsock1.SendData data
            ''''''''''''''''''''''''''''''''''''
        End If
        
    End If
    
End Sub


Private Sub DoProcPIC()

Dim DOavrSUM As Double
Dim DOavrCNT As Integer  ''Double
Dim DO_AVR As Integer    ''Double

''Public avrHeight As Long
''Public maxHH As Long

'''    ''''=====================================================1
'''        DOavrSUM = 0
'''        DOavrCNT = 0
'''        Dim k
'''        For k = 0 To 100
'''                DOavrSUM = DOavrSUM + scanD(k)  ''scanDY(k)
'''                DOavrCNT = DOavrCNT + 1
'''        Next k
'''        ''
'''        DO_AVR = DOavrSUM / DOavrCNT
'''
'''        If DO_AVR < 1000 Then
'''            picXbar.Cls  '''picXbar_Click
'''            '''''''''''''
'''            picXbar.ForeColor = vbRed  ''vbWhite  ''vbCyan
'''        Else
'''            picXbar.Cls  '''picXbar_Click
'''            '''''''''''''
'''            picXbar.ForeColor = vbBlack  ''vbRed  ''vbWhite  ''vbCyan
'''        End If
'''        picXbar.CurrentX = picXbar.Width / 20
'''        picXbar.CurrentY = 50
'''        picXbar.Print DO_AVR
'''    '''=====================================================1
'''
'''        If (DO_AVR < 1000) Then  '''500
'''            picPASScnt = picPASScnt + 1
'''        Else
'''            picPASScnt = 0
'''        End If
'''
'''        '''If (avrCNT > 500) Or ((avrCNT < 500) And (picPASScnt > 2)) Then
'''        If (DO_AVR < 1000) And (picPASScnt < 9) Then  ''3
'''
'''            '''picGET.Cls
'''            '''''''''''''''''''''''''''''''''''''''
'''            For k = 0 To 100
'''                scanD(k) = scanD_backup(k)
'''            Next k
'''            '''''''''''''''''''''''''''''''''''''''
'''        End If
'''    '''=====================================================1
    
        scanD_filt
        ''''''''''
        ''//''DoEvents
    
        picXbar.Cls
        
        picGET.Cls
        ''''''''''
        ''//''DoEvents
        
        picGET_Click
    
        picMON_Click
        ''''''''''''
        wsPause = False
        '''''''''''''''!!!
        
        ''//''DoEvents
        
End Sub



Private Sub tmrWS_Timer()
    
    rxMode = 1
    ''''''''''
    cmdCONN_Click
    '''''''''''''
    tmrWS.Enabled = False


    tmrScan1.Enabled = False
    tmrScan1.Interval = 9000  ''5000   '''wdt
    tmrScan1.Enabled = True


End Sub

Private Sub tmrScan1_Timer()  '''Resend~~~
'''
    If rxMode = 1 Then
    
    ElseIf rxMode = 2 Then
    
    ElseIf rxMode = 3 Then
                                ''ConfModeCM
    ElseIf rxMode = 4 Then
    
    ElseIf rxMode > 6 Then
    
    End If
    
    
    If rxMode < 7 Then
        wsock1.Close  '''(20170708)~
        ''''''''''''
        rxMode = 0              ''''ReStart!!!!
        tmrWS.Enabled = False   ''''(063)
    End If
    
    tmrScan1.Enabled = False
    
End Sub

Private Sub tmrWDT_Timer()
'''
'''    If tmrRun.Enabled = True Then  '''RUN-MODE'''
    
        wsock1.Close  '''(20170708)~
        ''''''''''''
        rxMode = 0
        tmrWS.Enabled = False   '''===>@tmrRun :: Restart!!!!!!!!!!
        
'''    End If
'''
End Sub

Private Sub tmrHmax_Timer()
'''
    tmrHmax.Enabled = False
    
    cmdHmax_Click
'''
End Sub


Private Sub UserControl_Initialize()

Dim i As Integer

    INITed = 0
    picPASScnt = 0

    tmrRunWDTcnt = 0

    txtMode.Height = 200
    txtRnn.Height = 200
    txtRXn.Height = 200
    txtTime1.Height = 200
    
    txtOpX.Height = 200
    txtOpMid.Height = 200
    txtOpBot.Height = 200
    
    cmdHmax.Height = 220
    txtHmax.Height = 250
    ''
    txtTypes.Height = 200
    
    
    
    inCNT = 0
    
    ''rxMode = 0
    ''''''''''''
    
    wsACT = False

    wsPause = False

    cmdCONN.BackColor = vbRed
    centerXsum = 0
    centerXcnt = 0

    ''''Sick:LMS-211''''
    startString = Chr(2) + Chr(0) + Chr(2) + Chr(0) + Chr(32) + Chr(36) + Chr(52) + Chr(8)
    stopString = Chr(2) + Chr(0) + Chr(2) + Chr(0) + Chr(32) + Chr(37) + Chr(53) + Chr(8)

End Sub

Private Sub UserControl_Resize()
    Dim i, d
    

    picGET.Width = Width - 60
    picGET.Left = 20
    picGET.Height = 3300  ''Height * 0.4
    picGET.Top = 1500  '''Height * 0.2
    
    picGET.Visible = True
    

    picXbar.Width = Width - 60
    picXbar.Left = 20
    picXbar.Height = 300
    picXbar.Top = picGET.Top - 320
    

    picMON.Width = Width - 60
    picMON.Left = 20
    picMON.Height = 3300  ''Height * 0.4
    picMON.Top = 1500  ''Height * 0.62
    ''
    picMON.Visible = False


    picCON.Width = Width - 60
    picCON.Left = 20
    picCON.Height = Width - 60
    picCON.Top = 1500  ''Height * 0.475
    ''
    picCON.Visible = False





    ''txtAVRheight



    
    RaiseEvent Resize
    
End Sub



Public Sub setBinID(BinName_I As String)
    BinName = BinName_I
    cmdCONN.Caption = UCindex + 1 & ") " & BinName
End Sub


Public Function getBinCaption() As String
    getBinCaption = cmdCONN.Caption
End Function


Public Sub setOptionD(dX As String, dM As String, dB As String)
    txtOpX.Text = Trim(dX)
    txtOpMid.Text = Trim(dM)
    txtOpBot.Text = Trim(dB)
End Sub


Public Sub setIDX(id As Integer, ip As String, port As String)
    
    UCindex = id
    
    ipAddr = ip
    ipPort = port
    
    wsock1.Close

    tmrRun.Enabled = True
    '''''''''''''''''''''
End Sub


Public Sub runCONN()
    cmdCONN_Click
End Sub


''''    void CLMSMANApp::ConfContiniousOutput(int Id)
''''    {
''''       BYTE telegram[8]={0x02,0x00,0x02,0x00,0x20,0x24,0x34,0x08};//Start
''''       LmsSendData((char *)&telegram, sizeof (telegram), Id);
''''    }
''''    void CLMSMANApp::StopContiniousOutput(int Id)
''''    {
''''       BYTE telegram[8]={0x02,0x00,0x02,0x00,0x20,0x25,0x35,0x08};//Stop
''''       LmsSendData((char *)&telegram, sizeof (telegram), Id);
''''    }

''''    ''''Sick:LMS-211''''
''''    startString = Chr(2) + Chr(0) + Chr(2) + Chr(0) + Chr(32) + Chr(36) + Chr(52) + Chr(8)
''''    stopString = Chr(2) + Chr(0) + Chr(2) + Chr(0) + Chr(32) + Chr(37) + Chr(53) + Chr(8)
    
Public Sub scan_STOP()
    If wsock1.State = sckConnected Then
        wsock1.SendData stopString
    End If
    
    tmrRun.Enabled = False
End Sub

Public Sub scan_RUN()

    Dim i As Integer
    For i = 0 To 100
        scanD(i) = 0
    Next i

    If wsock1.State = sckConnected Then
        wsock1.SendData startString
    End If
    
    wsPause = False
    '''''''''''''''
    
    tmrRun.Enabled = True
End Sub



''''    void CLMSMANApp::ConfBaudRate(int Id)   //Setting LMS Baud Rate
''''    {
''''       //DWORD dwBytesWriten;
''''       int mode = 0;  //1:19200, 2:38400
''''
''''       BYTE telegram[3][8]={
''''                            {0x02,0x00,0x02,0x00,0x20,0x42,0x52,0x08},  //9600 baud
''''                            {0x02,0x00,0x02,0x00,0x20,0x41,0x51,0x08},  //19200 baud
''''                            {0x02,0x00,0x02,0x00,0x20,0x40,0x50,0x08}   //38400 baud
''''                           };
''''
''''       LmsSendData((char *)&telegram[mode], sizeof (telegram[mode]), Id);
''''    }

''''    void CLMSMANApp::ConfModeCM(int Id)
''''    {
''''        BYTE telegram1[16]={0x02,0x00,0x0A,0x00,0x20,0x00,0x53,0x49,0x43,0x4B,0x5F,0x4C,0x4D,0x53,0xBE,0xC5};
''''
''''       LmsSendData((char *)&telegram1, sizeof (telegram1), Id);
''''    }

''''    void CLMSMANApp::ConfigAngleRes(int Id)
''''    {
''''       BYTE telegram[5][11]={
''''                                {0x02,0x00,0x05,0x00,0x3B,0x64,0x00,0x64,0x00,0x1D,0x0F},   //0도~100도 : 1도
''''                                {0x02,0x00,0x05,0x00,0x3B,0x64,0x00,0x32,0x00,0xB1,0x59},   //0도~100도 : 0.5도
''''                                {0x02,0x00,0x05,0x00,0x3B,0x64,0x00,0x19,0x00,0xE7,0x72},   //0도~100도 : 0.25도
''''                                {0x02,0x00,0x05,0x00,0x3B,0xB4,0x00,0x64,0x00,0x97,0x49},   //0도~180도 : 1도
''''                         {0x02,0x00,0x05,0x00,0x3B,0xB4,0x00,0x32,0x00,0x3B,0x1F}      //0도~180도 : 0.5도
''''                            };
''''
''''       LmsSendData((char *)&telegram[0], sizeof (telegram[0]), Id);
''''    }

Private Sub ConfBaudRate()
 Dim tele1 As String  ''{0x02,0x00,0x02,0x00,0x20,0x42,0x52,0x08},  //9600 baud
 tele1 = Chr(2) + Chr(0) + Chr(&O2) + Chr(0) + Chr(&H20) + Chr(&H42) + Chr(&H52) + Chr(&H8)
    LMSsendData tele1
End Sub

Private Sub ConfModeCM()
 Dim tele1 As String
 tele1 = Chr(2) + Chr(0) + Chr(&HA) + Chr(0) + Chr(&H20) + Chr(0) + Chr(&H53) + Chr(&H49) + Chr(&H43) + Chr(&H4B) + Chr(&H5F) + Chr(&H4C) + Chr(&H4D) + Chr(&H53) + Chr(&HBE) + Chr(&HC5)
    LMSsendData tele1
End Sub

Private Sub ConfigAngleRes()
 Dim tele1 As String
 tele1 = Chr(2) + Chr(0) + Chr(5) + Chr(0) + Chr(&H3B) + Chr(&H64) + Chr(0) + Chr(&H64) + Chr(0) + Chr(&H1D) + Chr(&HF)
    LMSsendData tele1
End Sub

Private Sub LMSsendData(sd As String)
    If wsock1.State = sckConnected Then
        wsock1.SendData sd
    End If
End Sub

Public Function picGET_width() As Integer
    picGET_width = picGET.Width
End Function

Public Function picGET_height() As Integer
    picGET_height = picGET.Height
End Function

Public Sub picCON_Cir1()

    picCON.ForeColor = vbRed  ''vbBlack
    picCON.Circle ((picCON.Width / 2) - 20, (picCON.Height / 4) - 20), (picCON.Width / 2) * 0.1 '' - 100 ''600


    picCON.ForeColor = vbCyan  ''vbRed  ''vbBlack
    
  If (UCindex <> 7) And (UCindex <> 8) Then
    picCON.Circle ((picCON.Width / 2) - 20, (picCON.Height / 2) - 20), (picCON.Width / 2) * 0.95 '' - 100 ''600
  End If


  If (UCindex >= 4) And (UCindex <= 6) Then
    picCON.ForeColor = vbBlue
    picCON.Circle ((picCON.Width / 2) - 20, (picCON.Height / 2) + 250), (picCON.Width / 2) * 0.2 '' - 100 ''600
    
  ElseIf (UCindex = 7) Or (UCindex = 8) Then
    picCON.ForeColor = vbBlack

''<가로>''
''    picCON.Line (40, (picCON.Height / 2) - 400)-(picCON.Width - 60, (picCON.Height / 2) - 400)
''    picCON.Line (40, (picCON.Height / 2) + 400)-(picCON.Width - 60, (picCON.Height / 2) + 400)
''    picCON.Line (40, (picCON.Height / 2) - 400)-(40, (picCON.Height / 2) + 400)
''    picCON.Line (picCON.Width - 60, (picCON.Height / 2) - 400)-(picCON.Width - 60, (picCON.Height / 2) + 400)
''
''    picCON.Line (640, (picCON.Height / 2) - 100)-(picCON.Width - 660, (picCON.Height / 2) - 100)
''    picCON.Line (640, (picCON.Height / 2) + 100)-(picCON.Width - 660, (picCON.Height / 2) + 100)
''    picCON.Line (640, (picCON.Height / 2) - 100)-(640, (picCON.Height / 2) + 100)
''    picCON.Line (picCON.Width - 660, (picCON.Height / 2) - 100)-(picCON.Width - 660, (picCON.Height / 2) + 100)
    
''<세로>''
    picCON.Line (picCON.Width / 2 - 400, (40))-(picCON.Width / 2 + 400, (40))
    picCON.Line (picCON.Width / 2 - 400, (picCON.Height - 60))-(picCON.Width / 2 + 400, (picCON.Height - 60))
    picCON.Line (picCON.Width / 2 - 400, (40))-(picCON.Width / 2 - 400, (picCON.Height - 60))
    picCON.Line (picCON.Width / 2 + 400, (40))-(picCON.Width / 2 + 400, (picCON.Height - 60))
    
    picCON.Line (picCON.Width / 2 - 100, (picCON.Height / 2 - 150))-(picCON.Width / 2 + 100, (picCON.Height / 2 - 150))
    picCON.Line (picCON.Width / 2 - 100, (picCON.Height / 2 + 150))-(picCON.Width / 2 + 100, (picCON.Height / 2 + 150))
    picCON.Line (picCON.Width / 2 - 100, (picCON.Height / 2 - 150))-(picCON.Width / 2 - 100, (picCON.Height / 2 + 150))
    picCON.Line (picCON.Width / 2 + 100, (picCON.Height / 2 - 150))-(picCON.Width / 2 + 100, (picCON.Height / 2 + 150))
        
  Else
    picCON.Circle ((picCON.Width / 2) + 250, (picCON.Height / 2) - 20), (picCON.Width / 2) * 0.2 '' - 100 ''600
    
  End If
    
    picCON.ForeColor = &HFFC0FF     ''&HFF00FF
    ''picCON.Line (40, (picCON.Height / 2) - 20)-(picCON.Width - 60, (picCON.Height / 2) - 20)
    picCON.Line (picCON.Width / 2 - 20, 40)-(picCON.Width / 2 - 20, (picCON.Height) - 60)
    
    ''체적,중량:1~12:: 400[m*m*m]--520Ton
    ''체적,중량: 8,9:: 150[m*m*m]--195Ton
    picCON.ForeColor = vbWhite  ''vbGreen

''<가로>''
''    If (UCindex = 7) Or (UCindex = 8) Then
''        picCON.CurrentX = picCON.Width / 2 - 300
''        picCON.CurrentY = picCON.Height - 400
''        picCON.Print "V:150"
''        picCON.CurrentX = picCON.Width / 2 - 300
''        picCON.CurrentY = picCON.Height - 240
''        picCON.Print "T:195"
''    Else
''        picCON.CurrentX = picCON.Width / 2 - 300
''        picCON.CurrentY = picCON.Height - 500
''        picCON.Print "V:400"
''        picCON.CurrentX = picCON.Width / 2 - 300
''        picCON.CurrentY = picCON.Height - 300
''        picCON.Print "T:520"
''    End If

''<세로>''
    If (UCindex = 7) Or (UCindex = 8) Then
        picCON.CurrentX = picCON.Width / 2 - 700
        picCON.CurrentY = picCON.Height - 900
        picCON.Print "V:150"
        picCON.CurrentX = picCON.Width / 2 - 700
        picCON.CurrentY = picCON.Height - 650
        picCON.Print "T:195"
    Else
        picCON.CurrentX = picCON.Width / 2 - 700
        picCON.CurrentY = picCON.Height - 900
        picCON.Print "V:400"
        picCON.CurrentX = picCON.Width / 2 - 700
        picCON.CurrentY = picCON.Height - 650
        picCON.Print "T:520"
    End If

End Sub


''Private Sub picCON_GotFocus()
''    picCON_Cir1
''End Sub


Private Sub cmdCONN_Click()

    inCNT = 0

    wsock1.Close

    wsock1.RemoteHost = ipAddr
    wsock1.RemotePort = ipPort
    
    wsock1.Connect

End Sub


Private Sub wsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''
    wsock1.Close
    
    rxMode = 0
    ''''''''''
End Sub


Private Sub wsock1_DataArrival(ByVal bytesTotal As Long)
'''


    Dim buffData As Variant 'This stores the incoming data from the buffer
    Dim i, j, c As Integer    'These are general counters
    Dim deg As Double       'This store the count for degrees
        
    Dim strHD(8) As Integer   'This stores the header to a packet of data
    
    
    Dim d1 As Long
    
    Dim Today
    Dim n1 As Integer
    
    Dim cCNT As Long
    Dim buff2590 As String   '''DPS-2590
    Dim scanN1 As Long
    
    
    'This is where the header for a data packet is assigned
    'NOTE: This changes based on settings of the SICK
    strHD(0) = &H2   ''2
    strHD(1) = &H80  ''128
    strHD(2) = &HCE  ''214(D6)
    strHD(3) = &H0  ''2
    strHD(4) = &HB0  ''176(B0)
    strHD(5) = &H65   ''105(69)
''
    strHD(6) = 65 ''(&h41)

    On Error GoTo exit1
'''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''(WDT)
    tmrWDT.Enabled = False
    tmrWDT.Interval = 10000  ''5000 '''10000
    tmrWDT.Enabled = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    tmrRunWDTcnt = 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If wsPause = True Then
        wsock1.GetData buffData
        ''//''DoEvents
        '''''''''''''''''''''''
        GoTo exit1  ''===============>>
    End If
    

    If ScanTYPE = 2590 Then  '''LD-LRS-3100,, DPS-2590

        wsock1.GetData buff2590$
        '''''''''''''''''''''''''

''        If rxSTOP > 0 Then
''            Exit Sub
''            '''''''''''''''''===>
''        End If


        txtMode.Text = rxMode
        txtRnn.Text = bytesTotal
        
        
        If (rxMode = 1) Then
            inCNT = 0
            inBUF2590 = ""
            rxMode = 7  '''==>RUN!!! DPS-2590
            Exit Sub  ''===============>>
        End If
        
    
        If (bytesTotal + inCNT) > 1999 Then
            inCNT = 0
            inBUF2590 = ""
            txtRXn.Text = inCNT
            Exit Sub  ''===============>>
        End If
    
    
        If (rxMode < 7) Then
            ''''            inCNT = 0
            ''''            inBUF2590 = ""
            ''''            rxMode = 7  '''==>RUN!!! DPS-2590
            Exit Sub  ''===============>>
        End If
    
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' $DPS2590_101,4F,4F,4F,4F,4F,4F,50,55,59,57,55,60,71,74,75,77,7A,EF,ED,EA,E7,E5,E3,E0,
        ''              DE,DC,DB,D9,D7,D5,D4,D2,D1,D0,CF,CE,CD,CD,CB,CA,CA,CA,C9,C8,C8,C8,C7,C7,
        ''              C6,C7,C7,C6,C7,C6,C7,C7,C7,C8,C8,C9,C9,CA,CB,CC,CC,CD,CF,D0,D1,D2,D3,D5,
        ''              D6,D7,D9,DA,DC,DF,E0,E3,E5,E7,EA,EC,EF,F1,F4,F8,FB,FE,102,106,10A,10E,
        ''              113,117,117,117,117,117,117
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        inBUF2590 = inBUF2590 & buff2590
        ''''''''''''''''''''''''''''''''
        inCNT = inCNT + bytesTotal
        ''''''''''''''''''''''''''
        If InStr(inBUF2590, "$DPS") > 0 Then
            inBUF2590 = Mid(inBUF2590, InStr(inBUF2590, "$DPS"))
            inCNT = Len(inBUF2590)
        End If
        ''''''''''''''''''''''''''
        txtRXn.Text = inCNT
        
        If (inCNT < 210) Then
            Exit Sub  ''===============>>
        End If
        
        ''' Check trailed CrLf and remove it
        If InStr(inBUF2590, vbCrLf) > 210 Then
            inBUF2590 = Left(inBUF2590, InStr(inBUF2590, vbCrLf) - 1)
            inCNT = Len(inBUF2590)  ''InStr(inBUF2590, vbCrLf)  ''Len(inBUF2590)
            txtRXn.Text = inCNT
            ''//''DoEvents
        Else
            Exit Sub  ''===============>>
        End If
        
        ''' Remove LF or CR if exist
        If InStr(inBUF2590, vbLf) > 210 Then
            inBUF2590 = Left(inBUF2590, InStr(inBUF2590, vbLf) - 1)
            inCNT = Len(inBUF2590)
            txtRXn.Text = inCNT
        End If
        If InStr(inBUF2590, vbCr) > 210 Then
            inBUF2590 = Left(inBUF2590, InStr(inBUF2590, vbCr) - 1)
            inCNT = Len(inBUF2590)
            txtRXn.Text = inCNT
        End If
        
        If (cmdHmax.BackColor = vbGreen) Then
            SaveStr2File UCindex & "_raw_", inBUF2590
        End If

        Dim strD1 As String
        n1 = 0
        inBUF2590 = Mid(inBUF2590, InStr(inBUF2590, ",") + 1)
        ''
        For i = 1 To inCNT - 1
            'calculate the distance measurement from the lower and upper byte
            
        
            If n1 > 100 Then
                Exit For  ''===>
            End If
            
        
            If (n1 = 100) And (Len(inBUF2590) >= 1) And (InStr(inBUF2590, ",") = 0) Then
                strD1 = inBUF2590
                ''''''''''''''''''''''''''''''''''''''''''''''''''
                scanN1 = Val("&H" & strD1) * 10
                scanD(n1) = scanN1
                
                If scanD(n1) < 600 Then
                    scanD(n1) = 499
                ElseIf scanD(n1) > 30000 Then  '''50000 ''Err~32767~~
                    scanD(n1) = 999
                End If
                
                Exit For  ''===>
            End If
        
            On Error GoTo exit1
        '''''''''''''''''''''''''''
            If InStr(inBUF2590, ",") > 1 Then
                strD1 = Left(inBUF2590, InStr(inBUF2590, ",") - 1)
                ''''''''''''''''''''''''''''''''''''''''''''''''''
            Else
                inCNT = 0
                Exit Sub  ''===============>> Err-Cancle!!
            End If
            
            
            scanN1 = Val("&H" & strD1) * 10
            '''''''''''''''''''''''''''''''
            scanD(n1) = scanN1
            
            If scanD(n1) < 600 Then
                scanD(n1) = 499
            End If
            
            If scanD(n1) > 30000 Then  '''50000 ''Err~32767~~
                scanD(n1) = 999
            End If

            n1 = n1 + 1
            
            inBUF2590 = Mid(inBUF2590, InStr(inBUF2590, ",") + 1)
            
        Next i



        ''scanD_filt
        ''''''''''
        
        '''picGET_Click
        ''''''''''''
        
        ''wsPause = True
        ''''''''''''''!!!
        
        INITed = 1  ''<---201706
        
        ''//''DoEvents
        ''''''''''''''''''''''''''''''''''''
        DoProcPIC   ''<===DPS-2590 !!!
        ''''''''''''''''''''''''''''''''''''
        
        inCNT = 0
        '''''''''


        '''inBUF2590 = ""

    
        ''//''DoEvents
        
        Exit Sub  '''''''===>
        
    Else
    
        wsock1.GetData buffData
        ''DoEvents
    
    End If
    
    
    ''''wsock1.GetData buffData
    
    
    



    txtRnn.Text = bytesTotal

    If (bytesTotal + inCNT) > 1999 Then
        inCNT = 0
        GoTo exit1  ''===============>>
    End If

    For i = 0 To bytesTotal - 1
        inBUF(inCNT + i) = buffData(i)
    Next i
    
    inCNT = inCNT + bytesTotal
    ''''''''''''''''''''''''''
    txtRXn.Text = inCNT

    txtMode.Text = rxMode



''-------------------------------------------------------
'''       if ( LmsMode[Id] == 1 && nSize == 24 )
'''       {
'''    //         LmsMode[Id] = 2;
'''    //         ConfigAngleRes(Id);
'''          return;
'''       }
'''       else if ( LmsMode[Id] == 1 && nSize > 24 )
'''       {
'''          StopContiniousOutput(Id);
'''          return;
'''       }
''-------------------------------------------------------
    ''If (rxMode = 1) And (bytesTotal = 24) Then
    ''Else
    If (rxMode = 1) And (inCNT > 24) Then
        If ScanTYPE = 211 Then
            wsock1.SendData stopString  ''scan_STOP
        End If
        inCNT = 0
        GoTo exit1  ''===============>>
    End If

   
''-------------------------------------------------------
'''   memcpy ( temp[Id]+nPos[Id], temp1, nSize);
'''   nPos[Id] += nSize;
'''   if ( LmsMode[Id] == 1 )
'''   {
'''      if ( nPos[Id] == 10 && temp[Id][nPos[Id]-1] == 0x0a )
'''      {  msg.Format("111111111111111");
'''         pClientDlg->AddListBox(msg);
'''
'''         memset( temp[Id], 0x00, nPos[Id] );
'''         nPos[Id] = 0;
'''         LmsMode[Id] = 2;
'''         ConfigAngleRes(Id);
'''      }
'''   }
''-------------------------------------------------------
    If (rxMode = 1) And (inCNT <= 10) And (inBUF(inCNT - 1) = &HA&) Then
        
        inCNT = 0
        
        If ScanTYPE = 2590 Then
            rxMode = 7  '''==>RUN!!! DPS-2590
            GoTo exit1  ''===============>>
        End If
        
        rxMode = 2
        ''''''''''
        ConfigAngleRes
        
        GoTo exit1  ''===============>>
        
''-------------------------------------------------------
'''   else if ( LmsMode[Id] == 2 )
'''   {
'''      msg.Format("111111111111111 nPos[%d] ID[%d]", nPos, Id);
'''      pClientDlg->AddListBox(msg);
'''
'''      //ConfigAngleRes Feedback
'''      if ( nPos[Id] == 14 && temp[Id][nPos[Id]-1] == 0xbd )
'''      {  msg.Format("222222222222222 ID[%d]", Id);
'''         pClientDlg->AddListBox(msg);
'''
'''         memset( temp[Id], 0x00, nPos[Id] );
'''         nPos[Id] = 0;
'''         LmsMode[Id] = 3;
'''         ConfModeCM(Id);
'''      }
'''   }
''-------------------------------------------------------
    ElseIf (rxMode = 2) And (inCNT = 14) And (inBUF(inCNT - 1) = &HBD&) Then
    ''//ConfigAngleRes Feedback
        
        inCNT = 0
        rxMode = 3 ''3
        ''''''''''
''        ConfBaudRate  ''ConfModeCM
        ConfBaudRate   ''ConfModeCM
        
        
        GoTo exit1  ''===============>>

''-------------------------------------------------------
'''   else if ( LmsMode[Id] == 3 )
'''   { //ConfModeCM Feedback
'''      if ( nPos[Id] == 10 && temp[Id][nPos[Id]-1] == 0x0a )
'''      {
'''         msg.Format("3333333333333333333 ID[%d]", Id);
'''         pClientDlg->AddListBox(msg);
'''
'''         memset( temp[Id], 0x00, nPos[Id] );
'''         nPos[Id] = 0;
'''         LmsMode[Id] = 4;
'''         ConfBaudRate(Id);
'''      }
'''   }
''-------------------------------------------------------
    ElseIf (rxMode = 3) Then
    ''//ConfModeCM Feedback
        
        ''If (inCNT <= 10) And (inBUF(inCNT - 1) = &HA) Then
        If (inBUF(inCNT - 1) = &HA) Then
        
            inCNT = 0
            rxMode = 4  ''5  ''4
            ''''''''''
            ConfBaudRate  ''ConfModeCM  ''ConfBaudRate

        End If
                
        GoTo exit1  ''===============>>
    
''-------------------------------------------------------
'''   else if ( LmsMode[Id] == 4 )
'''   { //ConfBaudRate Feedback
'''      if ( temp[Id][nPos[Id]-1] == 0x0A )
'''      {
'''         memset( temp[Id], 0x00, nPos[Id] );
'''         nPos[Id] = 0;
'''         LmsMode[Id] = 5;
'''         ConfBaudRate(Id);
'''      }
'''   }
''-------------------------------------------------------
'''   else if ( LmsMode[Id] == 5 )
'''   { //ConfBaudRate Feedback
'''      if ( temp[Id][nPos[Id]-1] == 0x0A )
'''      {
'''         memset( temp[Id], 0x00, nPos[Id] );
'''         nPos[Id] = 0;
'''         LmsMode[Id] = 6;
'''         ConfContiniousOutput(Id);
'''      }
'''   }
''-------------------------------------------------------
    ElseIf (rxMode = 4) And (inBUF(inCNT - 1) = &HA&) Then
    ''//ConfModeCM Feedback
        
        inCNT = 0
        rxMode = 5  '''5!!
        ''''''''''
        ConfBaudRate   ''ConfContiniousOutput
        
        GoTo exit1  ''===============>>

    ElseIf (rxMode = 5) And (inBUF(inCNT - 1) = &HA&) Then
    ''//ConfModeCM Feedback
        
        inCNT = 0
        rxMode = 6  '''5!!
        ''''''''''
        scan_RUN   ''ConfContiniousOutput
        
        GoTo exit1  ''===============>>
        

''-------------------------------------------------------
'''   else if ( LmsMode[Id] == 6 )
'''   { //ConfContiniousOutput Feedback
'''      if ( nPos[Id] >= 10 )
'''      {
'''         if ( temp[Id][9] == 0x0A )
'''         {
'''            memcpy( temp[Id], temp[Id]+10, nPos[Id]-10 );
'''            nPos[Id] -= 10;
'''
'''            LmsMode[Id] = 7;
'''         }
'''      }
'''   }
''-------------------------------------------------------
    ElseIf (rxMode = 6) And (inCNT > 9) Then
    ''//scan_RUN Feedback
        
        ''''''''''inCNT = 0
        rxMode = 7
        ''''''''''
        
        GoTo exit1  ''===============>>

    End If



''''''''''''''''''''''''''''        rxMode = 7
''''''''''''''''''''''''''''        ''''''''''




    If (inCNT < 212) Then

'        If (inCNT < 30) Then
'            txtRxD.Text = ""
'            For i = 0 To inCNT - 1
'                txtRxD.Text = txtRxD.Text & Hex(inBUF(i)) & " "
'            Next i
'        End If
        
''        If rxMode < 7 Then
''            inCNT = 0
''        End If
        

        GoTo exit1  ''===============>>
        
    End If
    
    
    

    i = 0   'This is the counter for the buffer
    c = 0   'This is the counter for the header string
    
    
''if ( (temp[Id][i+0] == 0x02) && (temp[Id][i+1] == 0x80) && (temp[Id][i+2] == 0xCE) && (temp[Id][i+3] == 0x00) &&
''//                 (temp[Id][i+4] == 0xB0) && (temp[Id][i+5] == 0x65) && (temp[Id][i+6] == 0x00) && (temp[Id][i+209] == 0x10) &&
''//                 (temp[Id][i+4] == 0xB0) && (temp[Id][i+5] == 0x65) && (temp[Id][i+209] == 0x10) &&
''     (temp[Id][i+4] == 0xB0) && (temp[Id][i+5] == 0x65) &&
''     nPos[Id] >= (i+212) )
''-------------------------------------------------------------------------------------------------------

    For i = 0 To inCNT - 1

        If (inBUF(i) = strHD(0)) And (inBUF(i + 1) = strHD(1)) And (inBUF(i + 2) = strHD(2)) And _
        (inBUF(i + 3) = strHD(3)) And (inBUF(i + 4) = strHD(4)) And (inBUF(i + 5) = strHD(5)) And _
        (inCNT >= i + 212) Then
        ''''''''''''''''''''''''''''''''''^^^^^^^^^^^^^^^^^
             
            n1 = 0
            For j = i + 7 To i + 7 + 200
                'calculate the distance measurement from the lower and upper byte
                
                d1 = ((CLng(inBUF(j + 1)) * CLng(256)) + CLng(inBUF(j))) * 10
                
                scanD(n1) = d1
                
                If d1 < 1000 Then
                    scanD(n1) = 499  ''20170616  '''1000
                End If
                
                If d1 > 50000 Then
                    scanD(n1) = 999  ''20170616  '''1000
                End If

                n1 = n1 + 1
                
                j = j + 1
            Next j
            
            
            ''scanD_filt
            ''''''''''
            
            '''picGET_Click
            ''''''''''''
            
            wsPause = True
            ''''''''''''''!!!
            
            INITed = 1  ''<---201706
            
            ''''''''''''''''''''''''''''''''''''
            '''DoProcPIC   ''<===DPS-2590 !!!
            ''''''''''''''''''''''''''''''''''''
            
            inCNT = 0
            '''''''''
            GoTo exit1  ''===============>>

        End If
        
        
        If (inCNT < i + 212) Then
            GoTo exit1  ''===============>>
        End If

    Next i


exit1:


End Sub



Private Sub scanD_filt()

Dim i, j As Integer

Dim Dcnt As Integer
Dim Dsum As Double
Dim DsumL As Long
Dim DsumM As Long
Dim DsumR As Long

    DsumM = 0
    Dcnt = 0
    For i = 0 To 15
        If (scanD(i) > 999) And (scanD(i) > DsumM) Then
            DsumM = scanD(i)
            Dcnt = Dcnt + 1
        End If
    Next i
    If Dcnt > 0 Then
        For i = 0 To 10
            If scanD(i) < 1000 Then
                scanD(i) = DsumM
            End If
        Next i
    End If
    
    Dcnt = 0
    For i = 10 To 40
        If (scanD(i) > 999) And (scanD(i) > DsumM) Then
            DsumM = scanD(i)
            Dcnt = Dcnt + 1
        End If
    Next i
    If Dcnt > 0 Then
        For i = 11 To 35
            If scanD(i) < 1000 Then
                scanD(i) = DsumM
            End If
        Next i
    End If
    
    Dcnt = 0
    For i = 30 To 70
        If (scanD(i) > 999) And (scanD(i) > DsumM) Then
            DsumM = scanD(i)
            Dcnt = Dcnt + 1
        End If
    Next i
    If Dcnt > 0 Then
        For i = 36 To 65
            If scanD(i) < 1000 Then
                scanD(i) = DsumM
            End If
        Next i
    End If
    
    Dcnt = 0
    For i = 63 To 92
        If (scanD(i) > 999) And (scanD(i) > DsumM) Then
            DsumM = scanD(i)
            Dcnt = Dcnt + 1
        End If
    Next i
    If Dcnt > 0 Then
        For i = 66 To 90
            If scanD(i) < 1000 Then
                scanD(i) = DsumM
            End If
        Next i
    End If
    
    Dcnt = 0
    For i = 86 To 100
        If (scanD(i) > 999) And (scanD(i) > DsumM) Then
            DsumM = scanD(i)
            Dcnt = Dcnt + 1
        End If
    Next i
    If Dcnt > 0 Then
        For i = 91 To 100
            If scanD(i) < 1000 Then
                scanD(i) = DsumM
            End If
        Next i
    End If
    
    '''201705 : Reverse!!! Dimmers!!!
    Dim scanDD(101) As Long
    For i = 0 To 100
        scanDD(i) = scanD(i)
    Next i

    ''''
    ''If (UCindex < 10) And (picPASScnt = 0) And (ScanTYPE = 211) Then  ''20170708~
    ''If (UCindex > 0) And (UCindex < 10) And (picPASScnt = 0) Then   ''20170708~
    If (UCindex < 10) And (picPASScnt = 0) Then   ''20170708~
        For i = 0 To 100
            scanD(i) = scanDD(100 - i)
        Next i
    End If
    '''''''''''''''''''''''''''''''''!!!

    For i = 0 To 100
        ''scanDfilt(i) = scanD(i)
        '''''''''''''''''''''''''(Law-Data!!)
        
        scanDfilt(i) = (scanD(i) * Sin(((i) + 40 + BinAngle) * (PI / 180)))
        
    Next i
    
End Sub






'###############################################################################
'##### User Control을 Main Map상에서 DragDrop하기 위해서 사용되는 Module
'###############################################################################
Private Sub UC_Click(Index As Integer)
    'UC(Index).ZOrder 0
    'RaiseEvent UCMouseClick(Index)
End Sub

Private Sub UC_GetFocus(Index As Integer)
    'RaiseEvent UCGotFocus(Index)
End Sub

Private Sub UC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'UC(Index).ZOrder 0
    'UC(Index).Drag 1
End Sub

Private Sub UC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'RaiseEvent UCMouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub UC_OutFocus(Index As Integer)
    'RaiseEvent UCLostFocus(Index)
End Sub

'###############################################################################



'###############################################################################
'###############################################################################
'###############################################################################



'''(end)'''


