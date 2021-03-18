VERSION 5.00
Begin VB.UserControl ucBC 
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   FillStyle       =   0  '단색
   ScaleHeight     =   1740
   ScaleWidth      =   5325
   Begin VB.PictureBox picBC 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      FillColor       =   &H00FF00FF&
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdSW1a 
         BackColor       =   &H0000FF00&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         MaskColor       =   &H00FFFF00&
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdSW1b 
         BackColor       =   &H0000FF00&
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         MaskColor       =   &H00FFFF00&
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.Label uc_lbBC 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00400000&
         BackStyle       =   0  '투명
         Caption         =   "CV-"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape spBC 
         BackColor       =   &H00C00000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00000040&
         FillColor       =   &H00C00000&
         FillStyle       =   0  '단색
         Height          =   255
         Left            =   0
         Top             =   720
         Width           =   4455
      End
   End
End
Attribute VB_Name = "ucBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit



''Private MapFile As String

Public Event Resize()


Private UCindex As Integer
Private swN As Integer





Private Sub UserControl_Initialize()

Dim i As Integer



End Sub

Private Sub UserControl_Resize()
    Dim i, d
    
    picBC.Width = Width
    
    If UCindex = 1 Then  ''''''''''Only:CV-05A(REC)
        spBC.Width = Width * 0.95 ''- 700
    Else
        spBC.Width = Width ''- 30
    End If

    uc_lbBC.Left = Width / 2 - uc_lbBC.Width / 2

    If UCindex < 3 Then
                        d = Width * 0.04  ''600
                        spBC.BorderColor = &H4080&
    Else
                        d = 100
                        spBC.BorderColor = &HC00000
    End If
    
    cmdSW1a(0).Left = d
    cmdSW1b(0).Left = d
    
    For i = 1 To swN - 1

        cmdSW1a(i).Left = (Width * (1 / 14)) * i + d  ''''cmdSW1a(0).Left + (i * 920)
        cmdSW1b(i).Left = (Width * (1 / 14)) * i + d  ''''cmdSW1b(0).Left + (i * 920)
        
'        cmdSW1a(i).Move (Width * (1 / 14)) * i + d
'        cmdSW1b(i).Move (Width * (1 / 14)) * i + d

    Next i
    
'    spBC.Move spBC.Left
    
    
    
    
    

    RaiseEvent Resize
    
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

'''Public Sub ucResize()
'''    UserControl_Resize
'''End Sub


Public Sub LoadSW1(ByVal ucID As Integer, ByVal n As Integer, cap As String)

    On Error GoTo Err:
    
    Dim i As Integer
    Dim d As Integer
    
        UCindex = ucID
        ''''''''''''
        swN = n
        '''''''
        
        If ucID < 3 Then
            spBC.FillColor = &H404080
        Else
            spBC.FillColor = &HC00000
        End If
        
        uc_lbBC.Left = Width / 2 - uc_lbBC.Width / 2
        uc_lbBC.Top = uc_lbBC.Top + 25
        uc_lbBC.Height = uc_lbBC.Height - 68
        uc_lbBC.Caption = uc_lbBC.Caption & cap
                
        
        cap = ""  ''''<=='''' cap = cap & "-"
        
        d = 100  ''80
        cmdSW1a(0).Left = d
        cmdSW1b(0).Left = d
    
        cmdSW1a(0).Caption = cap & Trim(Str(1))
        cmdSW1b(0).Caption = cap & Trim(Str(n * 2))

        For i = 1 To n - 1
            Load cmdSW1a(i)
            cmdSW1a(i).Left = (Width * (1 / 14)) * i + d  ''''cmdSW1a(0).Left + (i * 920)
            cmdSW1a(i).Top = cmdSW1a(i).Top
            cmdSW1a(i).Caption = cap & Trim(Str(i + 1))
            cmdSW1a(i).Visible = True
            Load cmdSW1b(i)
            cmdSW1b(i).Left = (Width * (1 / 14)) * i + d  ''''cmdSW1b(0).Left + (i * 920)
            cmdSW1b(i).Top = cmdSW1b(i).Top
            cmdSW1b(i).Caption = cap & Trim(Str(n + (n - i)))  '' 14 + (14-1) =
            cmdSW1b(i).Visible = True
        Next i
    
    spBC.Top = spBC.Top + 10
    spBC.Height = spBC.Height - 40
    
    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub


Public Sub SetSW1(ByVal n As Integer, pic As String)

    On Error GoTo Err:
    
    Dim i As Integer

    i = swN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''LoadPicture()
    If i >= n Then
        cmdSW1a(n - 1).Picture = LoadPicture(pic)
    Else
        cmdSW1b(i + (i - n)).Picture = LoadPicture(pic)
    End If

    
    '''cmdSW1b(i).Picture = LoadPicture("net-b.bmp") '', vbLPCustom, vbLPColor, 32, 32)   ' Load cursor.

    
    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub



Public Sub SetSW_OFF(ByVal n As Integer, pic As String)

    On Error GoTo Err:
    
    Dim i As Integer

    i = swN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''LoadPicture()
    If i >= n Then
        cmdSW1a(n - 1).Picture = LoadPicture(pic)
        cmdSW1a(n - 1).BackColor = vbRed
    Else
        cmdSW1b(i + (i - n)).Picture = LoadPicture(pic)
        cmdSW1b(i + (i - n)).BackColor = vbRed
    End If

    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub



Public Sub SetSW_ON(ByVal n As Integer, pic As String)

    On Error GoTo Err:
    
    Dim i As Integer

    i = swN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''LoadPicture()
    If i >= n Then
        cmdSW1a(n - 1).Picture = LoadPicture(pic)
        cmdSW1a(n - 1).BackColor = vbGreen
    Else
        cmdSW1b(i + (i - n)).Picture = LoadPicture(pic)
        cmdSW1b(i + (i - n)).BackColor = vbGreen
    End If

    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub


Public Sub SetSW_BAT(ByVal n As Integer, pic As String)

    On Error GoTo Err:
    
    Dim i As Integer

    i = swN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''LoadPicture()
    If i >= n Then
        cmdSW1a(n - 1).Picture = LoadPicture(pic)
        cmdSW1a(n - 1).BackColor = vbYellow
    Else
        cmdSW1b(i + (i - n)).Picture = LoadPicture(pic)
        cmdSW1b(i + (i - n)).BackColor = vbYellow
    End If

    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub


Public Sub SetSW_Mode(ByVal n As Integer, mode As Integer)

    On Error GoTo Err:
    
    Dim i As Integer
    Dim cS As ColorConstants
    Dim pic As String

    Select Case mode
        Case 1
                ''sMode = "[비상정지]": img1 = "sw6b.bmp"
                cS = vbRed
                pic = "sw6b.bmp"
                
        Case 2
                ''sMode = "[정상복귀]": img1 = "sw6a.bmp"
                cS = vbGreen
                pic = "sw6a.bmp"
                
        Case 3
                ''sMode = "[시험수신]": img1 = "sw6d.bmp"
                cS = vbBlue  ''vbInactiveBorder
                pic = "sw6d.bmp"
        
        Case 4
                ''sMode = "[LOW-BATT]": img1 = "sw6c.bmp"
                cS = vbYellow
                pic = "sw6c.bmp"
                
    End Select

    i = swN
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''LoadPicture()
    If i >= n Then
        ''cmdSW1a(n - 1).Picture = LoadPicture(pic)
        'cmdSW1a(n - 1).Picture = LoadPicture("")
        cmdSW1a(n - 1).BackColor = cS
        'cmdSW1a(n - 1).MaskColor = cS
    Else
        ''cmdSW1b(i + (i - n)).Picture = LoadPicture(pic)
        'cmdSW1b(n - 1).Picture = LoadPicture("")
        cmdSW1b(i + (i - n)).BackColor = cS
        'cmdSW1b(i + (i - n)).MaskColor = cS
    End If

    UserControl_Resize
    
    Exit Sub
    
Err:

End Sub



Public Sub SetSW_init()
Dim i As Integer

    For i = 0 To swN - 1
    
        ''cmdSW1a(i).Picture = LoadPicture("sw6a.bmp")
        cmdSW1a(i).Picture = LoadPicture("")
        cmdSW1a(i).BackColor = vbGreen
    
    Next i
    
End Sub




Private Sub cmdSW1a_Click(Index As Integer)
'''
    retSW (Index + 1)
    
End Sub


Private Sub cmdSW1b_Click(Index As Integer)
'''
    retSW (swN + (swN - Index))  ''(14+(14-0))=28, (14+(14-1))=27,..,(14+14-13)=15
    
End Sub


'''<<정상복귀>>'''
Private Sub retSW(id As Integer)  ''''(1~26:28) SWid '''

    Dim cS As Long
    
    If swN >= id Then
        cS = cmdSW1a(id - 1).BackColor
    Else
        ''cS = cmdSW1b(swN + (swN - id)).BackColor
        cS = cmdSW1b(swN - (id - swN)).BackColor '' (14-(28-14))=0, (14-(27-14))=1,..,(14-(15-14))=13
    End If

    If cS = vbGreen Then
        Exit Sub ''''''''-->>
    End If



    Dim Msg, Style, Title, Help, Ctxt, Response
    
    Msg = "풀코드 상태를 복귀하시겠습니까?"   ' 기본 메시지.
    Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
    Title = "수동/강제 복귀!"   ' 기본 제목.
    Help = ""                   ''"DEMO.HLP"   ' 기본 도움말 파일.
    Ctxt = 1000                 ' 기본 항목 ' 구문.  ' 메시지 화면 표시.
    
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then        ' 사용자가 예를 선택.
            SetSW_Mode id, 2
            ''''''''''''''''(복귀)
    Else                            ' 사용자가 아니오를 선택.
            MsgBox "수동복귀 취소, 기존상태를 유지합니다."
    End If

End Sub



'''(end)'''

