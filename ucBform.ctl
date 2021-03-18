VERSION 5.00
Begin VB.UserControl ucBform 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BackStyle       =   0  '투명
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   ForeColor       =   &H00FF00FF&
   ScaleHeight     =   540
   ScaleWidth      =   690
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Shape           =   3  '원형
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "ucBform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit


Public Event Resize()


Private Sub UserControl_Resize()

    Shape1.FillColor = vbCyan

    Shape1.Width = Width - 60
    Shape1.Left = 20
    Shape1.Height = Width - 60
    Shape1.Top = 20
    
    Shape1.Refresh
    
    
    RaiseEvent Resize
        
End Sub


Public Sub reDO()
    UserControl_Resize
End Sub

