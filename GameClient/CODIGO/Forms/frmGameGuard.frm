VERSION 5.00
Begin VB.Form frmGameGuard 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   Icon            =   "frmGameGuard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPres 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmGameGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\Recursos\Interface\gameguard.jpg")
    Unload frmLauncher
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FinGG = True
End Sub

Private Sub tmrPres_Timer()
'Static ticks As Long

'ticks = ticks + 1

'If ticks = 1 Then
    
'ElseIf ticks = 2 Then
    'Me.Picture = LoadPicture(App.Path & "\Recursos\Graficos\datafull.bmp")
'ElseIf ticks = 2 Then
'    Me.Picture = LoadPicture(App.Path & "\Recursos\Graficos\argentum.bmp")
'Else
FinGG = True
'End If

End Sub

