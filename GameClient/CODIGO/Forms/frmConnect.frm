VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Eternal Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":1171AA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3030
      Left            =   2250
      TabIndex        =   3
      Top             =   5220
      Width           =   7500
      ExtentX         =   13229
      ExtentY         =   5345
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000000&
      Height          =   2175
      Left            =   6690
      TabIndex        =   2
      Top             =   2100
      Width           =   3060
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox PwdTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   11640
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Logiar 
      Height          =   615
      Left            =   4800
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image crearpj 
      Height          =   630
      Left            =   2280
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   2010
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************* ****************
'ImperiumAO - v1.0
'************************************************* ****************
'Copyright (C) 2015 Gaston Jorge Martinez
'Copyright (C) 2015 Alexis Rodriguez
'Copyright (C) 2015 Luis Merino
'Copyright (C) 2015 Girardi Luciano Valentin
'
'Respective portions copyright by taxpayers below.
'
'This library is free software; you can redistribute it and / or
'Modify it under the terms of the GNU General Public
'License as published by the Free Software Foundation version 2.1
'The License
'
'This library is distributed in the hope that it will be useful,
'But WITHOUT ANY WARRANTY; without even the implied warranty
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'************************************************* ****************
'
'************************************************* ****************
'You can contact me at:
'Gaston Jorge Martinez (Zenitram@Hotmail.com)
'************************************************* ****************

Option Explicit

'//////////Ahora no te podran robar la cuenta tan facil mente///////////

Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Integer, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub Command1_Click()

End Sub

Private Sub crearpj_Click()
frmCrearCuenta.Show
End Sub

Private Sub eND_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call EndGame
    End If
End Sub

Private Sub Form_Load()

    lServer(1).port = 7666
    lServer(1).Ip = "127.0.0.1"
    lServer(1).name = "Servidor Primario"
   
    lServer(2).port = 7666
    lServer(2).Ip = "127.0.0.1"
    lServer(2).name = "BattleServer #1 (Muy Pronto)"
    
    WebBrowser1.Navigate "http://www.imperiumao.com.ar/es"
   
    lst_servers.AddItem lServer(1).name
    lst_servers.AddItem lServer(2).name
    lst_servers.ListIndex = 1
    
End Sub



Private Sub imgGetPass_Click()

End Sub


Private Sub Image1_Click()
End
End Sub

Private Sub Logiar_Click()
If frmMain.Socket1.Connected Then
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    DoEvents
End If
    
UserAccount = NameTxt.Text
UserPassword = PwdTxt.Text

If Not UserAccount = "" And Not UserPassword = "" Then
    EstadoLogin = ConectarCuenta
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
End If



End Sub

Private Sub lst_servers_Click()
CurServerIp = lServer(lst_servers.ListIndex + 1).Ip
    CurServerPort = lServer(lst_servers.ListIndex + 1).port
End Sub

Private Sub PwdTxt_Change()
PwdTxt.Locked = True
Dim loopc As Byte, i As Byte
loopc = RandomNumber(3, 254)
i = 0
Do While loopc > i
    Call keybd_event(VkKeyScan(RandomNumber(32, 126)), 0, 0, 0)
    i = i + 1
Loop
DoEvents       ' IMPORTANTE: No quitar este DoEvents!
PwdTxt.Locked = False
End Sub

Function RandomNumber(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
Randomize Timer
RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
