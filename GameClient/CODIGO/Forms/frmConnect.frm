VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      IntegralHeight  =   0   'False
      ItemData        =   "frmConnect.frx":4D282
      Left            =   6660
      List            =   "frmConnect.frx":4D284
      TabIndex        =   3
      Top             =   2400
      Width           =   3120
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2250
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox PwdTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2250
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3270
      Width           =   2355
   End
   Begin SHDocVwCtl.WebBrowser noticias 
      Height          =   2775
      Left            =   2250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4350
      Width           =   7530
      ExtentX         =   13282
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
   Begin VB.Image imgBorrarpj 
      Height          =   615
      Left            =   4155
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgOpciones 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image ImgCrearpj 
      Height          =   615
      Left            =   2220
      Top             =   7560
      Width           =   1785
   End
   Begin VB.Image imgconnect 
      Height          =   630
      Left            =   4770
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Image ImgPass 
      Height          =   615
      Left            =   6120
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Image imgGetPass 
      Height          =   615
      Left            =   6120
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   585
      Index           =   0
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   1770
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.1 MENDUZ DX8 VERSION www.noicoder.com
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit
Private cBotonCrearPj As ClsGraphicalButton
Private cBotonRecuperarPass As ClsGraphicalButton
Private cBotonBorrarPj As ClsGraphicalButton
Private cBotonConnect As ClsGraphicalButton
Private cBotonTeclas As ClsGraphicalButton

Public LastButtonPressed As ClsGraphicalButton
Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = 27 Then
        Call EndGame
    End If
End Sub

Private Sub Form_Load()

    Me.Picture = LoadPicture(App.path & "\Recursos\Interface\conectar.jpg")
    'Call noticias.Navigate("http://imperiumao.com.ar/es/")
    
   
    lServer(1).port = 7666
    lServer(1).Ip = "127.0.0.1"
    lServer(1).name = "Gotmul (Argentina)  " & "[" & Usuarios & "]"

    lst_servers.AddItem lServer(1).name
    
    LoadButtons
    EngineRun = False
End Sub

Private Sub ImgCrearPj_Click()
  Call Audio.Music_Load(48)
        
        EstadoLogin = E_MODO.Dados
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
End Sub

Private Sub imgconnect_Click()
    If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        
        'update user info
        UserName = NameTxt.text
        Dim aux As String
        aux = PwdTxt.text
        UserPassword = aux
        If CheckUserData(False) = True Then
            EstadoLogin = Normal
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
        End If
End Sub

Private Sub imgOpciones_Click()
frmOpciones.Show , frmConnect
End Sub

Private Sub ImgPass_Click()
frmNewPassword.Show , frmConnect
End Sub


Private Sub lst_servers_Click()
    CurServerIp = lServer(lst_servers.ListIndex + 1).Ip
    CurServerPort = lServer(lst_servers.ListIndex + 1).port
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = DirInterface
    
    Set cBotonCrearPj = New ClsGraphicalButton
    Set cBotonRecuperarPass = New ClsGraphicalButton
    Set cBotonBorrarPj = New ClsGraphicalButton
    Set cBotonConnect = New ClsGraphicalButton
    Set cBotonTeclas = New ClsGraphicalButton
    
    Set LastButtonPressed = New ClsGraphicalButton

        
    Call cBotonCrearPj.Initialize(ImgCrearpj, GrhPath & "imgCrearpj.jpg", _
                                               GrhPath & "botcrearover.jpg", _
                                    GrhPath & "botcreardown.jpg", Me)
                                  
    Call cBotonRecuperarPass.Initialize(ImgPass, GrhPath & "imgRecuperarContra.jpg", _
                                    GrhPath & "botrecuperarover.jpg", _
                                    GrhPath & "botrecuperardown.jpg", Me)
                                             
    Call cBotonBorrarPj.Initialize(imgBorrarpj, GrhPath & "ImgBorrarpj.jpg", _
                                    GrhPath & "botborrarover.jpg", _
                                    GrhPath & "botborrardown.jpg", Me)
                                    
                                    
    Call cBotonConnect.Initialize(imgconnect, GrhPath & "imgConnect.jpg", _
                                           GrhPath & "botconectarover.jpg", _
                                    GrhPath & "Botconectardown.jpg", Me)
                                    
    Call cBotonTeclas.Initialize(imgOpciones, GrhPath & "imgOpciones.jpg", _
                                      GrhPath & "botopcionesover.jpg", _
                                    GrhPath & "botopcionesdown.jpg", Me)

End Sub
Private Sub imgconnect_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_OVER)
End Sub
Private Sub imgpass_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_OVER)
End Sub
Private Sub imgopciones_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_OVER)
End Sub
Private Sub imgborrarpj_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_OVER)
End Sub
Private Sub imgCrearpj_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_OVER)
End Sub
Private Sub Form_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
