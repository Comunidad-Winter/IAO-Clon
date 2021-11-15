VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   4875
   ClientTop       =   2130
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   9
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   8
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   7
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   3165
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Picmacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   10200
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   12
      Top             =   7380
      Width           =   1455
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   600
         Shape           =   3  'Circle
         Top             =   600
         Width           =   60
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   9000
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   8
      Top             =   2220
      Width           =   2415
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8865
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Renderer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   210
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   6
      Top             =   2070
      Width           =   8175
      Begin VB.Timer macrotrabajo 
         Enabled         =   0   'False
         Left            =   6360
         Top             =   2040
      End
      Begin VB.Timer trueno 
         Enabled         =   0   'False
         Interval        =   2
         Left            =   2040
         Top             =   1200
      End
      Begin VB.Timer second 
         Left            =   3480
         Top             =   1200
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5790
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   210
      MaxLength       =   500
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":36A50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Imgquest 
      Height          =   450
      Left            =   9270
      Top             =   3765
      Width           =   1890
   End
   Begin VB.Image cmdmensaje 
      Height          =   255
      Left            =   7800
      Top             =   1740
      Width           =   555
   End
   Begin VB.Image Nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":36B14
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image Modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":36F52
      ToolTipText     =   "Modo combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9360
      TabIndex        =   19
      Top             =   8580
      Width           =   345
   End
   Begin VB.Label LblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9360
      TabIndex        =   18
      Top             =   8355
      Width           =   345
   End
   Begin VB.Image imggrupo 
      Height          =   450
      Left            =   9270
      Top             =   2010
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdcerrar 
      Height          =   225
      Left            =   11580
      Top             =   180
      Width           =   255
   End
   Begin VB.Image cmdminimizar 
      Height          =   225
      Left            =   11340
      Top             =   180
      Width           =   225
   End
   Begin VB.Image imgminicerra 
      Height          =   315
      Left            =   11325
      Top             =   150
      Width           =   510
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   17
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   16
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   15
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   14
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   13
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Image PicSeg 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":37390
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   2
      Left            =   10740
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2595
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   2
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3180
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1890
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
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
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   10
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8610
      TabIndex        =   9
      Top             =   180
      Width           =   2625
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   1
      Left            =   9660
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   0
      Left            =   8580
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image CMDInfo 
      Height          =   390
      Left            =   10650
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   945
   End
   Begin VB.Image cmdLanzar 
      Height          =   390
      Left            =   8775
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   0
      Left            =   11475
      Top             =   3405
      Width           =   300
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   1
      Left            =   11460
      Top             =   2910
      Width           =   300
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Top             =   1230
      Width           =   3240
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   3
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10275
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11985
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblexp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   8820
      TabIndex        =   1
      Top             =   885
      Width           =   1815
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8820
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10920
      TabIndex        =   0
      Top             =   840
      Width           =   435
   End
   Begin VB.Menu mhabla 
      Caption         =   "Mhabla"
      Visible         =   0   'False
      Begin VB.Menu mNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu Mgrupo 
         Caption         =   "Grupo"
      End
      Begin VB.Menu Mgms 
         Caption         =   "Gms"
      End
      Begin VB.Menu mGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mClan 
         Caption         =   "Clan"
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.1 MENDUZ DX8 VERSION www.noicoder.com
'nqe onda, mostrame que estabas haciendo y que te pasa lo habia compilado igul eh
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

Option Explicit

Public InMouseExp As Boolean
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long
Public Index As Integer
Public IsPlaying As Byte
Dim UltimoIndex As Integer
Public UltPos As Integer
Public UltPosInterface As Integer
Public UltPosSolapas As Integer

Private CentroActual As Byte


Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.ListIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case Index
        Case 1 'subir
            If hlst.ListIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index
        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
    End Select
End Sub

Private Sub lblExp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
InMouseExp = True
lblexp.Caption = UserExp & "/" & UserPasarNivel
If UserPasarNivel = 0 Then
    lblexp.Caption = "¡Nivel máximo!"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    'Audio.MusicActivated = Not Audio.MusicActivated
                    
                 Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                     Dim i As Integer
Clipboard.Clear
DoEvents
Call keybd_event(VK_SNAPSHOT, PS_TheScreen, 0, 0)
DoEvents
For i = 1 To 1000
        If Not FileExist(App.path & "\Screenshots\foto" & i & ".bmp", vbNormal) Then Exit For
Next
SavePicture Clipboard.GetData, App.path & "\screenshots\foto" & i & ".jpg"
Call AddtoRichTextBox(frmMain.RecTxt, "Screenshot guardada en " & App.path & "\Screenshots\foto" & i & ".bmp !", 255, 150, 50, False, False, False)
     
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                   frmMain.Modocombate.Visible = Not frmMain.Modocombate.Visible
                   frmMain.Nomodocombate.Visible = Not frmMain.Nomodocombate.Visible
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call WriteWork(eSkill.Domar)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call WriteWork(eSkill.Robar)
                            
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call WriteWork(eSkill.Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    
                    
            End Select
        Else
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    If LenB(CustomMessages.Message((KeyCode - 39) Mod 10)) <> 0 Then
                        Call WriteTalk(CustomMessages.Message((KeyCode - 39) Mod 10))
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
    Case vbKeyF12
Clipboard.Clear
DoEvents
Call keybd_event(VK_SNAPSHOT, PS_TheScreen, 0, 0)
DoEvents
For i = 1 To 1000
        If Not FileExist(App.path & "\Screenshots\foto" & i & ".bmp", vbNormal) Then Exit For
Next
SavePicture Clipboard.GetData, App.path & "\screenshots\foto" & i & ".jpg"
Call AddtoRichTextBox(frmMain.RecTxt, "Screenshot guardada en " & App.path & "\Screenshots\foto" & i & ".bmp !", 255, 150, 50, False, False, False)

            
     
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmestadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
         'Accion macros extracted imperiumAO 1.4
     Case vbKeyF1
       BotonElegido = 1
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
    Case vbKeyF2
       BotonElegido = 2
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
    Case vbKeyF3
       BotonElegido = 3
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
     Case vbKeyF4
       BotonElegido = 4
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
      Case vbKeyF5
       BotonElegido = 5
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
      Case vbKeyF6
       BotonElegido = 6
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
       Case vbKeyF7
       BotonElegido = 7
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
       Case vbKeyF8
       BotonElegido = 8
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
       Case vbKeyF9
       BotonElegido = 9
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
        Case vbKeyF10
       BotonElegido = 10
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
       Case vbKeyF11
       BotonElegido = 11
       If MacroKeys(BotonElegido).TipoAccion = 0 Then
        frmbindkey.Show vbModeless, frmMain
       Else
        Call frmbindkey.Bind_Accion(BotonElegido)
       End If
      'Fin accion macros
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG
            
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            Call WriteMeditate

        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
                        If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmestadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        Call EndGame
        Cancel = 1
    End If
End Sub


Private Sub imgCentros_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltPosSolapas = Index Then Exit Sub

If UltPosSolapas <> -1 Then Call RestaurarCentroActual
UltPosSolapas = Index

Select Case Index
    Case 0 'Inv
        imgCentros(0).Picture = LoadPicture(App.path & "\Recursos\Interface\[solapas]inventario-over.jpg")
    Case 1 'Hechizos
        imgCentros(1).Picture = LoadPicture(App.path & "\Recursos\Interface\[solapas]hechizos-over.jpg")
    Case 2 'Menu
        imgCentros(2).Picture = LoadPicture(App.path & "\Recursos\Interface\[solapas]menu-over.jpg")
End Select

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
  '  Case 0 'Grupo
     '   Image1(0).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]grupo-down.bmp")
    Case 1 'Estadisticas
        Image1(1).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]estadisticas-down.jpg")
    Case 2 'Guild
        Image1(2).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]clanes-down.jpg")
   ' Case 3 'Quest
   '     Image1(3).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]quests-down.bmp")
   ' Case 4 'Torneos
    '    Image1(4).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]torneos-down.bmp")
    Case 5 'Opciones
        Image1(0).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]opciones-down.jpg")
End Select

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index

    'Case 0 'Grupo
       ' Image1(0).Picture = General_Load_Picture_From_Resource("[menu]grupo-over.bmp")
    Case 1 'Estadisticas
        Image1(1).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]estadisticas-over.jpg")
    Case 2 'Guild
        Image1(2).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]clanes-over.jpg")
    'Case 3 'Quest
     '   Image1(3).Picture = General_Load_Picture_From_Resource("[menu]quests-over.bmp")
    'Case 4 'Torneos
     '   Image1(4).Picture = General_Load_Picture_From_Resource("[menu]torneos-over.bmp")
    Case 5 'Opciones
        Image1(0).Picture = LoadPicture(App.path & "\Recursos\Interface\[menu]opciones-over.jpg")
End Select

End Sub

Private Sub imgCentros_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        
        Case 0
            InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centroinventario.jpg")

            picInv.Visible = True
            
            lblInvInfo.Visible = True
            lblInvInfo = ""
        
            hlst.Visible = False
            CMDInfo.Visible = False
            cmdLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            cmdMoverHechi(0).Enabled = False
            cmdMoverHechi(1).Enabled = False
        
            Image1(0).Visible = False
            Image1(1).Visible = False
            Image1(2).Visible = False
            imggrupo.Visible = False
        Case 1
            InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centrohechizos.jpg")
            '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
            
            picInv.Visible = False
            lblInvInfo.Visible = False
            
            hlst.Visible = True
            CMDInfo.Visible = True
            cmdLanzar.Visible = True
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            cmdMoverHechi(0).Enabled = True
            cmdMoverHechi(1).Enabled = True
            
            Image1(0).Visible = False
            Image1(1).Visible = False
            Image1(2).Visible = False
            imggrupo.Visible = False
        Case 2
            InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centromenu.jpg")
            
            picInv.Visible = False
            
            lblInvInfo.Visible = False
        
            hlst.Visible = False
            CMDInfo.Visible = False
            cmdLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            cmdMoverHechi(0).Enabled = False
            cmdMoverHechi(1).Enabled = False
            
            Image1(0).Visible = True
            Image1(1).Visible = True
            Image1(2).Visible = True
            imggrupo.Visible = True
    End Select
End Sub

Private Sub imggrupo_Click()
Call Audio.PlayWave(SND_CLICK)
frmgrupo.Show , frmMain
End Sub

Private Sub imgHora_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgHora.ToolTipText = "La hora en el mundo es: " & Get_Time_String
End Sub

Private Sub imgquest_Click()
Call WriteQuestInformacion
End Sub



Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Application.IsAppActive() Then  'Implemento lo propuesto por GD, se puede usar macro aun que se esté en otra ventana
    '    Call DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    'Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
    
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
  '  Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

 Private Sub nomodocombate_Click()
    Call WriteCombatModeToggle
    IScombate = Not IScombate
    frmMain.Modocombate.Visible = Not frmMain.Modocombate.Visible
    frmMain.Nomodocombate.Visible = Not frmMain.Nomodocombate.Visible
End Sub

Private Sub modocombate_Click()
    Call WriteCombatModeToggle
    IScombate = Not IScombate
    frmMain.Modocombate.Visible = Not frmMain.Modocombate.Visible
    frmMain.Nomodocombate.Visible = Not frmMain.Nomodocombate.Visible
End Sub

Private Sub PicSeg_Click()
    Call WriteSafeToggle
End Sub

Private Sub Coord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     MouseX = x
     MouseY = y
     
    InfoMapAct = True
     
    Call InfoMapa
End Sub

Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_DblClick()
Call Form_DblClick
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
     If Button = 2 Then
      ActivarMacroTrabajo
    End If
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub second_Timer()
With luz_dia(Hour(time))
    Call engine.setup_ambient
    base_light = engine.change_day_effect(day_r_old, day_g_old, day_b_old, .r, .g, .b)
    End With
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.text = ""
        KeyCode = 0
        SendTxt.Visible = False
        End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    Call WritePickUp
End Sub

Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub Form_Click()
    
    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not frmForo.Visible And Not frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_Load()
        ModoHabla = 1
    mNormal.Checked = True
    Me.Picture = LoadPicture(App.path & "\Recursos\Interface\Todo.jpg")
    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\Centroinventario.jpg")
    Call Make_Transparent_Richtext(RecTxt.hWnd)
    Me.Height = 9000
    Me.Width = 12000

    Me.Left = 0
    Me.Top = 0
    UltPos = -1
UltimoIndex = -1
UltPosInterface = -1
UltPosSolapas = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x - Renderer.Left
    MouseY = y - Renderer.Top
    
    lblexpactivo = False
    
 InMouseExp = False
    If UserPasarNivel = 0 Then
        lblexp.Caption = "¡Nivel máximo!"
    Else
        frmMain.lblexp.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    End If
    
    Call RestaurarCentroActual
    InfoMapAct = False
    Call InfoMapa
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmestadisticas.Iniciar_Labels
            frmestadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        
        Case 2
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub cmdDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub

Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not modApplication.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
      (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
      (Not frmMSG.Visible) And (Not frmForo.Visible) And _
      (Not frmestadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) Then
        picInv.SetFocus
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.text)
            CharAscii = Asc(mid$(SendTxt.text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.text = tempstr
        End If
        
        stxtbuffer = SendTxt.text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If Not frmPasswd.Visible And Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name And Forms(i).name <> frmPasswd.name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    Nomodocombate.Visible = True
    Modocombate.Visible = False
    SkillPoints = 0
    Alocados = 0
    
frmMain.Nomodocombate.Visible = True
frmMain.Modocombate.Visible = False
    
    RemoveAllDialogs
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Winsock1.RemoteHostIP)
#End If
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
        
        If EstadoLogin = CrearNuevoPj Then

    frmPasswd.lblstatus.Caption = "Conectado. Enviando datos..."
End If

#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login


        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login

        Case E_MODO.Dados
            frmCrearPersonaje.Show
            
#If SeguridadAlkon Then
            Call ProtectForm(frmCrearPersonaje)
#End If
    End Select
    frmMain.Nomodocombate.Visible = True
frmMain.Modocombate.Visible = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
#If SeguridadAlkon Then
    Call DataReceived(Data)
#End If
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < Renderer.Left Or clicX > Renderer.Left + (32 * 17) Then Exit Function
    If clicY < Renderer.Top Or clicY > Renderer.Top + (32 * 13) Then Exit Function
    
    InGameArea = True
End Function


Private Sub cmdMinimizar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgminicerra.Picture = LoadPicture(App.path & "\Recursos\Interface\minimizardown.jpg")
End Sub

Private Sub cmdCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
End
End Sub

Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgminicerra.Picture = LoadPicture(App.path & "\Recursos\Interface\minimizarover.jpg")
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgminicerra.Picture = LoadPicture(App.path & "\Recursos\Interface\cerrarover.jpg")
End Sub

Private Sub cmdMinimizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.WindowState = vbMinimized
imgminicerra.Picture = Nothing
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgminicerra.Picture = LoadPicture(App.path & "\Recursos\Interface\cerrardown.jpg")
End Sub

Private Sub SolapasRestaurar(Index As Integer)

imgCentros(Index).Picture = Nothing
imgminicerra.Picture = Nothing
'cmdMensaje.Picture = Nothing

End Sub


Private Sub RestaurarCentroActual()

Select Case CentroActual
    Case CentroHechizos
        If UltPosInterface <> -1 Then Call CentroHechizosRestaurar(UltPosInterface)
    Case CentroInventario
    Case CentroMenu
        If UltPosInterface <> -1 Then Call CentroMenuRestaurar(UltPosInterface)
End Select

If UltPosSolapas <> -1 Then Call SolapasRestaurar(UltPosSolapas)

UltPosInterface = -1
UltPosSolapas = -1

imgminicerra.Picture = Nothing
'cmdMensaje.Picture = Nothing
lblInvInfo.Caption = ""

End Sub

Private Sub MostrarCentroHechizos()
    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centrohechizos.jpg")
    cmdLanzar.Visible = True
    CMDInfo.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    hlst.Visible = True
End Sub

Private Sub OcultarCentroHechizos()
    hlst.Visible = False
    cmdLanzar.Visible = False
    CMDInfo.Visible = False
    cmdMoverHechi(1).Visible = False
    cmdMoverHechi(0).Visible = False
End Sub

Private Sub MostrarCentroMenu()
    Image1(0).Visible = True
    Image1(1).Visible = True
    Image1(2).Visible = True
    Image1(3).Visible = True
    Image1(4).Visible = True
    Image1(5).Visible = True
    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centromenu.jpg")
End Sub

Private Sub OcultarCentroMenu()
    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    Image1(3).Visible = False
    Image1(4).Visible = False
    Image1(5).Visible = False
End Sub

Private Sub CambiaCentro(NuevoCentro As Byte)

CentroActual = NuevoCentro

If NuevoCentro = CentroMenu Then
    Call MostrarCentroMenu
    Call OcultarCentroHechizos
    Call OcultarCentroInventario
ElseIf NuevoCentro = CentroHechizos Then
    Call MostrarCentroHechizos
    Call OcultarCentroMenu
    Call OcultarCentroInventario
Else
    Call MostrarCentroInventario
    Call OcultarCentroHechizos
    Call OcultarCentroMenu
End If

End Sub


Private Sub MostrarCentroInventario()
    InvEqu.Picture = LoadPicture(App.path & "\Recursos\Interface\centroinventario.bmp")
    picInv.Visible = True
    lblInvInfo.Visible = True
    lblInvInfo = ""
End Sub

Private Sub OcultarCentroInventario()
    picInv.Visible = False
    lblInvInfo.Visible = False
End Sub

Private Sub CentroHechizosRestaurar(Index As Integer)

cmdLanzar.Picture = Nothing
CMDInfo.Picture = Nothing
cmdMoverHechi(1) = Nothing
cmdMoverHechi(0) = Nothing
End Sub

Private Sub CentroMenuRestaurar(Index As Integer)

Image1(Index).Picture = Nothing

End Sub

Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

BotonElegido = Index + 1

If MacroKeys(BotonElegido).TipoAccion = 0 Or Button = vbRightButton Then
    frmbindkey.Show vbModeless, frmMain
Else
    Call frmbindkey.Bind_Accion(Index + 1)
End If

End Sub

Private Sub picMacro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltimoIndex <> Index Then
    'If UltimoIndex >= 0 Then DibujarMenuMacros UltimoIndex + 1
    'DibujarMenuMacros Index + 1, 1
    UltimoIndex = Index
End If

End Sub

Private Sub cmdMensaje_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdmensaje.Picture = LoadPicture(App.path & "\Recursos\Interface\modotextodown.jpg")
End Sub

Private Sub cmdMensaje_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdmensaje.Picture = LoadPicture(App.path & "\Recursos\Interface\modotextoover.jpg")
End Sub

Private Sub cmdMensaje_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call cmdMensaje_MouseMove(Button, Shift, x, y)
PopUpMenu mhabla
cmdmensaje.Picture = LoadPicture(App.path & "\Recursos\Interface\modotextoover.jpg")
End Sub
Private Sub mNormal_Click()
    ModoHabla = 1
    mNormal.Checked = True
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = False
    Mgrupo.Checked = False
    mGlobal.Checked = False
    Mgms.Checked = False
End Sub
 
Private Sub mGritar_Click()
    ModoHabla = 2
    mNormal.Checked = False
    mGritar.Checked = True
    mPrivado.Checked = False
    mClan.Checked = False
    Mgrupo.Checked = False
    mGlobal.Checked = False
    Mgms.Checked = False
End Sub
 
Private Sub mPrivado_Click()
    ModoHabla = 3
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = True
    PrivateTo = InputBox("Escriba el nombre: ", "Mensajeria Privada", "")
    mClan.Checked = False
    Mgrupo.Checked = False
    mGlobal.Checked = False
    Mgms.Checked = False
End Sub
 
Private Sub mClan_Click()
    ModoHabla = 4
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = True
    Mgrupo.Checked = False
    mGlobal.Checked = False
    Mgms.Checked = False
End Sub

Private Sub mGrupo_Click()
    ModoHabla = 5
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = False
    Mgrupo.Checked = True
    mGlobal.Checked = False
    Mgms.Checked = False
End Sub

Private Sub mGlobal_Click()
    ModoHabla = 6
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = False
    Mgrupo.Checked = False
    mGlobal.Checked = True
    Mgms.Checked = False
End Sub
Private Sub mGms_Click()
    ModoHabla = 7
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = False
    Mgrupo.Checked = False
    mGlobal.Checked = False
    Mgms.Checked = True
End Sub
