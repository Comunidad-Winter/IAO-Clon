VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   4875
   ClientTop       =   15
   ClientWidth     =   12000
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
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":F172
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   12480
      Top             =   1680
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   12480
      Top             =   3120
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   7
         Left            =   0
         TabIndex        =   31
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   12360
      Top             =   2160
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   810
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   2
         Left            =   0
         TabIndex        =   26
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   1980
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   3
         Left            =   0
         TabIndex        =   27
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   4
         Left            =   0
         TabIndex        =   28
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3150
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   5
         Left            =   0
         TabIndex        =   29
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   6
         Left            =   0
         TabIndex        =   30
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   8
         Left            =   0
         TabIndex        =   32
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   9
         Left            =   0
         TabIndex        =   33
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   11
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   10
         Left            =   0
         TabIndex        =   34
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   8430
      Width           =   480
      Begin VB.Label lblMacro 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   1
         Left            =   0
         TabIndex        =   24
         Top             =   350
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   10170
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   7335
      Width           =   1500
      Begin VB.Shape clan 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   240
         Shape           =   1  'Square
         Top             =   120
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Shape Fadd 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   960
         Shape           =   1  'Square
         Top             =   360
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   600
         Shape           =   1  'Square
         Top             =   600
         Width           =   45
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
      TabIndex        =   7
      Top             =   2220
      Width           =   2415
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8850
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Renderer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   210
      ScaleHeight     =   419
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   546
      TabIndex        =   5
      Top             =   2040
      Width           =   8190
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
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1740
      Visible         =   0   'False
      Width           =   7455
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1485
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":2B1E4
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
      Top             =   5850
      Width           =   1350
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
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   6255
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
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15480
      TabIndex        =   39
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15240
      TabIndex        =   38
      Top             =   9960
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   16680
      TabIndex        =   37
      Top             =   10080
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15000
      TabIndex        =   36
      Top             =   8640
      Width           =   855
   End
   Begin VB.Label LBLITEMS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Left            =   9000
      TabIndex        =   35
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Image Modhab 
      Height          =   375
      Left            =   7800
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   11520
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9360
      TabIndex        =   22
      Top             =   8580
      Width           =   345
   End
   Begin VB.Label lblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9360
      TabIndex        =   21
      Top             =   8355
      Width           =   345
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9255
      Picture         =   "frmMain.frx":2B261
      ToolTipText     =   "Seguro"
      Top             =   7605
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":2B69F
      ToolTipText     =   "Modo Combate"
      Top             =   7605
      Width           =   300
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":2BADD
      ToolTipText     =   "Modo Combate"
      Top             =   7605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgHora 
      Height          =   555
      Left            =   6645
      Top             =   8400
      Width           =   1740
   End
   Begin VB.Image PicSeg 
      Height          =   255
      Left            =   9255
      Picture         =   "frmMain.frx":2BF1B
      ToolTipText     =   "Seguro"
      Top             =   7605
      Width           =   300
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   2
      Left            =   10680
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   2
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   4920
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
      TabIndex        =   9
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NicknamePJ"
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
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   240
      Width           =   2505
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   1
      Left            =   9720
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   0
      Left            =   8520
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Image CMDInfo 
      Height          =   495
      Left            =   10560
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   1065
   End
   Begin VB.Image cmdLanzar 
      Height          =   510
      Left            =   8760
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   540
      Index           =   0
      Left            =   11400
      Top             =   2880
      Width           =   420
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   540
      Index           =   1
      Left            =   11400
      Top             =   3360
      Width           =   420
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8565
      Picture         =   "frmMain.frx":2C359
      Top             =   1245
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
      TabIndex        =   2
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10200
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   300
   End
   Begin VB.Label exp 
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
      Height          =   150
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
      Left            =   11130
      TabIndex        =   0
      Top             =   840
      Width           =   135
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
   Begin VB.Menu mHabla 
      Caption         =   "ModosHabla"
      Visible         =   0   'False
      Begin VB.Menu mNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mGM 
         Caption         =   "MsgGM"
      End
      Begin VB.Menu mClan 
         Caption         =   "Clan"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Public InMouseExp As Boolean

Public Es_Real As Boolean

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

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

Private Sub exp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
InMouseExp = True
exp.Caption = UserExp & "/" & UserPasarNivel
If UserPasarNivel = 0 Then
    exp.Caption = "¡Nivel máximo!"
    ExpShp.Width = 121
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Es_Real = False Then
    Exit Sub
    End If
    Es_Real = True
    
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    'Audio.MusicActivated = Not Audio.MusicActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                     If IScombate = True Then
                        modocombate.Visible = True
                        nomodocombate.Visible = False
                    Else
                        modocombate.Visible = False
                        nomodocombate.Visible = True
                    End If
                
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
                    Call WriteSafeToggle
                    
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

        Case vbKeyF1 To vbKeyF11
            If Not frmOpciones.Macros.value = 0 Then
            Call UsarMacro(KeyCode - vbKeyF1 + 1)
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeMostrarFps)
        If MostrarFPS = False Then
          MostrarFPS = True
        Else
          MostrarFPS = False
        End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            Call WriteMeditate

        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
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
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not GetAsyncKeyState(KeyCode) < 0 Then
Es_Real = False
Exit Sub
End If
Es_Real = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        Call EndGame
        Cancel = 1
    End If
End Sub


Private Sub Image2_Click()
End
End Sub

Private Sub imgCentros_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        
        Case 0
            InvEqu.Picture = LoadPicture(App.path & "\Resources\Interface\Inventory.jpg")

            picInv.Visible = True
            
            LBLITEMS.Visible = True
            LBLITEMS = ""
        
            hlst.Visible = False
            cmdINFO.Visible = False
            cmdLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            cmdMoverHechi(0).Enabled = False
            cmdMoverHechi(1).Enabled = False
        
            Image1(0).Visible = False
            Image1(1).Visible = False
            Image1(2).Visible = False
            LBLITEMS.Visible = True
            
        Case 1
            InvEqu.Picture = LoadPicture(App.path & "\Resources\Interface\spells.jpg")
            '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
            
            picInv.Visible = False
            LBLITEMS.Visible = False
            
            hlst.Visible = True
            cmdINFO.Visible = True
            cmdLanzar.Visible = True
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            cmdMoverHechi(0).Enabled = True
            cmdMoverHechi(1).Enabled = True
            
            Image1(0).Visible = False
            Image1(1).Visible = False
            Image1(2).Visible = False
            LBLITEMS.Visible = False
            
        Case 2
            InvEqu.Picture = LoadPicture(App.path & "\Resources\Interface\menu.jpg")
            
            picInv.Visible = False
            
            LBLITEMS.Visible = False
        
            hlst.Visible = False
            cmdINFO.Visible = False
            cmdLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            cmdMoverHechi(0).Enabled = False
            cmdMoverHechi(1).Enabled = False
            
            Image1(0).Visible = True
            Image1(1).Visible = True
            Image1(2).Visible = True
            
            LBLITEMS.Visible = False
            
    End Select
End Sub

Private Sub imgHora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHora.ToolTipText = "La hora en el mundo es: " & Get_Time_String
End Sub



Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otPociones Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otruna Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otArmadura Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otcasco Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otescudo Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
    If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otMonturas Then
                Call DesactivarMacroTrabajo
                Exit Sub
        End If
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0

End Sub

Private Sub Minimap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call WriteWarpChar("YO", UserMap, IIf(X < 1, 1, X), Y)
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
                    If IScombate = True Then
                        modocombate.Visible = True
                        nomodocombate.Visible = False
                    Else
                        modocombate.Visible = False
                        nomodocombate.Visible = True
                    End If
End Sub
Private Sub ModoCombate_Click()
Call WriteCombatModeToggle
                    IScombate = Not IScombate
                    If IScombate = True Then
                        modocombate.Visible = True
                        nomodocombate.Visible = False
                    Else
                        modocombate.Visible = False
                        nomodocombate.Visible = True
                    End If
End Sub

Private Sub PicSeg_Click()
    Call WriteSafeToggle
End Sub

Private Sub Coord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MouseX = X
     MouseY = Y
     
    InfoMapAct = True
     
    Call InfoMapa
End Sub

Private Sub Picture1_Click()
End Sub

Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_DblClick()
Call Form_DblClick
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    If Button = 2 Then
      ActivarMacroTrabajo
    End If
End Sub


Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Second_Timer()

'**********************************
' Autor: Gaston Martinez
'**********************************

With luz_dia(Hour(time))
    Call Engine.setup_ambient
    base_light = Engine.change_day_effect(day_r_old, day_g_old, day_b_old, .r, .g, .b)
End With
    
'Programs detecting unsuitable to the game from the Second Timer
If Detected("sXe Injected.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: sXe Injected, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("Cheat Engine.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: Cheat Engine, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("svchost.exe.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: svchost.exe.exe, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("processhacker.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: Process Hacker, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("Sandboxie.exe") Then
Call MsgBox("ATENCIÒN: Se Detecto el Programa: SandBoxie.exe, Puede Causar el mal funcionamiento del Juego, Por favor Cierrelo", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

End If

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
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
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
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
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0
                                'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                '    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                               ' End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0
                                'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                'End With
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
    
    Me.Picture = LoadPicture(App.path & "\Resources\Interface\Main.jpg")
    InvEqu.Picture = LoadPicture(App.path & "\Resources\Interface\Inventory.jpg")
    Call Make_Transparent_Richtext(RecTxt.hWnd)
    
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - Renderer.Left
    MouseY = Y - Renderer.Top
    
    InMouseExp = False
    If UserPasarNivel = 0 Then
        exp.Caption = "¡Nivel máximo!"
        ExpShp.Width = 121
    Else
        frmMain.exp.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    End If
    
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
            LlegaronEstadisticas = False
            Call WriteRequestEstadisticas
            Call FlushBuffer
            
            Do While Not LlegaronEstadisticas
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronEstadisticas = False
        
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
    
    Call UsarItem
    Call EquiparItem

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
      (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) Then
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
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
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
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
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

''''''''''''''''''''''''''''''''''''''
'     Socket1                        '
''''''''''''''''''''''''''''''''''''''
Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Second.Enabled = True

    Call Login
    
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    Dim mifrm As Form
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    On Local Error Resume Next
    For Each mifrm In Forms
        If Not mifrm.name = Me.name And Not mifrm.name = frmCrearPersonaje.name And Not mifrm.name = frmConnect.name Then
            Unload mifrm
        End If
    Next
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
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
    


    SkillPoints = 0
    Alocados = 0
    
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

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        frmMensaje.Show
        frmMensaje.msg.Caption = "Si el Servidor no le Conecta en minutos, tiene Problemas Con la internet, Por Favor Verifiquelo. Para mas informacion www.EternalOnline.com.ar"
        Exit Sub
    ElseIf ErrorCode = 24061 Then
        frmMensaje.Show
        frmMensaje.msg.Caption = "No Hay Conecxion Con el Servidor. Puede ser que el servidor este Office o que tenga problemas con su conecxion a internet. Para mas informacion www.EternalOnline.com.ar"
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect

End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim Data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Debug.Print Asc(RD)
    Data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    
    'Put data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Macros_Click(Index As Integer)
 
    If frmMacro.Visible = True Then Exit Sub
 
    'Umm.. parece q no selecciono una accion para el macro.
    If MacroList(Index).mTipe = 0 Then
        MacroIndex = Index
        frmMacro.MacroLbl = "Tecla F" & Index
        frmMacro.Show , Me
        Exit Sub
    End If
   
    'acccion!
    Call UsarMacro(CByte(Index))
End Sub
Private Sub Macros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMacro.Visible = True Then Exit Sub
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    If Button = vbKeyRButton Then
        MacroIndex = Index
        frmMacro.MacroLbl = "Tecla F" & Index
        frmMacro.Show , Me
    End If
End Sub

Private Sub Modhab_Click()
PopUpMenu mHabla
End Sub

Private Sub mNormal_Click()
    ModoHabla = 1
    mNormal.Checked = True
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = False
End Sub
 
Private Sub mGritar_Click()
    ModoHabla = 2
    mNormal.Checked = False
    mGritar.Checked = True
    mPrivado.Checked = False
    mClan.Checked = False
End Sub
 
Private Sub mPrivado_Click()
    ModoHabla = 3
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = True
    PrivateTo = InputBox("Escriba el nombre: ", "Mensajeria Privada", "")
    mClan.Checked = False
End Sub
 
Private Sub mClan_Click()
    ModoHabla = 4
    mNormal.Checked = False
    mGritar.Checked = False
    mPrivado.Checked = False
    mClan.Checked = True
End Sub
