VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox headview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1695
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   4545
      Width           =   375
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":004C
      Left            =   840
      List            =   "frmCrearPersonaje.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":006F
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0071
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2490
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1050
      Width           =   5865
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0073
      Left            =   8550
      List            =   "frmCrearPersonaje.frx":0075
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3585
      Width           =   2745
   End
   Begin VB.Image menoshead 
      Height          =   600
      Left            =   1320
      Top             =   4440
      Width           =   390
   End
   Begin VB.Image mashead 
      Height          =   600
      Left            =   2160
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label modconstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   44
      Top             =   7140
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modcarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Left            =   2160
      TabIndex        =   43
      Top             =   6780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Left            =   2160
      TabIndex        =   42
      Top             =   6420
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Left            =   2160
      TabIndex        =   41
      Top             =   6060
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Left            =   2160
      TabIndex        =   39
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2835
   End
   Begin VB.Label Skill 
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
      Index           =   1
      Left            =   5280
      TabIndex        =   38
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   0
      Left            =   5280
      TabIndex        =   37
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   2
      Left            =   5280
      TabIndex        =   36
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   3
      Left            =   5280
      TabIndex        =   35
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   4
      Left            =   5280
      TabIndex        =   34
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   5
      Left            =   5280
      TabIndex        =   33
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   6
      Left            =   5280
      TabIndex        =   32
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   7
      Left            =   5280
      TabIndex        =   31
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   8
      Left            =   5280
      TabIndex        =   30
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   9
      Left            =   5280
      TabIndex        =   29
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   10
      Left            =   5280
      TabIndex        =   28
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   11
      Left            =   5280
      TabIndex        =   27
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   12
      Left            =   5280
      TabIndex        =   26
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   13
      Left            =   5280
      TabIndex        =   25
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   14
      Left            =   7335
      TabIndex        =   24
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   15
      Left            =   7335
      TabIndex        =   23
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   16
      Left            =   7335
      TabIndex        =   22
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   17
      Left            =   7335
      TabIndex        =   21
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0077
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":01C9
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":031B
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":046D
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":05BF
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0711
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0863
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":09B5
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0B07
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0C59
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0DAB
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":0EFD
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":104F
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":11A1
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":12F3
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1445
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1597
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":16E9
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":183B
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":198D
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1ADF
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1C31
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1D83
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":1ED5
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2027
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2179
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":22CB
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":241D
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":256F
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":26C1
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2813
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2965
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2AB7
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2C09
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2D5B
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2EAD
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":2FFF
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":3151
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":32A3
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":33F5
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":3547
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":3699
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   20
      Left            =   7335
      TabIndex        =   20
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   19
      Left            =   7335
      TabIndex        =   19
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   18
      Left            =   7335
      TabIndex        =   18
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   21
      Left            =   7335
      TabIndex        =   17
      Top             =   4950
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":37EB
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":393D
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   22
      Left            =   7335
      TabIndex        =   16
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":3A8F
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":3BE1
      Top             =   5430
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   23
      Left            =   7335
      TabIndex        =   15
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":3D33
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":3E85
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   24
      Left            =   7335
      TabIndex        =   14
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":3FD7
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":4129
      Top             =   6180
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   25
      Left            =   7335
      TabIndex        =   13
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":427B
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":43CD
      Top             =   6540
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   26
      Left            =   7335
      TabIndex        =   12
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":451F
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":4671
      Top             =   6930
      Width           =   195
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   525
      Left            =   2610
      TabIndex        =   11
      Top             =   8220
      Width           =   6795
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   6795
      TabIndex        =   6
      Top             =   7260
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   165
      Index           =   2
      Left            =   1200
      MouseIcon       =   "frmCrearPersonaje.frx":47C3
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   1260
   End
   Begin VB.Image boton 
      Height          =   615
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmCrearPersonaje.frx":4915
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1605
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   9600
      MouseIcon       =   "frmCrearPersonaje.frx":4A67
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1680
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2445
      TabIndex        =   4
      Top             =   6780
      Width           =   225
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2445
      TabIndex        =   3
      Top             =   6420
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2445
      TabIndex        =   2
      Top             =   7140
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2445
      TabIndex        =   1
      Top             =   6060
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2445
      TabIndex        =   0
      Top             =   5700
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Option Explicit
Public Actual As Integer
Public SkillPoints As Byte
Private MaxEleccion As Integer, MinEleccion As Integer
Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub Boton_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index
        Case 0
            
            Dim i As Integer
            Dim k As Object
            i = 1
            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next
            
            UserName = txtNombre.Text
            
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
            End If
            
            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex + 1
            UserClase = lstProfesion.ListIndex + 1
            
            UserAtributos(1) = Val(lbFuerza.Caption)
            UserAtributos(2) = Val(lbInteligencia.Caption)
            UserAtributos(3) = Val(lbAgilidad.Caption)
            UserAtributos(4) = Val(lbCarisma.Caption)
            UserAtributos(5) = Val(lbConstitucion.Caption)
            
            UserHogar = lstHogar.ListIndex + 1
            
            'Barrin 3/10/03
            If CheckData() Then
                frmPasswd.Show vbModal, Me
            End If
            
        Case 1
            Call Audio.Music_Load(2)
            
            frmConnect.Picture = LoadPicture(App.path & "\Recursos\Interface\conectar.jpg")
            Unload Me
            
            
        Case 2
            Call Audio.PlayWave(SND_DICE)
            Call TirarDados
    End Select
End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.path & "\Recursos\Interface\cp-interface.jpg")

SkillPoints = 10
Puntos.Caption = SkillPoints

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstHogar.Clear

For i = LBound(Ciudades()) To UBound(Ciudades())
    lstHogar.AddItem Ciudades(i)
Next i


lstRaza.Clear

For i = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(i)
Next i


lstProfesion.Clear

For i = LBound(ListaClases()) To UBound(ListaClases())
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 1
UserHead = 0
Call TirarDados

End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    Image1.Picture = LoadPicture(App.path & "\Recursos\Interface\" & lstProfesion.Text & ".jpg")
End Sub
Private Sub lstGenero_Click()
    Call DameOpciones
End Sub
Private Sub lstRaza_Click()
Call DameOpciones
        modfuerza.Visible = True
        modconstitucion.Visible = True
        modAgilidad.Visible = True
        modInteligencia.Visible = True
        modcarisma.Visible = True
    Select Case (lstRaza.List(lstRaza.ListIndex))
   Case Is = "Humano"
       modfuerza.Caption = "+1"
       modconstitucion.Caption = "+2"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = ""
        modcarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
       modconstitucion.Caption = "+1"
        modAgilidad.Caption = "+3"
       modInteligencia.Caption = "+1"
       modcarisma.Caption = "+2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+1"
      modconstitucion.Caption = ""
        modAgilidad.Caption = "+1"
       modInteligencia.Caption = "+2"
        modcarisma.Caption = "-3"
    Case Is = "Enano"
        modfuerza.Caption = "+3"
        modconstitucion.Caption = "+3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "-6"
        modcarisma.Caption = "-3"
    Case Is = "Gnomo"
        modfuerza.Caption = "-5"
        modAgilidad.Caption = "+4"
        modInteligencia.Caption = "+3"
        modcarisma.Caption = "+1"
    Case Is = "Orco"
        modfuerza.Caption = "+ 5"
        modconstitucion.Caption = "+6"
        modAgilidad.Caption = "- 2"
       modInteligencia.Caption = "-6"
       modcarisma.Caption = "-2"
End Select
End Sub
Private Sub MenosHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual - 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
headview.Cls
Call engine.DrawGrhtoHdc(headview.hdc, HeadData(Actual).Head(3).grhindex, 8, 5)
headview.Refresh
End Sub
Private Sub MasHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual + 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
headview.Cls
Call engine.DrawGrhtoHdc(headview.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
headview.Refresh
End Sub

Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Masculino"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).mStart
                MaxEleccion = Head_Range(HUMANO).mEnd
                MinEleccion = Head_Range(HUMANO).mStart
            Case "Elfo"
                Actual = Head_Range(ELFO).mStart
                MaxEleccion = Head_Range(ELFO).mEnd
                MinEleccion = Head_Range(ELFO).mStart
            Case "Elfo Oscuro"
                Actual = Head_Range(ElfoOscuro).mStart
                MaxEleccion = Head_Range(ElfoOscuro).mEnd
                MinEleccion = Head_Range(ElfoOscuro).mStart
            Case "Enano"
                Actual = Head_Range(Enano).mStart
                MaxEleccion = Head_Range(Enano).mEnd
                MinEleccion = Head_Range(Enano).mStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).mStart
                MaxEleccion = Head_Range(Gnomo).mEnd
                MinEleccion = Head_Range(Gnomo).mStart
            Case "Orco"
                Actual = Head_Range(Orco).mStart
                MaxEleccion = Head_Range(Orco).mEnd
                MinEleccion = Head_Range(Orco).mStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Femenino"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).fStart
                MaxEleccion = Head_Range(HUMANO).fEnd
                MinEleccion = Head_Range(HUMANO).fStart
            Case "Elfo"
                Actual = Head_Range(ELFO).fStart
                MaxEleccion = Head_Range(ELFO).fEnd
                MinEleccion = Head_Range(ELFO).fStart
            Case "Elfo Oscuro"
                Actual = Head_Range(ElfoOscuro).fStart
                MaxEleccion = Head_Range(ElfoOscuro).fEnd
                MinEleccion = Head_Range(ElfoOscuro).fStart
            Case "Enano"
                Actual = Head_Range(Enano).fStart
                MaxEleccion = Head_Range(Enano).fEnd
                MinEleccion = Head_Range(Enano).fStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).fStart
                MaxEleccion = Head_Range(Gnomo).fEnd
                MinEleccion = Head_Range(Gnomo).fStart
            Case "Orco"
                Actual = Head_Range(Orco).fStart
                MaxEleccion = Head_Range(Orco).fEnd
                MinEleccion = Head_Range(Orco).fStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
End Select
 
headview.Cls
Call engine.DrawGrhtoHdc(headview.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
headview.Refresh
End Sub
