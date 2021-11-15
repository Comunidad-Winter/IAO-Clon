VERSION 5.00
Begin VB.Form FrmPasajes 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pirata"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Barca 
      BackColor       =   &H80000007&
      Height          =   550
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton salir 
      BackColor       =   &H8000000A&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox Pasaje 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Viajar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¿ A dònde quieres Viajar?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Costo 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
End
Attribute VB_Name = "FrmPasajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If LenB(Pasaje.Text) <> 0 Then
 Call WriteViajer(Pasaje.ListIndex)
End If
Unload Me
End Sub

Private Sub Form_Load()
'Lugares a viajar
Pasaje.AddItem "Nix"
Pasaje.AddItem "Banderbill"
Pasaje.AddItem "Lindos"
Pasaje.AddItem "Ullathorpe"
End Sub


Private Sub Pasaje_Click()
'El oro
Select Case Pasaje
 Case "Nix"
  Costo.Caption = "Costo: 1700"
 Case "Banderbill"
  Costo.Caption = "Costo: 1850"
 Case "Lindos"
  Costo.Caption = "Costo: 1900"
 Case "Ullathorpe"
  Costo.Caption = "Costo: 1200"
End Select
End Sub


Private Sub salir_Click()
Unload Me
End Sub
