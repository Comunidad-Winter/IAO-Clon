VERSION 5.00
Begin VB.Form frmSastreria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sastreria"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4935
   LinkTopic       =   "Sastre"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "salir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tejer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2760
      Width           =   1710
   End
   Begin VB.ListBox lstsastrero 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4665
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   2340
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   4
      Top             =   2450
      Width           =   735
   End
End
Attribute VB_Name = "frmSastreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cantidad_Change()
If Val(cantidad.text) < 0 Then
    cantidad.text = 1
End If

If Val(cantidad.text) > 1000 Then
    cantidad.text = 1
End If
End Sub

Private Sub Command1_Click()
Call WriteCraftSastreria(ObjSastreria(lstsastrero.ListIndex + 1), Val(txtCantidad.text))
frmMain.ActivarMacroTrabajo
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Make_Transparent_Form(Me.hWnd, 210)
End Sub


