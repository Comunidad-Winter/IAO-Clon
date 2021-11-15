VERSION 5.00
Begin VB.Form FrmAlquimia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alquimia"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command 
      Caption         =   "Crear"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "1"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label cantidad 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2450
      Width           =   855
   End
End
Attribute VB_Name = "FrmAlquimia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click(Index As Integer)
On Error Resume Next

  Select Case Index
   Case 0
    Unload Me
    
   Case 1
    Call WriteCraftAlquimia(objalquimia(lst.ListIndex + 1), Val(text.text))
    frmMain.ActivarMacroTrabajo
    Unload Me
  End Select
  
End Sub

Private Sub Form_Load()
Call Make_Transparent_Form(Me.hWnd, 210)
End Sub

Private Sub Text_Change()

If Val(text.text) <= 0 Then
    text.text = 1
End If

If Val(text.text) > 1000 Then
    text.text = 1
End If

End Sub

