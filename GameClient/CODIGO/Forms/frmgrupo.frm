VERSION 5.00
Begin VB.Form frmgrupo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grupo"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton crear 
      Caption         =   "crear"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton salir 
      Caption         =   "salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton abandonar 
      Caption         =   "abandonar"
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
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton invitar 
      Caption         =   "invitar"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton expulsar 
      Caption         =   "Expulsar"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox user 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox text 
      Height          =   1425
      ItemData        =   "frmgrupo.frx":0000
      Left            =   120
      List            =   "frmgrupo.frx":0002
      TabIndex        =   1
      Top             =   1200
      Width           =   4710
   End
   Begin VB.Label lid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lider :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label LG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lider"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmgrupo.frx":0004
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4635
   End
End
Attribute VB_Name = "frmgrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abandonar_Click()
 Call WritePartyLeave
   Unload Me
End Sub

Private Sub crear_Click()
   Call WritePartyCreate
   Call WriteRequestPartyForm
   Unload Me
End Sub

Private Sub expulsar_Click()
  Call WritePartyKick(text.text)
   Call WriteRequestPartyForm
   Unload Me
End Sub

Private Sub Form_Load()
text.Clear
Call WriteRequestPartyForm
text.Clear
Call Make_Transparent_Form(Me.hWnd, 210)
End Sub

Private Sub User_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
End If
End Sub

Private Sub User_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub invitar_Click()
 If Len(user) > 0 Then
   
        If Not IsNumeric(user) Then
        
            Call WritePartyAcceptMember(Trim(user.text))
            
            Unload Me
            
            Call WriteRequestPartyForm
            
        End If
        
    End If
   
   'Call WritePartyAcceptMember(User.Text)
   'Call WriteRequestPartyForm
   Unload Me
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub usar_Change()

End Sub
