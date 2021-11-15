VERSION 5.00
Begin VB.Form Frmquest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de quest"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Abandonar 
      Caption         =   "Abandonar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton BotonSalir 
      Caption         =   "Rechazar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton BotonAceptar 
      Caption         =   "Aceptar propuesta"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Lbquestinfoitem 
      BackStyle       =   0  'Transparent
      Caption         =   "QuestInfo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lbdesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la quest"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lbQuestInfoCantidad 
      BackStyle       =   0  'Transparent
      Caption         =   "QuestInfo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3840
   End
End
Attribute VB_Name = "Frmquest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Abandonar_Click()
Call WriteQuestAbandonar
Unload Me
End Sub

Private Sub BotonAceptar_Click()
Call WriteQuestAceptar
Unload Me
End Sub

Private Sub BotonSalir_Click()
Unload Me
End Sub

