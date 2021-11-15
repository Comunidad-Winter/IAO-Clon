VERSION 5.00
Begin VB.Form frmRecuperar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Olvidaste Tu Contraseña ?"
   ClientHeight    =   3420
   ClientLeft      =   9900
   ClientTop       =   5235
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Contraseña Nueva:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Mail de la cuenta:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre de la cuenta :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Recuperar la Contraseña de su cuenta debe ingresar los siguentes datos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
