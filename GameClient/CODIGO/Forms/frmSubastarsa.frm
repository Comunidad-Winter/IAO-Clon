VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centro de Subastas"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton Command5 
         Caption         =   "Nueva Subasta"
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
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Buscar"
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
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   445
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar Solamente Mis Subastas."
         Height          =   255
         Left            =   5640
         TabIndex        =   5
         Top             =   5880
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
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
         Left            =   7320
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Comprar (Directo)"
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
         Left            =   120
         TabIndex        =   2
         Top             =   5880
         Width           =   2535
      End
      Begin VB.ListBox lstObjetos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   4905
         ItemData        =   "frmSubastarsa.frx":0000
         Left            =   120
         List            =   "frmSubastarsa.frx":0007
         TabIndex        =   1
         Top             =   840
         Width           =   8415
      End
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'JAO
Call WriteCompraSubasta(0, lstObjetos.ListIndex + 1)
End Sub

Private Sub Command2_Click()
Unload frmSubastar
End Sub

Private Sub Command5_Click()
frmSubastas.Show
End Sub

