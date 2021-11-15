VERSION 5.00
Begin VB.Form frmHogar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   Picture         =   "frmHogar.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro que quieres cambiar de ciudad ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2040
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmHogar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Call WriteHogar(2)
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
