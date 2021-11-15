VERSION 5.00
Begin VB.Form frmSubastas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subastas"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPrecio 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "100.000"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estime un precio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Objeto seleccionado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSubastas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not IsNumeric(txtPrecio) Then
Call MsgBox("El precio no es numérico.", vbInformation)
Else
Call WriteSubastarObjeto(txtPrecio)
Unload frmSubastas
End If
End Sub

Private Sub Form_Load()
Dim ItemSlot As Byte
ItemSlot = Inventario.SelectedItem
lblItem.Caption = Inventario.ItemName(ItemSlot)
End Sub

