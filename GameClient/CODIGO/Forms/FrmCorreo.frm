VERSION 5.00
Begin VB.Form FrmCorreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correo"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Enviar Mensaje"
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   6495
      Begin VB.PictureBox ITEM 
         BackColor       =   &H00000000&
         Height          =   550
         Left            =   4920
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   480
         Width           =   550
      End
      Begin VB.ListBox lstItems 
         Height          =   2595
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox CheckObjeto 
         Caption         =   "Adjuntar Item"
         Height          =   195
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDestinatario 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtMensaje 
         Height          =   1995
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Enviar"
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Costo:"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Mensaje:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Para:"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensajes"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar Item"
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   5040
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.PictureBox OBJREGALADO 
         BackColor       =   &H80000007&
         Height          =   550
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   1800
         Width           =   550
      End
      Begin VB.TextBox LblCorreo 
         Height          =   1455
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.ListBox LstCorreos 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   2000
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckObjeto_Click()
If CheckObjeto.value = 1 Then
Label6.Caption = "Costo: 1300"
Else
Label6.Caption = "Costo:"
End If
End Sub

Private Sub Form_Load()


    '   Cargamos la lista
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ItemName(i) = "" Then
            lstItems.AddItem "Nada"
        Else
            lstItems.AddItem Inventario.ItemName(i)
        End If
    Next i
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
Call WritePacketsCorreo(1, LstCorreos.ListIndex + 1)
End Sub

Private Sub Command6_Click()
If CheckObjeto.value <> 0 Then 'Con objeto
    Call WriteEnviarCorreo(txtDestinatario.Text, txtMensaje.Text, Inventario.OBJIndex(lstItems.ListIndex + 1), txtCantidad.Text)
Else
    Call WriteEnviarCorreo(txtDestinatario.Text, txtMensaje.Text, 0, 0)
End If

End Sub

Private Sub Command7_Click()
Call WritePacketsCorreo(1, LstCorreos.ListIndex + 1)
txtDestinatario.Text = ""
txtMensaje.Text = ""
CheckObjeto.value = 0
Label6.Caption = "Costo:"
End Sub

Private Sub lstItems_Click()
ITEM.Cls
 Call Engine.DrawGrhToHdc(ITEM.hdc, Inventario.grhindex(lstItems.ListIndex + 1), 0, 0)
End Sub

