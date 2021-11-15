VERSION 5.00
Begin VB.Form FrmNuevoCorreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo correo"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtMensaje 
      Height          =   1725
      Left            =   120
      TabIndex        =   4
      Text            =   "Mensaje"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtDestinatario 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Destinatario"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Cantidad"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ListBox lstItems 
      Height          =   1815
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CheckBox CheckObjeto 
      Caption         =   "Objeto"
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmNuevoCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If CheckObjeto.value <> 0 Then 'Con objeto
    Call WriteEnviarCorreo(txtDestinatario.Text, txtMensaje.Text, Inventario.OBJIndex(lstItems.ListIndex + 1), txtCantidad.Text)
Else
    Call WriteEnviarCorreo(txtDestinatario.Text, txtMensaje.Text, 0, 0)
End If

End Sub

Private Sub Form_Load()


    '   Cargamos la lista
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ItemName(i) = "" Then
            lstItems.AddItem "Nada"
        Else
            lstItems.AddItem Inventario.ItemName(i) & "(" & Inventario.Amount(i) & ")"
        End If
    Next i
End Sub
