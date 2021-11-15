VERSION 5.00
Begin VB.Form frmDELETEPJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrar Personaje"
   ClientHeight    =   1665
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1250
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Estas seguro que desea borrar este personaje ? En ese caso escriba: DELETE"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmDELETEPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Y si no quiero escribir DELETE ?
If Not Text1.Text = "DELETE" Then
MsgBox "Por favor Escirba DELETE", vbInformation
Exit Sub
End If



MsgBox "Falta sistemaaaaaaa Mierda", vbCritical
Unload Me

End Sub
