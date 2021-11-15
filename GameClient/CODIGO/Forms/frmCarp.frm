VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carpintero"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1170
      TabIndex        =   3
      Text            =   "1"
      Top             =   2160
      Width           =   3165
   End
   Begin VB.ListBox lstArmas 
      Height          =   1815
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      MouseIcon       =   "frmCarp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
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
      Height          =   435
      Left            =   255
      MouseIcon       =   "frmCarp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2670
      Width           =   1710
   End
   Begin VB.Label cantidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cantidad"
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
      Left            =   300
      TabIndex        =   4
      Top             =   2190
      Width           =   855
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmHerrero - ImperiumAO -  AOshao v1.3.0
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Pablo Ignacio Márquez (morgolock@speedy.com.ar)
'   - First Relase
'   - 2016
'*****************************************************************
Option Explicit

Private Sub Command3_Click()
    On Error Resume Next

    Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), Val(txtCantidad.text))
    frmMain.ActivarMacroTrabajo
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call Make_Transparent_Form(Me.hWnd, 210)
End Sub

Private Sub txtCantidad_Change()
If Val(txtCantidad.text) < 0 Then
    txtCantidad.text = 1
End If

If Val(txtCantidad.text) > 1000 Then
    txtCantidad.text = 1
End If

End Sub

