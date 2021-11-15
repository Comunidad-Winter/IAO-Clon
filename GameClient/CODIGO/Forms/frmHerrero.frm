VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Herrero"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtherrero 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "1"
      Top             =   3060
      Width           =   3735
   End
   Begin VB.CommandButton boton 
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
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "salir"
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
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "Ar&maduras"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton boton 
      Caption         =   "&Armas"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2205
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   3075
      Width           =   855
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmHerrero - ImperiumAO - Aoshao v1.3.0
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
'*****************************************************************


Option Explicit

Private Sub Command1_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub Boton_Click(Index As Integer)
 Select Case Index
 
   Case 0
    lstArmaduras.Visible = False
    lstArmas.Visible = True
    
   Case 1
    lstArmaduras.Visible = True
    lstArmas.Visible = False

   Case 2
    Unload Me
    
   Case 3
    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), Val(txtherrero.text))
   frmMain.ActivarMacroTrabajo
    End If
    
    If lstArmaduras.Visible Then
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), Val(txtherrero.text))
    frmMain.ActivarMacroTrabajo
    End If
    
    Unload Me

 End Select
End Sub

Private Sub Form_Load()
Call Make_Transparent_Form(Me.hWnd, 210)
End Sub

