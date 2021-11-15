VERSION 5.00
Begin VB.Form frmCarp 
   Caption         =   "Carpintero"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
      Height          =   315
      Left            =   3000
      MouseIcon       =   "frmCarp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
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
      Height          =   315
      Left            =   120
      MouseIcon       =   "frmCarp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2425
      Width           =   855
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************* ****************
'ImperiumAO - v1.0
'************************************************* ****************
'Copyright (C) 2015 Gaston Jorge Martinez
'Copyright (C) 2015 Alexis Rodriguez
'Copyright (C) 2015 Luis Merino
'Copyright (C) 2015 Girardi Luciano Valentin
'
'Respective portions copyright by taxpayers below.
'
'This library is free software; you can redistribute it and / or
'Modify it under the terms of the GNU General Public
'License as published by the Free Software Foundation version 2.1
'The License
'
'This library is distributed in the hope that it will be useful,
'But WITHOUT ANY WARRANTY; without even the implied warranty
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'************************************************* ****************
'
'************************************************* ****************
'You can contact me at:
'Gaston Jorge Martinez (Zenitram@Hotmail.com)
'************************************************* ****************

Option Explicit

Private Sub Command3_Click()
    On Error Resume Next
 
    If Int(Val(Cantidad)) < 1 Or Int(Val(Cantidad)) > 10000 Then
    MsgBox "La cantidad es invalida.", vbCritical
    Exit Sub
End If
    Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), (Cantidad))
    If frmMain.macrotrabajo.Enabled Then _
        MacroBltIndex = ObjCarpintero(lstArmas.ListIndex + 1) & Cantidad.Text
   
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

