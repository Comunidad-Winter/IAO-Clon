VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herrero"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "1"
      Top             =   2530
      Width           =   615
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
      MouseIcon       =   "frmHerrero.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
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
      Left            =   2880
      MouseIcon       =   "frmHerrero.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   1350
   End
   Begin VB.ListBox lstArmas 
      Height          =   2010
      Left            =   150
      TabIndex        =   2
      Top             =   450
      Width           =   4080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Armaduras"
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
      Left            =   2280
      MouseIcon       =   "frmHerrero.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Armas"
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
      Left            =   130
      MouseIcon       =   "frmHerrero.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   2040
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2010
      Left            =   135
      TabIndex        =   5
      Top             =   465
      Width           =   4080
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2560
      Width           =   855
   End
End
Attribute VB_Name = "frmHerrero"
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

Private Sub Command1_Click()
lstArmaduras.Visible = False
lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), (Cantidad))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmasHerrero(lstArmas.ListIndex + 1) & Cantidad.Text
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), (Cantidad))
        
        If frmMain.macrotrabajo.Enabled Then _
            MacroBltIndex = ArmadurasHerrero(lstArmaduras.ListIndex + 1) & Cantidad.Text
    End If

    Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

