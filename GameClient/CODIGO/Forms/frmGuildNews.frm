VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GuildNews"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Clanes aliados"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox aliados 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0000
         Left            =   120
         List            =   "frmGuildNews.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clanes con los que estamos en guerra"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
      Begin VB.ListBox guerra 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0004
         Left            =   120
         List            =   "frmGuildNews.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GuildNews"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox news 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmGuildNews.frx":0008
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5880
      Width           =   4335
   End
End
Attribute VB_Name = "frmGuildNews"
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
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub

