VERSION 5.00
Begin VB.Form frmRetirar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   Picture         =   "frmRetirar.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRetirar.frx":5B69
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
      Left            =   720
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2040
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmRetirar"
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

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
Call WriteRetirar
End Sub
