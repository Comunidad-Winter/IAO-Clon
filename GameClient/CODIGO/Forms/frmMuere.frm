VERSION 5.00
Begin VB.Form frmMuere 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   2685
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   3915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMuere.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMuere.frx":8D25A
   ScaleHeight     =   2685
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1920
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Has Muerto ! Si queres volver a la ciudad  click en Aceptar y si no en Cancelar"
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
End
Attribute VB_Name = "frmMuere"
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

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
Call WriteRegresar
Unload frmMuere
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image1.Tag = "0" Then
            Image1.Tag = "1"
        End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Image1.Tag = "1"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Tag = "1" Then
    Image1.Picture = Nothing
    Image1.Tag = "0"
End If
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

