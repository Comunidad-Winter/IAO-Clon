VERSION 5.00
Begin VB.Form frmGameGuard 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   Icon            =   "frmGameGuard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPres 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmGameGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************* ****************
'Eternal Online - v1.0
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

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\Resources\Interface\gameguard.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FinGG = True
End Sub

Private Sub tmrPres_Timer()
'Static ticks As Long

'ticks = ticks + 1

'If ticks = 1 Then
    
'ElseIf ticks = 2 Then
    'Me.Picture = LoadPicture(App.Path & "\Resources\graphics\datafull.bmp")
'ElseIf ticks = 2 Then
'    Me.Picture = LoadPicture(App.Path & "\Resources\graphics\argentum.bmp")
'Else
FinGG = True
'End If

End Sub

