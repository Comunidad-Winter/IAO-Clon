VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPres 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmPres - ImperiumAO - v1.3.0
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

Private Sub Form_Load()
'Me.Icon = frmIniciando.Icon
Me.Picture = LoadPicture(App.path & "\Recursos\Interface\cargandomundo.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FinPres = True
End Sub

Private Sub tmrPres_Timer()
'Static ticks As Long

'ticks = ticks + 1

'If ticks = 1 Then
    
'ElseIf ticks = 2 Then
    'Me.Picture = LoadPicture(App.Path & "\Recursos\Graficos\datafull.bmp")
'ElseIf ticks = 2 Then
'    Me.Picture = LoadPicture(App.Path & "\Recursos\Graficos\argentum.bmp")
'Else
FinPres = True
'End If

End Sub

