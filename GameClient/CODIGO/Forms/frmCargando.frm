VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Eternal Online"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCargando.frx":1171AA
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image picLoad 
      Height          =   630
      Left            =   2220
      Picture         =   "frmCargando.frx":18F713
      Top             =   8040
      Width           =   7560
   End
End
Attribute VB_Name = "frmCargando"
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
Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 595
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
 
Public Sub progresoConDelay(ByVal porcentaje As Integer)
 
If porcentaje = porcentajeActual Then Exit Sub
 
Dim step As Integer, stepInterval As Integer, Timer As Long, tickCount As Long
 
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
 
Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (Timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    Timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call establecerProgreso(porcentajeActual)
Loop
 
End Sub
 
 
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    picLoad.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    picLoad.Width = DEFAULT_PROGRESS_WIDTH
Else
    picLoad.Width = 0
End If
porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    compararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    compararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
 
End Function

