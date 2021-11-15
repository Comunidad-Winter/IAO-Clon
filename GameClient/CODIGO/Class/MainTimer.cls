VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


Private Const CANTIDADTIMERS As Byte = 8

Private Type Timer
    Interval As Long
    CurrentTick As Long
    StartTick As Long
    Run As Boolean
End Type

Dim Timer(1 To CANTIDADTIMERS) As Timer

Public Enum TimersIndex
    Attack = 1
    Work = 2
    UseItemWithU = 3
    UseItemWithDblClick = 4
    SendRPU = 5
    CastSpell = 6
    Arrows = 7
    CastAttack = 8
End Enum

Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)

If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
Timer(TimerIndex).Interval = Interval

End Sub
Public Sub Start(ByVal TimerIndex As TimersIndex)

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True
    
End Sub
Public Sub Continue(ByVal TimerIndex As TimersIndex)

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True
    
    Timer(TimerIndex).StartTick = Timer(TimerIndex).StartTick + Timer(TimerIndex).CurrentTick
    
End Sub
Public Sub Pause(ByVal TimerIndex As TimersIndex)

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = False
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick
    
End Sub
Public Function Check(ByVal TimerIndex As TimersIndex, Optional Restart As Boolean = True) As Boolean

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    If Not Timer(TimerIndex).Run Then Exit Function
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick
    
    If Timer(TimerIndex).CurrentTick >= Timer(TimerIndex).Interval Then
        Check = True
        If Restart Then
            Timer(TimerIndex).StartTick = GetTickCount
            If (TimerIndex = TimersIndex.Attack) Or (TimerIndex = TimersIndex.CastSpell) Then
                Timer(TimersIndex.CastAttack).StartTick = GetTickCount
            ElseIf TimerIndex = TimersIndex.CastAttack Then
                Timer(TimersIndex.Attack).StartTick = GetTickCount
                Timer(TimersIndex.CastSpell).StartTick = GetTickCount
            End If
        End If
    End If
    
End Function
Public Sub Restart(ByVal TimerIndex As TimersIndex)

    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).StartTick = GetTickCount
    
End Sub
