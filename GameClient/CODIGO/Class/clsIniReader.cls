VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
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
Option Base 0

Private Type ChildNode
    Key As String
    value As String
End Type

Private Type MainNode
    name As String
    values() As ChildNode
    numValues As Integer
End Type

Private fileData() As MainNode

Private MainNodes As Long
Private Sub Class_Terminate()

    Dim i As Long
    
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        
        Erase fileData
    End If
    
End Sub
Public Sub Initialize(ByVal file As String)

    Dim handle As Integer
    Dim Text As String
    Dim Pos As Long
    
    handle = FreeFile
    
    Open file For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, Text
        
        If Len(Text) Then
            If Left$(Text, 1) = "[" Then
                Pos = InStr(2, Text, "]")
                If Pos Then
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    
                    MainNodes = MainNodes + 1
                End If
            Else
                Pos = InStr(2, Text, "=")
                If Pos Then
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            ReDim Preserve .values(.numValues) As ChildNode
                            
                            .values(.numValues).value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).Key = UCase$(Left$(Text, Pos - 1))
                            
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close handle
    
    Dim i As Long
    
    If MainNodes Then
        Call SortMainNodes(0, MainNodes - 1)
        
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then _
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
End Sub
Private Sub SortChildNodes(ByRef Node As MainNode, ByVal first As Integer, ByVal last As Integer)

    Dim min As Integer
    Dim max As Integer
    Dim comp As String
    Dim temp As ChildNode
    
    min = first
    max = last
    
    With Node
        comp = .values((min + max) \ 2).Key
    
        Do While min <= max
            Do While .values(min).Key < comp And min < last
                min = min + 1
            Loop
            Do While .values(max).Key > comp And max > first
                max = max - 1
            Loop
            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    
    If first < max Then SortChildNodes Node, first, max
    If min < last Then SortChildNodes Node, min, last
    
End Sub
Private Sub SortMainNodes(ByVal first As Integer, ByVal last As Integer)

    Dim min As Integer
    Dim max As Integer
    Dim comp As String
    Dim temp As MainNode
    
    min = first
    max = last
    
    comp = fileData((min + max) \ 2).name
    
    Do While min <= max
        Do While fileData(min).name < comp And min < last
            min = min + 1
        Loop
        Do While fileData(max).name > comp And max > first
            max = max - 1
        Loop
        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If first < max Then SortMainNodes first, max
    If min < last Then SortMainNodes min, last
    
End Sub
Public Function GetValue(ByVal Main As String, ByVal Key As String) As String

    Dim i As Long
    Dim j As Long
    
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        j = FindKey(fileData(i), UCase$(Key))
        
        If j >= 0 Then GetValue = fileData(i).values(j).value
    End If
    
End Function
Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long

    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = Node.numValues - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If Node.values(mid).Key < Key Then
            min = mid + 1
        ElseIf Node.values(mid).Key > Key Then
            max = mid - 1
        Else
            FindKey = mid
            Exit Function
        End If
    Loop

    FindKey = Not mid
End Function
Private Function FindMain(ByVal name As String) As Long

    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If fileData(mid).name < name Then
            min = mid + 1
        ElseIf fileData(mid).name > name Then
            max = mid - 1
        Else
            FindMain = mid
            Exit Function
        End If
    Loop

    FindMain = Not mid
End Function




