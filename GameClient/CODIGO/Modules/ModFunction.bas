Attribute VB_Name = "ModFunction"
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

Public CurServerIp As String
Public CurServerPort As Integer
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Public Function Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Public Function LoadInterface(ByVal FileName As String) As IPictureDisp
    Extract_File Interface, App.path & "\resources\Interfaces\", LTrim(FileName) & ".jpg", App.path & "\resources\Interface\"
       Set LoadInterface = LoadPicture(App.path & "\resources\Interface\" & FileName & ".jpg")
End Function
