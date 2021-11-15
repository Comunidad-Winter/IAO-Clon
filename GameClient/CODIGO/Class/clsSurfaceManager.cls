VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
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
 
Const HASH_TABLE_SIZE As Long = 337
Private Const BYTES_PER_MB As Long = 1048576
Private Const MIN_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB
 
Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type
 
Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type
 
Private TexList(HASH_TABLE_SIZE - 1) As HashNode
 
Private mD3D As D3DX8
Private device As Direct3DDevice8
 
Private mCantidadGraficos As Integer
Private maxBytesToUse As Long
Private mFreeMemoryBytes As Long
 
Private Declare Function GetTickCount Lib "kernel32" () As Long
 
Private Sub Class_Terminate()

    Dim i As Long
    Dim j As Long
   
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
           
            Erase .SurfaceEntry
        End With
    Next i
End Sub
 
Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
If FileName = 0 Then
Debug.Print "0 GRH ATMPT TO BE LOADED"
Exit Function
End If
 
    Dim i As Long
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With
 
    Set GetTexture = CrearGrafico(FileName, textwidth, textheight)
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean

    mCantidadGraficos = 0
   

    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
   
    Init = True
End Function
 
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8

On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim Index As Integer
    Index = Archivo Mod HASH_TABLE_SIZE
    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            .FileName = Archivo
           
            .UltimoAcceso = GetTickCount
   
            Set .Texture = mD3D.CreateTextureFromFileEx(device, App.path & "\Resources\graphics\" & LTrim(str(Archivo)) & ".bmp", _
                D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
               
            .Texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .size = surface_desc.size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .Texture
            mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
        End With
    End With
    Do While mFreeMemoryBytes < 0
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".bmp"

End Function
 
Private Function RemoveLRU() As Boolean

    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
   
    LRUtime = GetTickCount
   
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
   
    Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)
   
    Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
   
    If LRUj Then
        RemoveLRU = True
       
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
           
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
   
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function

