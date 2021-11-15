VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLight"
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

Option Base 0
Private Type tLight
    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_x As Byte
    map_y As Byte
    range As Byte
End Type
 
Private light_list() As tLight
Private NumLights As Byte
Public Function Create_Light_To_Map(ByVal map_x As Byte, ByVal map_y As Byte, Optional range As Byte = 3, Optional ByVal red As Byte = 255, Optional ByVal green = 255, Optional ByVal blue As Byte = 255)

    NumLights = NumLights + 1
   
    ReDim Preserve light_list(1 To NumLights) As tLight
   
    light_list(NumLights).RGBcolor.r = red
    light_list(NumLights).RGBcolor.g = green
    light_list(NumLights).RGBcolor.b = blue
   
    light_list(NumLights).RGBcolor.a = 255
   
    light_list(NumLights).range = range
   
    light_list(NumLights).active = True
   
    light_list(NumLights).map_x = map_x
    light_list(NumLights).map_y = map_y
   
    Call LightRender(NumLights)
    
End Function
Public Function Delete_Light_To_Map(ByVal X As Byte, ByVal Y As Byte)
   
    Dim i As Long
   
    For i = 1 To NumLights
        If light_list(i).map_x = X And light_list(i).map_y = Y Then
            Delete_Light_To_Index i
        End If
    Next i
 
End Function
Public Function Delete_Light_To_Index(ByVal light_index As Byte)
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
 
    light_list(light_index).active = False
 
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = 0
                MapData(Xa, Ya).light_value(1) = 0
                MapData(Xa, Ya).light_value(2) = 0
                MapData(Xa, Ya).light_value(3) = 0
            End If
        Next Xa
    Next Ya
   
End Function
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
   
    Dim CurrentColor As D3DCOLORVALUE
   
    pRadio = cRadio * 32
   
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
   
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
   
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.r), Round(CurrentColor.g), Round(CurrentColor.b))
    Else
        LightCalculate = TileLight
    End If
End Function
Private Sub LightRender(ByVal light_index As Integer)
 
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.r = 128
    AmbientColor.g = 128
    AmbientColor.b = 128
   
    LightColor = light_list(light_index).RGBcolor
       
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
       
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
 
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
   
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).range, light_list(light_index).map_x * 32, light_list(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
               
            End If
        Next Xa
    Next Ya
End Sub
