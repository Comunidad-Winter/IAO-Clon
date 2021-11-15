Attribute VB_Name = "LuzItem"
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
Public Const NumItemsLuz As Byte = 2 'Numeros de items con luz
Public GrhItemLuz(1 To NumItemsLuz) As Integer
Public GrhLuze(1 To 12) As Integer
 
 Function TieneLuz(ByVal X As Byte, ByVal Y As Byte)
'*******************************************
'*        Carga de luces en items          *
'*******************************************
Dim i As Byte
For i = 1 To NumItemsLuz
  If GrhItemLuz(1) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 250, 149, 48
  End If
Next i
  If GrhLuze(1) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 0, 0
  End If
  If GrhLuze(2) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 255, 0
  End If
  If GrhLuze(3) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 128, 0
  End If
  If GrhLuze(4) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 0, 0
  End If
  If GrhLuze(5) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 0, 255, 255
  End If
  If GrhLuze(6) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 0, 255, 0
  End If
  If GrhLuze(7) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 128, 0
  End If
  If GrhLuze(8) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 255, 0
  End If
  If GrhLuze(9) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 255, 255
  End If
  If GrhLuze(10) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 255, 0
  End If
  If GrhLuze(11) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 0, 255, 0
  End If
  If GrhLuze(12) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
  MapData(X, Y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map X, Y, 1, 255, 255, 255
  End If
End Function
 
Function DeletLuz(ByVal X As Byte, ByVal Y As Byte)
'*******************************************
'*        Aqui se Borran Las Luces         *
'*******************************************

Dim i As Byte
For i = 1 To NumItemsLuz
    If GrhItemLuz(i) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
Next i
    If GrhLuze(1) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(2) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(3) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(4) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(5) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(6) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(7) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(8) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(9) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(10) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(11) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
    If GrhLuze(12) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
End Function

 
Sub ObjLuz()

'*******************************************
'*     Carga de luz en item desde Grh      *
'*******************************************

GrhItemLuz(1) = 1521 'Fogata
GrhItemLuz(2) = 510 'Daga comun
GrhLuze(1) = 716 'Espada Mata Dragones
GrhLuze(2) = 873 'Baculo Dm +10
GrhLuze(3) = 892 'Harbinger Kin
GrhLuze(4) = 900 'Daga Infernal
GrhLuze(5) = 997 'Espada AO
GrhLuze(6) = 19590 'Baculo Lazurt
GrhLuze(7) = 19592 'Baculo +20
GrhLuze(8) = 19595 'Espada Saramiana
GrhLuze(9) = 1003 'Escudo De Reflex +30
GrhLuze(10) = 25679 'Nudillos de oro
GrhLuze(11) = 18378  'Collar de Rykan
GrhLuze(12) = 17504  'Pendiente de experto

End Sub

