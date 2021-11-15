Attribute VB_Name = "LuzItem"

Option Explicit
Public Const NumItemsLuz As Byte = 2 'Numeros de items con luz
Public GrhItemLuz(1 To NumItemsLuz) As Integer
Public GrhLuze(1 To 12) As Integer
 
 Function TieneLuz(ByVal x As Byte, ByVal y As Byte)
'*******************************************
'*        Carga de luces en items          *
'*******************************************
Dim i As Byte
For i = 1 To NumItemsLuz
  If GrhItemLuz(1) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 250, 149, 48
  End If
Next i
  If GrhLuze(1) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 0, 0
  End If
  If GrhLuze(2) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 255, 0
  End If
  If GrhLuze(3) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 128, 0
  End If
  If GrhLuze(4) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 0, 0
  End If
  If GrhLuze(5) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 0, 255, 255
  End If
  If GrhLuze(6) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 0, 255, 0
  End If
  If GrhLuze(7) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 128, 0
  End If
  If GrhLuze(8) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 255, 0
  End If
  If GrhLuze(9) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 255, 255
  End If
  If GrhLuze(10) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 255, 0
  End If
  If GrhLuze(11) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 0, 255, 0
  End If
  If GrhLuze(12) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 0 Then
  MapData(x, y).OBJInfo.TieneLuz = 1
  Light.Create_Light_To_Map x, y, 1, 255, 255, 255
  End If
End Function
 
Function DeletLuz(ByVal x As Byte, ByVal y As Byte)
'*******************************************
'*        Aqui se Borran Las Luces         *
'*******************************************

Dim i As Byte
For i = 1 To NumItemsLuz
    If GrhItemLuz(i) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
Next i
    If GrhLuze(1) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(2) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(3) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(4) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(5) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(6) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(7) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(8) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(9) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(10) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(11) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
    End If
    If GrhLuze(12) = MapData(x, y).ObjGrh.grhindex And MapData(x, y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map x, y
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


