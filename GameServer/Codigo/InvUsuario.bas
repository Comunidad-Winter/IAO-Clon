Attribute VB_Name = "InvUsuario"
'AoshaoServer 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'AoshaoServer is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo ErrHandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If Cantidad > 10000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.map
            For j = UserList(UserIndex).Pos.X - 10 To UserList(UserIndex).Pos.X + 10
                For k = UserList(UserIndex).Pos.Y - 10 To UserList(UserIndex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ORO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Call LogDesarrollo(UserList(UserIndex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If EsGM(UserIndex) Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            Dim AuxPos As WorldPos
            
            If UserList(UserIndex).clase = eClass.drakkar And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

ErrHandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .Amount <= Cantidad Then
            If .Equipped = 1 Then
                Call Desequipar(UserIndex, Slot)
                 ElseIf UserList(UserIndex).flags.Montando = 1 Then
                If Slot = UserList(UserIndex).Invent.MonturaSlot Then
                    DoEquita UserIndex, .ObjIndex, Slot
                End If
            End If
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With
End Sub


Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim obj As obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
        obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If num + MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.Amount
        End If
        
        obj.Amount = num
        
        Call MakeObj(obj, map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If ObjData(obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call WriteConsoleMsg(UserIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
        End If
        
        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
        
        'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(obj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " tiró al piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
        ElseIf obj.Amount = 1000 Then 'Es mucha cantidad?
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(obj.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " tiró del piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
            End If
        End If
  Else
    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount - num

If MapData(map, X, Y).ObjInfo.Amount <= 0 Then
    MapData(map, X, Y).ObjInfo.ObjIndex = 0
    MapData(map, X, Y).ObjInfo.Amount = 0
    
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByRef obj As obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
        MapData(map, X, Y).ObjInfo.Amount = MapData(map, X, Y).ObjInfo.Amount + obj.Amount
    Else
        MapData(map, X, Y).ObjInfo = obj
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(obj.ObjIndex).GrhIndex, X, Y))
    End If
End If
 '//Agregamos las pos de los objetos
            If ObjData(obj.ObjIndex).OBJType <> otFogata And _
                Not EsObjetoFijo(ObjData(obj.ObjIndex).OBJType) Then
                Dim xPos As WorldPos
                xPos.map = map
                xPos.X = X
                xPos.Y = Y
                AgregarObjetoLimpieza xPos
                End If
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo ErrHandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
ErrHandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim obj As ObjData
Dim MiObj As obj

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex
        'If DropMode = 1 Then 'y que es lo que no te anda que lo manda siempre a la bille pero esta funcion es de agarrar lo tira y va a la bille pero coso lo tendria que tirar y va directo esta mal este call
        If obj.OBJType = 5 Then ''el 5 es el oro? si, queres que si esta en 1 haga esto, sino que tendria que hacer? osea esta en 1 lo tira si esta en 0 va a la bille de una y aparece en consola has ganado,tambien si lo agarras del suelo aparece has ganado X de oro y queres que valla a la bille y no al inventario eso ya esta chekeado pero el tema es que lo manda a la bille de una sin tirarlo estando en 1 que lo tendria que tir
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
        Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.Amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has ganado " & MiObj.Amount & " monedas de oro.", FontTypeNames.FONTTYPE_ORO)
        Exit Sub 'fañtan las comillas adentro del & lo estas poiniendo mal
       ' End If
End If
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.Amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            
            'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            ElseIf MiObj.Amount = 1000 Then 'Es mucha cantidad?
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                    Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
                End If
            End If
            
        End If
        
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
     
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
     
    Case eOBJType.otNudillos
        UserList(UserIndex).Invent.NudiEqpIndex = 0
        UserList(UserIndex).Invent.NudiEqpSlot = 0
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otAnillo
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
    
    Case eOBJType.otArmadura ' Puede ser un escudo, casco , o vestimenta
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        Select Case ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).Subtipo
        
            Case 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(UserIndex)
                If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
            Case 1
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

            Case 2
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        
        End Select
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo ErrHandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(UserIndex) Then
        FaccionPuedeUsarItem = esArmada(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(UserIndex) Then
        FaccionPuedeUsarItem = esCaos(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Public Function NivelPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    If UserList(UserIndex).Stats.ELV >= ObjData(ObjIndex).MinELV Then
        NivelPuedeUsarItem = True
    Else
        NivelPuedeUsarItem = False
    End If
End Function

Public Function SkillsPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
If ObjData(ObjIndex).OBJType = otBarcos Then
    If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) >= ObjData(ObjIndex).MinSkill Then
        SkillsPuedeUsarItem = True
    Else
        SkillsPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).OBJType = otMonturas Then
    If UserList(UserIndex).Stats.UserSkills(eSkill.Domar) >= ObjData(ObjIndex).MinSkill Then
        SkillsPuedeUsarItem = True
    Else
        SkillsPuedeUsarItem = False
    End If
Else
    SkillsPuedeUsarItem = True
End If
End Function
Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo ErrHandler

'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
obj = ObjData(ObjIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If

If Not obj.MinELV = 0 Then
    If obj.MinELV > UserList(UserIndex).Stats.ELV Then
        Call WriteConsoleMsg(UserIndex, "Debes ser nivel " & obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If
Select Case obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                   If UserList(UserIndex).flags.Montando = True Then
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        End If
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
                Else
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.WeaponAnim = obj.WeaponAnim
            Else
                UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
      
    Case eOBJType.otNudillos
        If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                UserList(UserIndex).Char.WeaponAnim = NingunArma
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Exit Sub
            End If
            
            'Quitamos el arma si tiene alguna equipada
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudiEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.NudiEqpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.NudiEqpSlot = Slot

            UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otAnillo
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
  Case eOBJType.otArmadura
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
      
               'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Call DarCuerpoDesnudo(UserIndex)
                    If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                
           'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.body = obj.Ropaje
            Else
                UserList(UserIndex).Char.body = obj.Ropaje
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            UserList(UserIndex).flags.Desnudo = 0
            

       If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                UserList(UserIndex).flags.Desnudo = 0
            Else
                Call WriteConsoleMsg(UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
    Case eOBJType.otCASCO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.CascoEqpSlot = Slot
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.CascoAnim = obj.CascoAnim
            Else
                UserList(UserIndex).Char.CascoAnim = obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(UserIndex, Slot)
                 If UserList(UserIndex).flags.Mimetizado = 1 Then
                     UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(UserIndex).Invent.Object(Slot).Equipped = 1
             UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
             UserList(UserIndex).Invent.EscudoEqpSlot = Slot
             
             If UserList(UserIndex).flags.Mimetizado = 1 Then
                 UserList(UserIndex).CharMimetizado.ShieldAnim = obj.ShieldAnim
             Else
                 UserList(UserIndex).Char.ShieldAnim = obj.ShieldAnim
                 
                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             End If
         Else
             Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
ErrHandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub
Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo ErrHandler

If ObjData(ItemIndex).RazaTipo > 0 Then
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).raza = eRaza.Humano Or _
       UserList(UserIndex).raza = eRaza.Elfo Or _
       UserList(UserIndex).raza = eRaza.Orco Or _
       UserList(UserIndex).raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 1)
    ElseIf UserList(UserIndex).raza = eRaza.Orco Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 3)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 2)
    End If
Else
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).raza = eRaza.Humano Or _
       UserList(UserIndex).raza = eRaza.Elfo Or _
       UserList(UserIndex).raza = eRaza.Drow Or _
       UserList(UserIndex).raza = eRaza.Orco Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If
End If

Exit Function
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function
Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase drakkar y Pescador.
'*************************************************

Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj

If UserList(UserIndex).Invent.Object(Slot).Amount = 0 Then Exit Sub

obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If obj.OBJType = eOBJType.otWeapon Then
    If obj.proyectil = 1 Then
        If Not UserList(UserIndex).flags.ModoCombate Then
            Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + obj.MinHam
        If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then _
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)
         
     Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(UserIndex).Stats.ELV < 25 Then
            If UserList(UserIndex).clase <> eClass.fisher And UserList(UserIndex).clase <> eClass.Sastre Then
                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If UserList(UserIndex).Stats.ELV < 20 Then
                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        
        If UserList(UserIndex).flags.Montando = 0 Then
            If ((LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True, False)) _
                    And UserList(UserIndex).flags.Navegando = 0) _
                    Or UserList(UserIndex).flags.Navegando = 1 Then
                Call DoNavega(UserIndex, obj, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No disponible cabalgando.", FontTypeNames.FONTTYPE_INFO)
        End If
        
          Case eOBJType.otMonturas
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Los muertos no dominan los animales. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                If Not ((LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True, False)) _
                    And UserList(UserIndex).flags.Montando = 0) _
                     Or UserList(UserIndex).flags.Montando = 1 Then
                        Call DoEquita(UserIndex, ObjIndex, Slot)
        Else
                        Call WriteConsoleMsg(UserIndex, "No puedes montar en el mar.", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
    End If
    
    
    Case eOBJType.otGuita
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(Slot).Amount
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call WriteConsoleMsg(UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
       
        If ObjData(ObjIndex).proyectil = 1 Then
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not UserList(UserIndex).flags.ModoCombate Then
                Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
         Else
            If UserList(UserIndex).flags.TargetObj = Leña Then
                If UserList(UserIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, _
                         UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                End If
            End If
        End If
        
       
        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjIndex
            Case CAÑA_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
            Case HACHA_LEÑADOR
                Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
            Case TIJERA_ALQUIMIA
                   Call WriteWorkRequestTarget(UserIndex, eSkill.Botanica)
            Case PIQUETE_MINERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
             Case OLLA_ALQUIMIA
                Call EnivarObjAlquimiaConstruibles(UserIndex)
                Call WriteShowAlquimiaForm(UserIndex)
            Case COSTURERO_SASTRE
                Call EnivarObjSastreriaConstruibles(UserIndex)
                Call WriteShowSastreriaForm(UserIndex)
        End Select
        
        
         Case eOBJType.OtMapa
          Call WriteShowMapaForm(UserIndex)
          
          
      Case eOBJType.otBolsaOro
      If UserList(UserIndex).flags.Muerto = 1 Then
      Call WriteConsoleMsg(UserIndex, " ¡¡Estas Muerto!!", FontTypeNames.FONTTYPE_INFO)
      Exit Sub
      End If
      
      If obj.CuantoAgrega Then
      UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + obj.CuantoAgrega
      Call WriteConsoleMsg(UserIndex, "Has ganado " & obj.CuantoAgrega & " Monedas de oro", FontTypeNames.FONTTYPE_ORO)
      Call QuitarUserInvItem(UserIndex, Slot, 1)
      Call WriteUpdateGold(UserIndex)
     Call UpdateUserInv(False, UserIndex, Slot)
      End If
     
      
   Case eOBJType.otPociones
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteAtacar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = obj.TipoPocion
                
        Select Case UserList(UserIndex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteAgilidad(UserIndex)
            Case 2 'Modif la fuerza
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Call WriteFuerza(UserIndex)
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then _
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)
                If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado = 1 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Case 6  ' Pocion Negra
                If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
       End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
        Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = obj.clave Then
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
            Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.MaxMAN > 0 Then
            If UserList(UserIndex).flags.Hambre = 0 And _
                UserList(UserIndex).flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If UserList(UserIndex).flags.Muerto = 1 Then
             Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
       Case eOBJType.otMonturas
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            
            If UserList(UserIndex).flags.Navegando = 1 Then
                Call WriteConsoleMsg(UserIndex, "No disponible navegando.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call DoEquita(UserIndex, ObjIndex, Slot)
            End If

    Case eOBJType.otPasajes
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(UserIndex).flags.TargetNPC <> 0 Then
            If Left$(Npclist(UserList(UserIndex).flags.TargetNPC).name, 6) <> "Pirata" Then
                Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el Pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el Pirata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not MapaValido(obj.HastaMap) Then
            Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(UserIndex).Stats.UserSkills(5) < obj.CantidadSkill Then
            Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
        Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_CENTINELA)
        UserList(UserIndex).Stats.MinAGU = 0
        UserList(UserIndex).Stats.MinHam = 0
        UserList(UserIndex).flags.Sed = 1
        UserList(UserIndex).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)
End Select

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarObjAlquimiaConstruibles(ByVal UserIndex As Integer)

Call WriteAlquimiaObjects(UserIndex)

End Sub

Sub EnivarObjSastreriaConstruibles(ByVal UserIndex As Integer)

Call WriteSastreriaObjects(UserIndex)

End Sub
Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Or UserList(UserIndex).Pos.map = 34 Or UserList(UserIndex).Pos.map = 20 Or UserList(UserIndex).Pos.map = 1 Or UserList(UserIndex).Pos.map = 59 Or UserList(UserIndex).Pos.map = 60 Or UserList(UserIndex).Pos.map = 37 Or UserList(UserIndex).Pos.map = 208 Or UserList(UserIndex).Pos.map = 150 Or UserList(UserIndex).Pos.map = 151 Or UserList(UserIndex).Pos.map = 34 Or UserList(UserIndex).Pos.map = 217 Or UserList(UserIndex).Pos.map = 218 Or UserList(UserIndex).Pos.map = 61 Or UserList(UserIndex).Pos.map = 111 Or UserList(UserIndex).Pos.map = 112 Or UserList(UserIndex).Pos.map = 64 Or UserList(UserIndex).Pos.map = 62 Or UserList(UserIndex).Pos.map = 63 Or UserList(UserIndex).Pos.map = 179 Or UserList(UserIndex).Pos.map = 194 Or UserList(UserIndex).Pos.map = 183 Or UserList(UserIndex).Pos.map = 184 Or UserList(UserIndex).Pos.map = 185 Then Exit Sub

Call TirarTodosLosItems(UserIndex)

Dim Cantidad As Long
Cantidad = UserList(UserIndex).Stats.GLD - CLng(UserList(UserIndex).Stats.ELV) * 10000

If Cantidad > 0 Then _
    Call TirarOro(Cantidad, UserIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es drakkar y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(UserIndex).clase = eClass.drakkar And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Or UserList(UserIndex).Pos.map = 34 Or UserList(UserIndex).Pos.map = 20 Or UserList(UserIndex).Pos.map = 1 Or UserList(UserIndex).Pos.map = 59 Or UserList(UserIndex).Pos.map = 60 Or UserList(UserIndex).Pos.map = 37 Or UserList(UserIndex).Pos.map = 208 Or UserList(UserIndex).Pos.map = 150 Or UserList(UserIndex).Pos.map = 151 Or UserList(UserIndex).Pos.map = 34 Or UserList(UserIndex).Pos.map = 217 Or UserList(UserIndex).Pos.map = 218 Or UserList(UserIndex).Pos.map = 61 Or UserList(UserIndex).Pos.map = 111 Or UserList(UserIndex).Pos.map = 112 Or UserList(UserIndex).Pos.map = 64 Or UserList(UserIndex).Pos.map = 62 Or UserList(UserIndex).Pos.map = 63 Or UserList(UserIndex).Pos.map = 179 Or UserList(UserIndex).Pos.map = 194 Or UserList(UserIndex).Pos.map = 183 Or UserList(UserIndex).Pos.map = 184 Or UserList(UserIndex).Pos.map = 185 Then Exit Sub
'es0? SI YA ESTA
For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
