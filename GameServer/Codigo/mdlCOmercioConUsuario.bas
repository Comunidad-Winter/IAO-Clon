Attribute VB_Name = "mdlCOmercioConUsuario"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 1000

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo ErrHandler
Dim userindex As Integer
'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Origen)
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Destino)
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call WriteConsoleMsg(Destino, UserList(Origen).name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    

If Not UserList(userindex).flags.ModoCombate = False Then
Call WriteConsoleMsg(userindex, "Debes desactivar el modo combate para realizar esta acci�n", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If
Call FlushBuffer(Destino)

Exit Sub
ErrHandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    ObjInd = iORO
Else
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

If ObjInd > 0 And ObjCant > 0 Then
    Call WriteChangeUserTradeSlot(AQuien, ObjInd, ObjCant)
    Call FlushBuffer(AQuien)
End If

End Sub

Public Sub FinComerciarUsu(ByVal userindex As Integer)
    With UserList(userindex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(userindex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.cant = 0
        .ComUsu.DestUsu = 0
        .ComUsu.Objeto = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal userindex As Integer)
Dim Obj1 As obj, Obj2 As obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False

If UserList(userindex).ComUsu.DestUsu <= 0 Or UserList(userindex).ComUsu.DestUsu > MaxUsers Then
    TerminarAhora = True
End If

OtroUserIndex = UserList(userindex).ComUsu.DestUsu

If Not TerminarAhora Then
    If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(userindex).flags.UserLogged = False Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(OtroUserIndex).ComUsu.DestUsu <> userindex Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(OtroUserIndex).name <> UserList(userindex).ComUsu.DestNick Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(userindex).name <> UserList(OtroUserIndex).ComUsu.DestNick Then
        TerminarAhora = True
    End If
End If

If TerminarAhora = True Then
    Call FinComerciarUsu(userindex)
    
    If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
        Call FinComerciarUsu(OtroUserIndex)
        Call Protocol.FlushBuffer(OtroUserIndex)
    End If
    
    Exit Sub
End If

UserList(userindex).ComUsu.Acepto = True
TerminarAhora = False

If UserList(OtroUserIndex).ComUsu.Acepto = False Then
    Call WriteConsoleMsg(userindex, "El otro usuario aun no ha aceptado tu oferta.", FontTypeNames.FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    Obj1.ObjIndex = iORO
    If UserList(userindex).ComUsu.cant > UserList(userindex).Stats.GLD Then
        Call WriteConsoleMsg(userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(userindex).ComUsu.cant
    Obj1.ObjIndex = UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).Amount Then
        Call WriteConsoleMsg(userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    Obj2.ObjIndex = iORO
    If UserList(OtroUserIndex).ComUsu.cant > UserList(OtroUserIndex).Stats.GLD Then
        Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(userindex)
    
    Call FinComerciarUsu(OtroUserIndex)
    Call FlushBuffer(OtroUserIndex)
    Exit Sub
End If

Call FlushBuffer(OtroUserIndex)

'[CORREGIDO]
'Desde ac� correg� el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.cant
    If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " solto oro en comercio seguro con " & UserList(userindex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
    Call WriteUpdateUserStats(OtroUserIndex)
    'y se la doy al otro
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(OtroUserIndex).ComUsu.cant
    'If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " recibio oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
    'Esta linea del log es al pedo.
    Call WriteUpdateUserStats(userindex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(userindex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(userindex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
    
    'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
    If ObjData(Obj2.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(userindex).name & " " & Obj2.Amount & " " & ObjData(Obj2.ObjIndex).name)
    End If
    'Es mucha cantidad?
    If Obj2.Amount > MAX_OBJ_LOGUEABLE Then
    'Si no es de los prohibidos de loguear, lo logueamos.
        If ObjData(Obj2.ObjIndex).NoLog <> 1 Then
            Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(userindex).name & " " & Obj2.Amount & " " & ObjData(Obj2.ObjIndex).name)
        End If
    End If
End If

'pone el oro directamente en la billetera
If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - UserList(userindex).ComUsu.cant
    If UserList(userindex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(userindex).name & " solt� oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(userindex).ComUsu.cant)
    Call WriteUpdateUserStats(userindex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(userindex).ComUsu.cant
    'If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " recibio oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
    'Esta linea del log es al pedo.
    Call WriteUpdateUserStats(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj1)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, userindex)
    
    'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
    If ObjData(Obj1.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(userindex).name & " le pas� en comercio seguro a " & UserList(OtroUserIndex).name & " " & Obj1.Amount & " " & ObjData(Obj1.ObjIndex).name)
    End If
    'Es mucha cantidad?
    If Obj1.Amount > MAX_OBJ_LOGUEABLE Then
    'Si no es de los prohibidos de loguear, lo logueamos.
        If ObjData(Obj1.ObjIndex).NoLog <> 1 Then
            Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(userindex).name & " " & Obj1.Amount & " " & ObjData(Obj1.ObjIndex).name)
        End If
    End If
    
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, userindex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(userindex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]