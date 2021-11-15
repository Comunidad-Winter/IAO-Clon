Attribute VB_Name = "modPacketAntiCheat"
Option Explicit
 
Private Enum modosecu
porsegundo
intervaloentrepaquetes
ambos
End Enum
 
Private Const Cantidadxsegundo As Byte = 25 'paquetes por segundo
Private Const paqinterval As Byte = 30 'milisegundos
 
Private Const modoSec As Byte = modosecu.porsegundo
'Esto lo pueden cambiar, pueden poner paquetes por segundo, intervalos entre paquetes o ambos.
 
Public Type ElsantoSec '(?
lastpaqtime As Long
Paquetes As Byte
End Type
 
Public Sub AgregarPaquete(ByVal UserIndex As Integer)
With UserList(UserIndex).controlpaqs
If modoSec = modosecu.porsegundo Or modoSec = modosecu.ambos Then
.Paquetes = .Paquetes + 1
If .Paquetes > Cantidadxsegundo Then
Call WriteErrorMsg(UserIndex, "Has sido echado por posible uso de cheats.")
Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido echado del servidor por posible uso de cheats", FontTypeNames.FONTTYPE_TALK))
Call CloseSocket(UserIndex)
End If
End If
 
If modoSec = modosecu.ambos Or modoSec = modosecu.intervaloentrepaquetes Then
Dim tiempoo As Long
tiempoo = GetTickCount
If tiempoo - .lastpaqtime > paqinterval Then
Call WriteErrorMsg(UserIndex, "Has sido echado por posible uso de cheats.")
Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido echado del servidor por posible uso de cheats", FontTypeNames.FONTTYPE_TALK))
Call CloseSocket(UserIndex)
End If
.lastpaqtime = tiempoo
End If
End With
End Sub
 
Public Sub PasarsegundoSec()
Dim d As Long
For d = 1 To LastUser
UserList(d).controlpaqs.Paquetes = 0
Next d
End Sub
 
