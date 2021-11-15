Attribute VB_Name = "Mod_limpieza"
Option Explicit
 
Const MAXITEMS As Integer = 1000
Private ArrayLimpieza(MAXITEMS) As WorldPos
 
'//Desde acá establecemos el ultimo slot que se uso
Public UltimoSlotLimpieza As Integer
 
 
Public Sub AgregarObjetoLimpieza(Pos As WorldPos)
 
    '//Primera pos almacenada
    If UltimoSlotLimpieza = -1 Then
        ArrayLimpieza(0) = Pos
        UltimoSlotLimpieza = 0
        Exit Sub
    End If
   
    '//En caso de no ser cero, significa que ya hay objetos, seguimos sumando +1
    UltimoSlotLimpieza = UltimoSlotLimpieza + 1
   
    ArrayLimpieza(UltimoSlotLimpieza) = Pos
   
    '¿Alcanzamos los slots que permite almacenar?
    If UltimoSlotLimpieza = MAXITEMS Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpiando objetos tirados.", FontTypeNames.FONTTYPE_SERVER))
        Call BorrarObjetosLimpieza
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza de objetos finalizada.", FontTypeNames.FONTTYPE_SERVER))
        'Comenzamos de nuevo
        UltimoSlotLimpieza = 0
    End If
End Sub
 
Public Sub BorrarObjetosLimpieza()
Dim i As Long
 
    For i = 0 To MAXITEMS
        With ArrayLimpieza(i)
            Call EraseObj(10000, .map, .x, .Y)
        End With
    Next i
End Sub
 
 
