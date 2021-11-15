Attribute VB_Name = "mod_AntiName"
Public NombreOriginal As String
Public NoCambiesNombre As String
 
Public Function NoCambio() As Boolean
If NombreOriginal <> NoCambiesNombre Then
NoCambio = True
Exit Function
End If
NoCambio = False
End Function
 
Public Sub ErrorCambiName()
MsgBox "Se ha detectado cambio de nombre en el ejecutable", vbCritical
End Sub
