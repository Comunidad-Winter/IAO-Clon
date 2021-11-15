VERSION 5.00
Begin VB.Form frmbindkey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar acción"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBindKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleMode       =   0  'User
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Optaccion 
      Caption         =   "Trabajar"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   3135
   End
   Begin VB.OptionButton Optaccion 
      Caption         =   "Enviar Comando"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.OptionButton Optaccion 
      Caption         =   "Lanzar hechizo elegido"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2430
      Width           =   3135
   End
   Begin VB.TextBox txtcomandoenvio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "/"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdaccept 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OptionButton Optaccion 
      Caption         =   "Equipar ítem elegido"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   2970
      Width           =   3135
   End
   Begin VB.OptionButton Optaccion 
      Caption         =   "Usar ítem elegido"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lbltecla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3375
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   1400
      Y2              =   1400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBindKey.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmbindkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAccept_Click()
Dim i As Integer

For i = Optaccion.LBound To Optaccion.UBound
    If Optaccion(i).value = True Then
        MacroKeys(BotonElegido).TipoAccion = i + 1
        Exit For
    End If
Next i

Select Case MacroKeys(BotonElegido).TipoAccion
    
    Case 1
    If txtcomandoenvio.text = "Comando a enviar" Then
        MensajeAdvertencia ("Debes escribir un comando válido a enviar")
        Exit Sub
    End If
        
        MacroKeys(BotonElegido).SendString = UCase$(txtcomandoenvio.text)
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).invslot = 0
    
    Case 2
        MacroKeys(BotonElegido).hlist = frmMain.hlst.ListIndex + 1
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = 0

    Case 3
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = Inventario.SelectedItem
      
    Case 4
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = Inventario.SelectedItem
    
    Case 5
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).SendString = ""
        MacroKeys(BotonElegido).invslot = 0

        End Select

Call DibujarMenuMacros(BotonElegido)

Dim lC As Byte
Dim file As String

file = "macros-" & UCase$(UserName) & ".dat"

If FileExist(App.path & "/Recursos\Init\Macros\" & file, vbNormal) Then _
    Kill App.path & "/Recursos\Init\Macros\" & file
    
Open App.path & "/recursos\Init\Macros\" & file For Append As #1
    For lC = 1 To 11
        Print #1, "[BIND" & lC & "]"
        Print #1, "Accion=" & MacroKeys(lC).TipoAccion
        Print #1, "InvSlot=" & MacroKeys(lC).invslot
        Print #1, "SndString=" & MacroKeys(lC).SendString
        Print #1, "hlist=" & MacroKeys(lC).hlist
        Print #1, "Trabaja=" & MacroKeys(lC).invslot
        Print #1, "" 'Separacion entre macro y macro
    Next lC
Close #1

Unload Me
End Sub


Private Sub cmdCancel_Click()
'MacroKeys(BotonElegido).TipoAccion = 0
Unload Me
End Sub

Private Sub optAccion_Click(Index As Integer)

If Index = 0 Then
    txtcomandoenvio.Enabled = True
Else
    txtcomandoenvio.Enabled = False
End If

End Sub

Private Sub Form_Load()

'el problema que tengo es que no me chekea la cabeza y tengo del lado del sv y cliente hecho bien pero no me la u

lbltecla.Caption = "Tecla: F" & BotonElegido

If MacroKeys(BotonElegido).TipoAccion <> 0 Then

    Select Case MacroKeys(BotonElegido).TipoAccion
        Case 1 'Envia comando
            Optaccion(0).value = True
            txtcomandoenvio.text = MacroKeys(BotonElegido).SendString
            txtcomandoenvio.Enabled = True
        Case 2 'Lanza hechizo
            Optaccion(1).value = True
        Case 3 'Equipa
            Optaccion(2).value = True
        Case 4 'Usa
            Optaccion(3).value = True
         Case 5 'Usa
            Optaccion(4).value = True
    End Select
    
End If
    

End Sub
Public Sub DibujarMenuMacros(Optional ActualizarCual As Integer = 0, Optional AlphaEffect As Byte = 0)

On Error Resume Next

Dim i As Integer

If ActualizarCual <= 0 Then
    For i = 1 To 11
        frmMain.Picmacro(i - 1).Cls
        Select Case MacroKeys(i).TipoAccion
            Case 1 'Envia comando
                Call Grh_Render_To_Hdc(frmMain.Picmacro(i - 1).hdc, 17506, 0, 0)
                frmMain.Picmacro(i - 1).ToolTipText = "Enviar comando: " & MacroKeys(i).SendString
            Case 2 'Lanza hechizo
                Call Grh_Render_To_Hdc(frmMain.Picmacro(i - 1).hdc, 609, 0, 0)
                frmMain.Picmacro(i - 1).ToolTipText = "Lanzar hechizo elegido: " & frmMain.hlst.List(MacroKeys(i).hlist - 1)
            Case 3 'Equipa
                Call Grh_Render_To_Hdc(frmMain.Picmacro(i - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
                frmMain.Picmacro(i - 1).ToolTipText = "Equipar objeto: " & Inventario(MacroKeys(i).invslot).name
            Case 4 'Usa
                Call Grh_Render_To_Hdc(frmMain.Picmacro(i - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
                frmMain.Picmacro(i - 1).ToolTipText = "Usar objeto: " & Inventario(MacroKeys(i).invslot).name
            Case 5 'Trabaja
                Call Grh_Render_To_Hdc(frmMain.Picmacro(i - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
                frmMain.Picmacro(i - 1).ToolTipText = "Usar objeto: " & Inventario(MacroKeys(i).invslot).name
        End Select
        frmMain.Picmacro(i - 1).Refresh
    Next i
Else
    frmMain.Picmacro(ActualizarCual - 1).Cls
    
    Select Case MacroKeys(ActualizarCual).TipoAccion
        Case 1 'Envia comando
            Call Grh_Render_To_Hdc(frmMain.Picmacro(ActualizarCual - 1).hdc, 17506, 0, 0)
            frmMain.Picmacro(ActualizarCual - 1).ToolTipText = "Enviar comando: " & MacroKeys(ActualizarCual).SendString
        Case 2 'Lanza hechizo
            Call Grh_Render_To_Hdc(frmMain.Picmacro(ActualizarCual - 1).hdc, 609, 0, 0)
            frmMain.Picmacro(ActualizarCual - 1).ToolTipText = "Lanzar hechizo elegido: " & frmMain.hlst.List(MacroKeys(ActualizarCual).hlist - 1)
        Case 3 'Equipa
            Call Grh_Render_To_Hdc(frmMain.Picmacro(ActualizarCual - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
            frmMain.Picmacro(ActualizarCual - 1).ToolTipText = "Equipar objeto: " & Inventario(MacroKeys(ActualizarCual).invslot).name
        Case 4 'Usa
            Call Grh_Render_To_Hdc(frmMain.Picmacro(ActualizarCual - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
            frmMain.Picmacro(ActualizarCual - 1).ToolTipText = "Usar objeto: " & Inventario.ItemName(MacroKeys(ActualizarCual).invslot)
        Case 5 'trabaja
        Call Grh_Render_To_Hdc(frmMain.Picmacro(ActualizarCual - 1).hdc, 505, 0, 0)
            frmMain.Picmacro(ActualizarCual - 1).ToolTipText = "trabajar " & MacroKeys(ActualizarCual).SendString
    End Select

    frmMain.Picmacro(ActualizarCual - 1).Refresh
   
End If

End Sub

Public Sub Bind_Accion(ByVal FNUM As Integer)

If MacroKeys(FNUM).TipoAccion = 0 Then Exit Sub

Select Case MacroKeys(FNUM).TipoAccion

Case 1 'Envia comando

    Call ParseUserCommand(MacroKeys(FNUM).SendString)
    
Case 2 'Lanza hechizo
    If frmMain.hlst.List(MacroKeys(FNUM).hlist - 1) <> "Nada" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then Exit Sub
        Call WriteCastSpell(MacroKeys(FNUM).hlist)
        Call WriteWork(eSkill.magia)
        UsaMacro = True
        
     If frmMain.hlst.List(MacroKeys(FNUM).hlist - 1) = "Invisibilidad" Then
        Call WriteWorkLeftClick(UserPos.X, UserPos.Y - 1, eSkill.magia)
        Call WriteCastSpell(MacroKeys(FNUM).hlist)
        Call WriteWork(eSkill.magia)
        Call WriteWorkLeftClick(UserPos.X, UserPos.Y, eSkill.magia)
     End If
       
     If frmMain.hlst.List(MacroKeys(FNUM).hlist - 1) = "Honor del Paladín" Then
        Call WriteWorkLeftClick(UserPos.X, UserPos.Y - 1, eSkill.magia)
        Call WriteCastSpell(MacroKeys(FNUM).hlist)
        Call WriteWork(eSkill.magia)
        Call WriteWorkLeftClick(UserPos.X, UserPos.Y, eSkill.magia)
     End If
 
    End If
Case 3 'Equipa
    If UserEstado = 1 Then Exit Sub
    Call WriteEquipItem(MacroKeys(FNUM).invslot)
Case 4 'Usa
    If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteUseItem(MacroKeys(FNUM).invslot)
Case 5 'trabaja
    frmMain.macrotrabajo.Enabled = Not frmMain.macrotrabajo.Enabled
    Call frmMain.ActivarMacroTrabajo
End Select

End Sub

