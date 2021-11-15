VERSION 5.00
Begin VB.Form frmMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar acción"
   ClientHeight    =   3765
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
   ScaleHeight     =   3765
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Salir 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3270
      Width           =   1455
   End
   Begin VB.CommandButton Guardar 
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
      Left            =   1800
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Text            =   "/"
      Top             =   2070
      Width           =   3135
   End
   Begin VB.OptionButton Accion3 
      Caption         =   "Equipar ítem elegido"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   3135
   End
   Begin VB.OptionButton Accion4 
      Caption         =   "Usar ítem elegido"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2700
      Width           =   3135
   End
   Begin VB.OptionButton Accion2 
      Caption         =   "Lanzar hechizo elegido"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2430
      Width           =   3135
   End
   Begin VB.OptionButton Accion1 
      Caption         =   "Enviar Comando"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label MacroLbl 
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
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
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
      TabIndex        =   8
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Sub Accion1_Click()
Text1.Enabled = True
End Sub
 
Private Sub Accion2_Click()
Text1.Enabled = False
End Sub
 
Private Sub Accion3_Click()
Text1.Enabled = False
End Sub
 
Private Sub Accion4_Click()
Text1.Enabled = False
End Sub

Private Sub Guardar_Click()
'Usar Item
    If Accion4.value Then
        If Inventario.OBJIndex(Inventario.SelectedItem) = 0 Or _
           UsarYequiparObjValido(Inventario.OBJType(Inventario.SelectedItem), True) = False Then
            Call MsgBox("Item Invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aUsar
            MacroList(MacroIndex).Grh = Inventario.grhindex(Inventario.SelectedItem)
            MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
            MacroList(MacroIndex).OBJIndex = Inventario.OBJIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).slot = Inventario.SelectedItem
            Call SaveMacros(UserName)
           ' Call frmMain.ActualizarMacros(CByte(MacroIndex), False)
           frmMain.Macros(MacroIndex).Cls
           Call Grh_Render_To_Hdc(frmMain.Macros(MacroIndex).hdc, Inventario.grhindex(Inventario.SelectedItem), 0, 0)
            Unload Me
        End If
    End If
 
    'Equipar Item
    If Accion3.value Then
        If Inventario.OBJIndex(Inventario.SelectedItem) = 0 Or _
           UsarYequiparObjValido(Inventario.OBJType(Inventario.SelectedItem), False) = False Then
            Call MsgBox("Item Invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aEquipar
            MacroList(MacroIndex).Grh = Inventario.grhindex(Inventario.SelectedItem)
            MacroList(MacroIndex).Nombre = Inventario.ItemName(Inventario.SelectedItem)
            MacroList(MacroIndex).OBJIndex = Inventario.OBJIndex(Inventario.SelectedItem)
            MacroList(MacroIndex).slot = Inventario.SelectedItem
            Call SaveMacros(UserName)
            'Call frmMain.ActualizarMacros(CByte(MacroIndex), False)
            frmMain.Macros(MacroIndex).Cls
            Call Grh_Render_To_Hdc(frmMain.Macros(MacroIndex).hdc, Inventario.grhindex(Inventario.SelectedItem), 0, 0)
            Unload Me
        End If
    End If
 
    'Usar comandos/Hablar
    If Accion1.value Then
        If Text1.Text = "" Then
            Call MsgBox("Escriba un comando o una palabra.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aComando
            MacroList(MacroIndex).Grh = 17506
            MacroList(MacroIndex).Nombre = Text1.Text
            Call SaveMacros(UserName)
            'Call frmMain.ActualizarMacros(CByte(MacroIndex), False)
            frmMain.Macros(MacroIndex).Cls
            Call Engine.DrawGrhToHdc(frmMain.Macros(MacroIndex).hdc, 17506, 0, 0)
            Unload Me
        End If
    End If
 
    'Usar Hechizo
    If Accion2.value Then
        If frmMain.hlst.List(frmMain.hlst.ListIndex) = "(None)" Or _
           frmMain.hlst.ListIndex = -1 Then
            Call MsgBox("Hechizo invalido,seleccione otro.", vbCritical + vbOKOnly)
        Else
            MacroList(MacroIndex).mTipe = eMacros.aLanzar
            MacroList(MacroIndex).Grh = 609
            MacroList(MacroIndex).Nombre = frmMain.hlst.List(frmMain.hlst.ListIndex)
            MacroList(MacroIndex).SpellSlot = frmMain.hlst.ListIndex
            Call SaveMacros(UserName)
            'Call frmMain.ActualizarMacros(CByte(MacroIndex), False)
            frmMain.Macros(MacroIndex).Cls
            Call Engine.DrawGrhToHdc(frmMain.Macros(MacroIndex).hdc, 609, 0, 0)
            Unload Me
        End If
    End If
End Sub


Private Sub salir_Click()
    Unload Me
End Sub


