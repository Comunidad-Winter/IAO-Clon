VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion de Controles."
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMSens 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   70
      Text            =   "10"
      Top             =   3120
      Width           =   495
   End
   Begin VB.HScrollBar scrSens 
      Height          =   375
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   69
      Top             =   3120
      Value           =   10
      Width           =   2805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   46
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Teclas por defecto"
      Height          =   375
      Left            =   3720
      TabIndex        =   45
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   11400
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   2280
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   2280
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   2280
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   2280
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   2280
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   2280
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   2280
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Modo Combate"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   60
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Modo Seguro"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   59
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Seg. de Resucitación"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   58
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Salir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Macro Trabajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Macro Hechizos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Meditar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Opciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar FPS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Hablar"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11640
      TabIndex        =   3
      Top             =   6000
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   1920
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar al Clan"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar a Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Caption         =   "Acciones"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7575
      Begin VB.TextBox Text9 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Esc"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "C"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "L"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "M"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Bloq Num"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "N"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "IMP PANT"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "*"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   1920
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "Salir"
         Height          =   255
         Left            =   5520
         TabIndex        =   77
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label35 
         Caption         =   "Modo Combate"
         Height          =   255
         Left            =   5520
         TabIndex        =   76
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label34 
         Caption         =   "Corregir Posicion"
         Height          =   255
         Left            =   5520
         TabIndex        =   74
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label33 
         Caption         =   "Activar/Desactivar Musica"
         Height          =   255
         Left            =   5520
         TabIndex        =   72
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label31 
         Caption         =   "Bloqueo de Movimiento"
         Height          =   255
         Left            =   3720
         TabIndex        =   68
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "Ocultar / Mostrar NICK"
         Height          =   255
         Left            =   3720
         TabIndex        =   66
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Sacar Foto"
         Height          =   255
         Left            =   3720
         TabIndex        =   64
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "Mostrar FPS"
         Height          =   255
         Left            =   3720
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Atacar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Usar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tirar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ocultar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000004&
         Caption         =   "Robar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000004&
         Caption         =   "Domar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000004&
         Caption         =   "Equipar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000004&
         Caption         =   "Agarrar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   11400
      TabIndex        =   1
      Top             =   600
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Movimiento"
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   7800
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Derecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   "Izquierda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Caption         =   "Abajo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Arriba"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label32 
      Caption         =   "Sensibilidad : "
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   2880
      Width           =   2895
   End
End
Attribute VB_Name = "frmCustomKeys"
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

Option Explicit

Private Sub Command1_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Command2_Click()

Dim i As Long

For i = 1 To CustomKeys.Count
    If LenB(Text1(i).Text) = 0 Then
        Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
        Exit Sub
    End If
Next i

Call CustomKeys.SaveCustomKeys

Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub


Private Sub scrSens_Change()
MouseS = scrSens.value
Call General_Set_Mouse_Speed(MouseS)
txtMSens.Text = scrSens.value
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
