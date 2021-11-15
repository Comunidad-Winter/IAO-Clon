VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Apariencia y performance"
      Height          =   2865
      Left            =   3480
      TabIndex        =   23
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox chkop 
         Caption         =   "Ver nombre del mapa"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   27
         Top             =   300
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Ver diálogos en la consola"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   26
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Ver nombres de los jugadores"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   25
         Top             =   840
         Width           =   2715
      End
      Begin VB.ListBox lstSkin 
         Height          =   1230
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1410
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Skins instalados"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Top             =   1200
         Width           =   2925
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "General"
      Height          =   3375
      Left            =   3480
      TabIndex        =   18
      Top             =   3090
      Width           =   3285
      Begin VB.CheckBox chkop 
         Caption         =   "Uso inteligente de consola"
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Habilitar mensajes globales"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   20
         Top             =   570
         Width           =   2715
      End
      Begin VB.ListBox lstIgnore 
         Height          =   2010
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1140
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de ignorados (click derecho: menú)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   900
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6090
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   4065
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
      End
      Begin VB.CheckBox chkMidi 
         Caption         =   "Reproducir midi default de la zona"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1230
         Width           =   2775
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   2385
         TabIndex        =   9
         Top             =   1845
         Width           =   345
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Efectos de navegación"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   2715
      End
      Begin VB.CheckBox chkInvertir 
         Caption         =   "Invertir canales de audio (L / R)"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1530
         Width           =   2775
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
      End
      Begin VB.Label lblMidi 
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo midi número"
         Height          =   255
         Left            =   195
         TabIndex        =   16
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label lblBackMidi 
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2265
         TabIndex        =   15
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label lblNextMidi 
         Caption         =   "»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   14
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   2190
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   2790
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   3360
         Width           =   2865
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información"
      Height          =   1335
      Left            =   135
      TabIndex        =   1
      Top             =   4290
      Width           =   3255
      Begin VB.CommandButton cmdManual 
         Caption         =   "¿Necesitás &ayuda?"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&www.imperiumao.com.ar"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdCustomKeys 
      Caption         =   "Con&figuración de controles"
      Height          =   360
      Left            =   150
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5700
      Width           =   3255
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.1 MENDUZ DX8 VERSION www.noicoder.com
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Argentum Online is based on Baronsoft's VB6 Online RPG
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

Private loading As Boolean

Private Sub chkop_Click(Index As Integer)
Select Case Index
Case 0
    If chkop(0).value = vbUnchecked Then
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).value = Audio.MusicVolume
            End If
     Case 1
            If chkop(1).value = vbUnchecked Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
            Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).value = Audio.SoundVolume
            End If
    Case 2
    
     If FxNavega = 1 Then
            FxNavega = 0
        Else
            FxNavega = 1
        End If
        
     Case 3
      Nombres = Not Nombres
    End Select
      
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdManual_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/aomanual/", "", App.path, 0)
End Sub

Private Sub cmdCerrar_Click()
    Me.Visible = False
End Sub

Private Sub customMsgCmd_Click()
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub Form_Load()
On Error Resume Next
    loading = True      'Prevent sounds when setting check's values
    
    If Audio.MusicActivated Then
        chkop(0).value = vbChecked
        Slider1(0).Enabled = True
        Slider1(0).value = Audio.MusicVolume
    Else
        chkop(0).value = vbUnchecked
        Slider1(0).Enabled = False
    End If
    
    If Audio.SoundActivated Then
        chkop(1).value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).value = Audio.SoundVolume
    Else
        chkop(1).value = vbUnchecked
        Slider1(1).Enabled = False
    End If
    
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
        Case 1
            Audio.SoundVolume = Slider1(1).value
    End Select
End Sub
