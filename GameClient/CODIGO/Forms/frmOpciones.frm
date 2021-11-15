VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4425
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
   ScaleHeight     =   4425
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "General"
      Height          =   2415
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox Macros 
         Caption         =   "MACROS: Activados."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ListBox lstIgnore 
         Height          =   1620
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de ignorados (click derecho: menú)"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   3705
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox chkMidi 
         Caption         =   "Reproducir midi default de la zona"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1230
         Width           =   2775
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   2385
         TabIndex        =   11
         Top             =   1845
         Width           =   345
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Efectos de navegación"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   900
         Width           =   2715
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   2715
      End
      Begin VB.HScrollBar Slider1 
         Height          =   315
         Index           =   1
         LargeChange     =   15
         Left            =   150
         Max             =   100
         SmallChange     =   2
         TabIndex        =   7
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar Slider1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         LargeChange     =   15
         Left            =   150
         Max             =   100
         SmallChange     =   2
         TabIndex        =   6
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CheckBox chkInvertir 
         Caption         =   "Invertir canales de audio (L / R)"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1530
         Width           =   2775
      End
      Begin VB.Label lblMidi 
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo midi número"
         Height          =   255
         Left            =   195
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   2160
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   2905
         Width           =   2865
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información"
      Height          =   1335
      Left            =   3480
      TabIndex        =   1
      Top             =   2520
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
         Caption         =   "http//www.eternalonline.com"
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
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3960
      Width           =   3255
   End
End
Attribute VB_Name = "frmOpciones"
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

Private loading As Boolean

Private Sub Check1_Click(Index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index
        Case 0
            If Check1(0).value = vbUnchecked Then
                Audio.MusicActivated = False
                Audio.AmbientActivated = False
                Call Audio.MusicMP3Stop
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Audio.MusicActivated = True
                Audio.AmbientActivated = True
            End If
        Case 1
            If Check1(1).value = vbUnchecked Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
            Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).value = Audio.SoundVolume
            End If
    End Select
End Sub

Private Sub chkop_Click(Index As Integer)
    Nombres = Not Nombres
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdManual_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://www.Eternal-Online.com.ar/", "", App.path, 0)
End Sub

Private Sub cmdCerrar_Click()
    Me.Visible = False
End Sub

Private Sub customMsgCmd_Click()
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub cmdWeb_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://www.Eternal-Online.com.ar/", "", App.path, 0)
End Sub

Private Sub Form_Load()
On Error Resume Next
    loading = True      'Prevent sounds when setting check's values
    
    If Audio.MusicActivated Then
        Check1(0).value = vbChecked
        Slider1(0).Enabled = True
        Slider1(0).value = Audio.MusicVolume
    Else
        Check1(0).value = vbUnchecked
        Slider1(0).Enabled = False
    End If
    
    If Audio.SoundActivated Then
        Check1(1).value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).value = Audio.SoundVolume
    Else
        Check1(1).value = vbUnchecked
        Slider1(1).Enabled = False
    End If
    
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub Macros_Click()
Dim i As Integer
If frmOpciones.Macros.value = 0 Then
frmOpciones.Macros.Caption = "MACROS: Desactivados."
For i = 1 To 11
frmMain.Macros(i).Visible = False
Next i
Else
frmOpciones.Macros.Caption = "MACROS: Activados."

For i = 1 To 11
frmMain.Macros(i).Visible = True
Next i
End If

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
