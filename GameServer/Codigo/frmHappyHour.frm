VERSION 5.00
Begin VB.Form frmHappyHour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Happy Hour"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Finalizar"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "2"
      Top             =   720
      Width           =   1320
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Iniciar Evento"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox TipoEvent 
      Height          =   315
      ItemData        =   "frmHappyHour.frx":0000
      Left            =   120
      List            =   "frmHappyHour.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Evento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmHappyHour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHappyHour
' Author    : Shermie80
' Date      : 10/03/2015
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdEnd_Click()
    tHappyHour.Iniciado = False
    tHappyHour.tExp = 0
    tHappyHour.tOro = 0
    tHappyHour.tDrop = 0
    tHappyHour.tModo = 0
    Call SendData(ToAll, 0, PrepareMessageConsoleMsg("¡Ha finalizado el evento Happy Hour!", FontTypeNames.FONTTYPE_INFO))
    Unload Me
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(132, NO_3D_SOUND, NO_3D_SOUND))
End Sub

Public Sub IniciarHappyHour(ByVal Tipoevento As Byte, ByVal Multiplicador As Byte)
tHappyHour.Iniciado = False
tHappyHour.tExp = 1
tHappyHour.tOro = 1
tHappyHour.tDrop = 1
tHappyHour.tModo = 0

Select Case Tipoevento
    Case 1
        tHappyHour.Iniciado = True
        tHappyHour.tModo = Tipoevento
        tHappyHour.tExp = Multiplicador
        Call SendData(ToAll, 0, PrepareMessageConsoleMsg("¡Ha comenzado el evento de Experiencia x" & Multiplicador & "! En 30 minutos finalizará el evento.", FontTypeNames.FONTTYPE_INFO))
        CountTimerHP = 30
        frmHappyHour.Timer1.Enabled = True
    Exit Sub
    
    Case 2
        tHappyHour.Iniciado = True
        tHappyHour.tModo = Tipoevento
        tHappyHour.tOro = Multiplicador
        Call SendData(ToAll, 0, PrepareMessageConsoleMsg("¡Ha comenzado el evento de Oro x" & Multiplicador & "! En 30 minutos finalizará el evento.", FontTypeNames.FONTTYPE_INFO))
        CountTimerHP = 30
        frmHappyHour.Timer1.Enabled = True
    Exit Sub
    
    Case 3
        tHappyHour.Iniciado = True
        tHappyHour.tModo = Tipoevento
        tHappyHour.tExp = Multiplicador
        tHappyHour.tOro = Multiplicador
        Call SendData(ToAll, 0, PrepareMessageConsoleMsg("¡Ha comenzado el evento de Experiencia x" & Multiplicador & " y Oro x" & Multiplicador & "! En 30 minutos finalizará el evento.", FontTypeNames.FONTTYPE_INFO))
        CountTimerHP = 30
        frmHappyHour.Timer1.Enabled = True
    Exit Sub
    
    Case 4
        tHappyHour.Iniciado = True
        tHappyHour.tModo = Tipoevento
        tHappyHour.tDrop = Multiplicador
        Call SendData(ToAll, 0, PrepareMessageConsoleMsg("¡¡¡Ha comenzado el evento de Drop x" & Multiplicador & "!!! En 30 minutos finalizará el evento.", FontTypeNames.FONTTYPE_INFO))
        CountTimerHP = 30
        frmHappyHour.Timer1.Enabled = True
    Exit Sub
End Select
End Sub

Private Sub cmdEnviar_Click()
Call IniciarHappyHour(TipoEvent.ListIndex + 1, txtCantidad.Text)
Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(132, NO_3D_SOUND, NO_3D_SOUND))
Unload Me
End Sub

Private Sub Form_Load()
    TipoEvent.AddItem ("Experiencia")
    TipoEvent.AddItem ("Oro")
    TipoEvent.AddItem ("Exp y Oro")
    TipoEvent.AddItem ("¡Drop!")
    TipoEvent.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
CountTimerHP = CountTimerHP - 1

If CountTimerHP <= 0 Then
    Select Case tHappyHour.tModo
        Case 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Evento de Experiencia x" & tHappyHour.tExp & " ha finalizado.", FontTypeNames.FONTTYPE_INFO))
        Case 2
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Evento de Oro x" & tHappyHour.tOro & " ha finalizado.", FontTypeNames.FONTTYPE_INFO))
        Case 3
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Evento de Exp x" & tHappyHour.tExp & " y Oro x" & tHappyHour.tOro & " ha finalizado.", FontTypeNames.FONTTYPE_INFO))
        Case 4
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Evento de Drop x" & tHappyHour.tDrop & " ha finalizado.", FontTypeNames.FONTTYPE_INFO))
    End Select
    tHappyHour.Iniciado = False
    tHappyHour.tExp = 1
    tHappyHour.tOro = 1
    tHappyHour.tDrop = 1
    tHappyHour.tModo = 0
    Timer1.Enabled = False
End If

If CountTimerHP = 15 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> ¡Restan " & CountTimerHP & " minutos de Evento!", FontTypeNames.FONTTYPE_INFO))
End If
End Sub

