VERSION 5.00
Begin VB.Form frmLauncher 
   BorderStyle     =   0  'None
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmLauncher.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox ComboRes 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "800x600"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.ComboBox Tuwindows 
      Enabled         =   0   'False
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Intel R"
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Image correroff 
      Height          =   375
      Left            =   720
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image CorrerOn 
      Height          =   375
      Left            =   720
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgnotas 
      Height          =   615
      Left            =   6120
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Image imgpres 
      Height          =   615
      Left            =   4680
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Image imgmanual 
      Height          =   615
      Left            =   3120
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgforo 
      Height          =   615
      Left            =   1680
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image imgverweb 
      Height          =   615
      Left            =   237
      Top             =   5160
      Width           =   1450
   End
   Begin VB.Image imgsolucionarproblemas 
      Height          =   975
      Left            =   5160
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Image MenosP 
      Height          =   135
      Left            =   3840
      Top             =   3720
      Width           =   135
   End
   Begin VB.Image masP 
      Height          =   135
      Left            =   3840
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Precarga 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   315
   End
   Begin VB.Image Imgjugar 
      Height          =   615
      Left            =   5640
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Image imgsalir 
      Height          =   615
      Left            =   240
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a AOshao! Haz click en iniciar juego para jugar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   6120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer
Private cIniciarJuego As ClsGraphicalButton
Private cSalir As ClsGraphicalButton
Private cForo As ClsGraphicalButton
Private cVerWeb As ClsGraphicalButton
Private cNotasversion As ClsGraphicalButton
Private cSolucionarProblemas As ClsGraphicalButton
Private cManual As ClsGraphicalButton
Private cpreguntasfrecuentes As ClsGraphicalButton


Public LastButtonPressed As ClsGraphicalButton

Private picCheckBox As Picture

Private Sub CorrerOn_Click()
CorrerOn.Picture = picCheckBox
End Sub

Private Sub Imgjugar_Click()
Call Main
End Sub
Private Sub imgjugar_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = "Haz click aquí para jugar Aoshao."
Info.font.bold = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = "Bienvenidos a Aoshao. Haz click en Iniciar Juego para jugar."
Info.font.bold = True
LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgsalir_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = "Volver al escritorio de Windows."
Info.font.bold = False
End Sub
Private Sub Imgsalir_click()
End
End Sub

Private Sub masP_Click()
m_Increment = 1
Tmrnumber.Interval = 500
Tmrnumber.Enabled = True
End Sub

Private Sub MenosP_Click()
m_Increment = -1
Tmrnumber.Interval = 500
Tmrnumber.Enabled = True
End Sub

Private Sub tmrNumber_Timer()
Const MIN_NUMBER = 1
Const MAX_NUMBER = 4

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If
    Precarga = format$(m_Number)
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        Tmrnumber.Interval = m_Interval
    End If
End Sub


Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = DirInterface

    Set cIniciarJuego = New ClsGraphicalButton
    Set cSalir = New ClsGraphicalButton
    Set cForo = New ClsGraphicalButton
    Set cVerWeb = New ClsGraphicalButton
    Set cManual = New ClsGraphicalButton
    Set cNotasversion = New ClsGraphicalButton
    Set cManual = New ClsGraphicalButton
    Set cpreguntasfrecuentes = New ClsGraphicalButton
    Set cSolucionarProblemas = New ClsGraphicalButton
    
    Set LastButtonPressed = New ClsGraphicalButton
    
    Call cIniciarJuego.Initialize(Imgjugar, GrhPath & "ImgJugar.jpg", _
                                    GrhPath & "iniciarover.jpg", _
                                    GrhPath & "iniciardown.jpg", Me)
                                    
    Call cSalir.Initialize(imgsalir, GrhPath & "Imgsalir.jpg", _
                                    GrhPath & "salirover.jpg", _
                                    GrhPath & "salirdown.jpg", Me)
                                    
    Call cVerWeb.Initialize(imgverweb, GrhPath & "ImgVerweb.jpg", _
                                    GrhPath & "sitioover.jpg", _
                                    GrhPath & "sitiodown.jpg", Me)
                                    
    Call cNotasversion.Initialize(imgnotas, GrhPath & "ImgNotas.jpg", _
                                    GrhPath & "NotasOver.jpg", _
                                    GrhPath & "NotasDown.jpg", Me)
                                    
    Call cForo.Initialize(imgforo, GrhPath & "ImgForo.jpg", _
                                    GrhPath & "foroover.jpg", _
                                    GrhPath & "forodown.jpg", Me)
                                    
                                    
    Call cpreguntasfrecuentes.Initialize(imgpres, GrhPath & "ImgPreg.jpg", _
                                    GrhPath & "faqover.jpg", _
                                    GrhPath & "faqdown.jpg", Me)
                                    
    Call cSolucionarProblemas.Initialize(imgsolucionarproblemas, GrhPath & "ImgSolucionar.jpg", _
                                    GrhPath & "SolucOver.jpg", _
                                    GrhPath & "SolucDown.jpg", Me)
                                    
    Call cManual.Initialize(imgmanual, GrhPath & "ImgManual.jpg", _
                                    GrhPath & "manualover.jpg", _
                                    GrhPath & "manualdown.jpg", Me)
                                    
    Set picCheckBox = LoadPicture(GrhPath & "CorrerOn.jpg")
End Sub
Private Sub Form_Load()
LoadButtons
CorrerOn.Picture = Nothing
End Sub

