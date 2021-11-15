VERSION 5.00
Begin VB.Form frmHlp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ayuda"
   ClientHeight    =   6225
   ClientLeft      =   2355
   ClientTop       =   1845
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHlp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296.605
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFAQ 
      Caption         =   "Preguntas &frecuentes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2130
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdBPage 
      Caption         =   "< Página &anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdNPage 
      Caption         =   "Página &siguiente >"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Ayuda 
      Caption         =   "Introducción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label lblHlp 
         Caption         =   $"frmHlp.frx":0442
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5085
      End
   End
End
Attribute VB_Name = "frmHlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EnQuePagina As Integer

Private Sub cmdFAQ_Click()
    lblHlp.Caption = "."
End Sub

Private Sub cmdBPage_Click()
    EnQuePagina = EnQuePagina - 1
    Call CambioPagina
End Sub

Private Sub cmdNPage_Click()
    EnQuePagina = EnQuePagina + 1
    Call CambioPagina
End Sub

Private Sub CambioPagina()

Select Case EnQuePagina

Case 0
    Ayuda.Caption = "Introducción"
    lblHlp.Caption = "En ImperiumAO encontrarás todo un nuevo mundo por explorar, sin fronteras, en el que no hay profesión que sobresalga sobre otra y aún el ambiente se mantiene vivo. A continuación podés encontrar algunas de las preguntas más frecuentes que se realizan los viajeros de estas tierras."
    cmdBPage.Visible = False
    cmdNPage.Visible = True
Case 1
    Ayuda.Caption = "Empezando a jugar"
    lblHlp.Caption = "En un principio, tendrás, como mínimo, ropa, agua, comida, un mapa y un arma. Un buen lugar para empezar es el Newbie Dungeon, que posee dos zonas. En la primera encontrarás tenebrosas criaturas que te desafiarán, en la segunda una zona en donde entrenar tus habilidades de combate. Estando un alguna de las ciudades iniciales, podrás tomar el teleport al mismo fácilmente. Mientras subas de nivel, irás ganando oro, con lo que puedes comprar hechizos, armas, ropa, etc. Cuando tu nivel sea 13 o superior, dejarás de ser newbie y ya no ganarás oro al subir de nivel. Deberás ir pensando en acceder a los bosques, los dungeons, y todo el mundo que te espera."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 2
    Ayuda.Caption = "Entrenando"
    lblHlp.Caption = "Cada clase es única. Al haber elegido cierta clase, estas sujeto a ser mejor en algunas áreas que en otras. Aprovéchalas. Si eres mago, dedícate a la magia. Si eres guerrero, dedícate al combate con armas. Cuando subas de nivel, ganarás skillpoints, los cuales puedes decidir utilizar o no. Lo más conveniente es dejarlos para cuando seas de mayor nivel. Por ahora, mientras entrenes, irás ganando naturalmente los skillpoints. A mayor skill en tal área, mejor es el rendimiento que tendrá tu personaje al realizar la actividad. Explora: es la mejor manera de conocer el mundo en el que vive tu personaje. Encontrarás miles de criaturas, nuevos retos y gente con quien afrontarlos. Los mejores lugares para comenzar a entrenar luego de haber dejado el Newbie Dungeon son el Bosque Dorck (Mapas 39 y 38) y las Colinas de Nix (Mapa 35 y alrededores) En estas zonas encontrarás tus primeros retos: Arañas Gigantes y Zombies."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 3
    Ayuda.Caption = "Interactuando"
    lblHlp.Caption = "A lo largo de tus viajes, encontrarás muchas maneras de interactuar con el mundo en el que tu personaje vive. Encontrarás ciudades, dungeons, bosques, y muchos lugares que llamarán tu atención. Si estás en una ciudad una buena idea es, si es que ésta tiene un puerto o río cercano, es pescar. Puedes adquirir una caña o red de pesca en el gremio de pescadores. Otra opción es talar. A medida que explores las diferentes maneras de ganarte la vida, decidirás cuál crees como más apropiada. No dejes de probar la minería y herrería, pues es una actividad muy productiva e útil si es que tu personaje tiene habilidades de combate. En ImperiumAO estas áreas han sido rebalanceadas para obtener, así, un mundo más justo y próspero."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 4
    Ayuda.Caption = "Facciónes"
    lblHlp.Caption = "En la medida de que llegues a niveles superiores, podrás decidir enlistarte o no en una facción. Puedes elegir las Tropas Reales o bien las Legiones del Caos. Para ingresar en las Tropas Reales debes haber matado, como mínimo, cincuenta criminales. Si en algún momento asesinaste gente inocente no serás admitido. Tu nivel no deberá ser inferior a 20. Por otro lado, si tu deseo es ingresar en las Legiones del Caos, deberás haber matado como mínimo 125 ciudadanos y ser nivel 25 o superior. Cuando llegues a la centena de criminales o ciudadanos asesinados, serás recompensado y, además de ganar algo de experiencia, subirás un rango de jerarquía, lo que, además de posibilitarte nuevos ítems especiales, hará que tu rango sea superior. Recuerda, no hay vuelta atrás a la hora de seleccionar una facción. Se cuidadoso a la hora de elegir."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 5
    Ayuda.Caption = "Magia"
    lblHlp.Caption = "ImperiumAO fue creado ni mas ni menos que con la quinta esencia, lo que hace a la magia y a lo desconocido. Aprovéchala. Está en todas partes. Pues, entonces, se han desarrollado nuevos tipos de hechizos. Los hechizos de metamorfosis permiten que el cuerpo de tu personaje cambie a determinada forma, temporalmente. Esto no sólo te da ventajas técnicas sino también de combate. Si aparentas ser más poderoso que una criatura, esta simplemente te ignorará. Puedes conseguir estos hechizos por precios razonables en la ciudad de Nueva Esperanza. Por otro lado, la magia se ha expandido tanto, que hay no sólo hechizos de metamorfosis, sino también de materialización, y mucho más que tienes por explorar. Se dice que hay magos con poderes tales, que pueden abrir portales planares con su magia para viajar entre el espacio. Muchos otros dicen haber visto magia con poder suficiente para matar hasta al más poderoso guerrero de estas tierras."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 6
    Ayuda.Caption = "Viajando"
    lblHlp.Caption = "Hay, pues, varias maneras no convencionales de viajar hacia otras tierras. Si hay un puerto en las cercanías y quieres llegar rápido a tu destino ¿Por qué no tomar un barco? Es muy fácil, sólo debes acercarte el muelle, comerciar con el pirata y luego mostrarle el pasaje, así, llegarás al destino elegido. Los pases no caen al morir y no pueden ser robados. Lleva un mapa. Te salvará un situaciónes comprometedoras. Recuerda que para viajar a zonas lejanas debes poder usar una Barca."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 7
    Ayuda.Caption = "Objetos mágicos"
    lblHlp.Caption = "Anillos, brazaletes, botas, armas, armaduras, escudos y todo tipo de extraños objetos han sido impregnados de magia por los más poderosos y Arcanos hechizeros. Ellos, pues, han dispersado esos codiciados ítems por todas estas tierras. Actualmente en poder de las criaturas más poderosas, estos objetos pueden darte propiedades mágicas que van desde la inmunidad a hechizos, hasta el aumento del tiempo en que recuperas tus puntos de mana. Al morir una de estas criaturas, si tienes suerte, podrás recibir uno de estos objetos."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 8
    Ayuda.Caption = "Cazando Recompensas"
    lblHlp.Caption = "En la ciudad de Rinkel no sólo se han establecido piratas, sino todo un mercado negro organizado, que está constantemente trabajando por la gente que necesita, por ejemplo, que alguien sea asesinado a toda costa. Revisa la ciudad, allí encontrarás un cazarecompensas con el que podrás, además de ofrecer dinero por la cabeza de un usuarios, ver quiénes son buscados y así poder ganar dinero de una manera fácil y efectiva."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 9
    Ayuda.Caption = "Familiares y mascotas permanentes"
    lblHlp.Caption = "Los familiares son criaturas de otros planos que, gracias a tu poder mágico, han decidido permancer junto a ti. Durante la creación de tu personaje, si eres Mago, podrás elegir un Familiar. Por otro lado, si eres Cazador o Druida, tendrás la posibilidad de tener una mascota permanente. Los familiares y mascotas permanentes tienen niveles de vida, mueren, revivien, se hacen más fuertes entrenando, exactamente como lo haces vos. Una de estas criaturas será tu fiel seguidor y te ayudará en los momentos más difíciles. Si tienes un familiar, podrás ver sus estadísticas al ver las tuyas, haciendo click en el Boton que se encuentra en el menú inferior. Todos ganan habilidades especiales al llegar a diferentes niveles."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 10
    Ayuda.Caption = "Nuevas fronteras"
    lblHlp.Caption = "Nuevos ámbitos se han descubierto. Si eres un Bardo o tu conocimiento en Artes Marciales es 10 o superior, al combatir sin armas, tendrás una probabilidad (que depende de tu skill) de hacer movimientos especiales al atacar. Estos movimientos terminan en efectos que van desde parálisis, ceguera, estupidez, hasta el desarmado de la víctima (la víctima perderá su arma) Al recibir daño con hechizos, tu personaje irá aprendiendo diferentes modos de resisitirlo. A mayor resistencia, menor daño te harán los hechizos. Shurikens, lanzas, jabalinas, dagas arrojadizas, y muchísimas otras armas te esperan. Con tener sólo las municiones alcanza, no necesitas nada más."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 11
    Ayuda.Caption = "Macros Integrados"
    lblHlp.Caption = "Se han integrado en el cliente una serie de teclas para mejorar la dinámica del juego. Su uso es completamente legal. Está TOTALMENTE prohibido el uso de cualquier programa externo al juego. Para poder utilizar estas teclas, prueba presionando F1, F2, F3, etc..."
    cmdBPage.Visible = True
    cmdNPage.Visible = True
Case 12
    Ayuda.Caption = "Game Masters"
    lblHlp.Caption = "Este mundo se encuentra regido por dioses, quienes mantienen el órden y están constantemente trabajando por su mejoría. Puedes invocarlos escribiendo /GM. De este modo, los dioses acudirán en tu ayuda. Los mismos son fácilmente reconocibles, su nombre, habla, y descripción se encuentra en verde."
    cmdBPage.Visible = True
    cmdNPage.Visible = False
End Select
End Sub
