VERSION 5.00
Object = "{AF8B3B7F-5EEF-45C5-8DF5-C063AA68663C}#2.0#0"; "listadoservers.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   5010
      TabIndex        =   4
      Top             =   3735
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   420
      Left            =   5130
      TabIndex        =   3
      Top             =   3120
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   405
      Left            =   5310
      TabIndex        =   2
      Top             =   2385
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5250
      TabIndex        =   1
      Top             =   1890
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   5280
      TabIndex        =   0
      Top             =   1410
      Width           =   1770
   End
   Begin ListaServers.ListadoServers ctrListaServers1 
      Height          =   2895
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5106
      ColorSombra     =   8421504
      ColorLabel      =   65280
      ColorDireccion  =   65280
      ColorFondo      =   0
      BeginProperty TipoLetraLabels {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TipoLetraDireccion {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PunteroItems    =   2
      PunteroImagenItems=   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selected As Integer

Private Sub Command1_Click()
Randomize
ctrListaServers1.AddItem CStr(Rnd), "lala", 7666
End Sub

Private Sub Command2_Click()
ctrListaServers1.Resetear
End Sub

Private Sub Command3_Click()
ctrListaServers1.Remover selected
End Sub

Private Sub Command4_Click()
Dim i As Integer
For i = 1 To 20
    Randomize
    ctrListaServers1.AddItem CStr(Rnd), "lala", 7666
Next i
End Sub

Private Sub Command5_Click()
ctrListaServers1.ColorSombra = vbBlack
ctrListaServers1.ColorDireccion = vbBlack
ctrListaServers1.ColorLabel = vbYellow
ctrListaServers1.ColorFondo = vbRed
End Sub

Private Sub ctrListaServers1_Click(Index As Integer, item As String, direccion As String, puerto As Long)
Me.Caption = item & " (" & direccion & ":" & puerto & ")"
selected = Index
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrListaServers1.MouseOut
End Sub

Private Sub Form_Resize()
ctrListaServers1.Height = Me.ScaleHeight - 15
ctrListaServers1.Width = Me.ScaleWidth - 15
End Sub
