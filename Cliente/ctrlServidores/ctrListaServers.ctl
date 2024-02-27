VERSION 5.00
Begin VB.UserControl ListadoServers 
   BackColor       =   &H00000000&
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   Begin VB.VScrollBar VScroll 
      Height          =   3825
      LargeChange     =   37
      Left            =   4470
      SmallChange     =   17
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   292
      TabIndex        =   0
      Top             =   0
      Width           =   4380
      Begin VB.Image ImagenClick 
         Height          =   555
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   4725
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "No hay servidores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   4335
      End
      Begin VB.Label LabelDireccion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   4665
      End
      Begin VB.Label LabelSombra 
         BackStyle       =   0  'Transparent
         Caption         =   "No hay servidores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   4335
      End
   End
End
Attribute VB_Name = "ListadoServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type tServers
    item As String
    server As String
    puerto As Long
End Type

Private servers As Collection
Private lstServers() As tServers

Public Event Click(Index As Integer, item As String, direccion As String, puerto As Long)
Attribute Click.VB_Description = "Evento que se lanza al clikear sobre una opci�n de la lista"

Private Sub ImagenClick_Click(Index As Integer)
    On Error GoTo fin
    If servers(Index + 1) = "" Then Exit Sub
    RaiseEvent Click((Index + 1), lstServers(servers.item(Index + 1)).item, lstServers(servers.item(Index + 1)).server, lstServers(servers.item(Index + 1)).puerto)
fin:
    On Error GoTo 0
End Sub

Private Sub ImagenClick_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To Label.Count - 1
        If Not i = Index Then LabelSombra(i).Visible = False
    Next i
    
    LabelSombra(Index).Visible = True
End Sub

Private Sub picScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseOut
End Sub

Private Sub picScroll_Resize()
    Dim i As Integer
    For i = 0 To LabelDireccion.Count - 1
        LabelDireccion(i).Width = picScroll.ScaleWidth - 3
        ImagenClick(i).Width = picScroll.ScaleWidth
    Next i
End Sub

Private Sub UserControl_Initialize()
    Set servers = New Collection
    ReDim lstServers(0)
End Sub

Public Function AddItem(ByVal item As String, ByVal server As String, ByVal puerto As Long)
Attribute AddItem.VB_Description = "Agrega un nuevo Item a la lista"
    'Hice IFs anidados porque todo junto tiraba true no se pq
    If UBound(lstServers) = 0 Then
        If lstServers(0).item <> "" Then ReDim Preserve lstServers(UBound(lstServers) + 1)
    Else
        ReDim Preserve lstServers(UBound(lstServers) + 1)
    End If
    lstServers(UBound(lstServers)).item = item
    lstServers(UBound(lstServers)).server = server
    lstServers(UBound(lstServers)).puerto = puerto
    
    servers.Add UBound(lstServers)
    'If ImagenClick.Count > 1 Or Not ImagenClick(0).Tag = "" Then Load ImagenClick(ImagenClick.Count)
    'ImagenClick(ImagenClick.Count - 1).Tag = UBound(lstServers)
    
    MostrarServers
    NecesitaScroll
End Function

Private Sub MostrarServers(Optional reload As Boolean)
    Dim i As Integer, actserv As Integer
    
    ' Resetea el listado actual ^_^
    Label(0).Caption = ""
    LabelSombra(0).Caption = ""
    LabelDireccion(0).Caption = ""
    If reload Then
        For i = 1 To Label.Count - 1
            Unload ImagenClick(i)
            Unload Label(i)
            Unload LabelSombra(i)
            Unload LabelDireccion(i)
        Next i
    End If
    
    On Error Resume Next
    
    For i = 0 To servers.Count - 1
        actserv = servers(i + 1)
        
        If i > 0 Then
            Load ImagenClick(i)
            Load Label(i)
            Load LabelSombra(i)
            Load LabelDireccion(i)
        End If
        
        Label(i).Caption = lstServers(actserv).item
        LabelSombra(i).Caption = lstServers(actserv).item
        LabelDireccion(i).Caption = lstServers(actserv).server & ":" & lstServers(actserv).puerto
            
        If Not Label(i).Visible Then Label(i).Visible = True
        If Not LabelSombra(i).Visible Then LabelSombra(i).Visible = False
        If Not LabelDireccion(i).Visible Then LabelDireccion(i).Visible = True
        If Not ImagenClick(i).Visible Then ImagenClick(i).Visible = True
            
        Label(i).Top = (37 * (i)) + 3
        LabelSombra(i).Top = (37 * (i)) + 2
        LabelDireccion(i).Top = (37 * (i)) + 20
        ImagenClick(i).Top = (37 * (i))
    Next i
    picScroll.Height = IIf(i = 0, 37, 37 * i)
    
    On Error GoTo 0
    
    If Label(0).Caption = "" Then GoTo noservers
    
    Exit Sub
    
noservers:
    Label(0).Visible = True
    Label(0).Caption = "No hay servidores"
    LabelSombra(0).Caption = Label(0).Caption
    LabelDireccion(0).Caption = ""
End Sub

Public Sub MouseOut()
Attribute MouseOut.VB_Description = "Esta subrutina se debe ejecutar cuando se lanza un evento MouseMove fuera del control, quita las sombras"
    Dim i As Integer
    For i = 0 To Label.Count - 1
        LabelSombra(i).Visible = False
    Next i
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseOut
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.ColorSombra = PropBag.ReadProperty("ColorSombra", Me.ColorSombra)
    Me.ColorLabel = PropBag.ReadProperty("ColorLabel", Me.ColorLabel)
    Me.ColorDireccion = PropBag.ReadProperty("ColorDireccion", Me.ColorDireccion)
    Me.ColorFondo = PropBag.ReadProperty("ColorFondo", Me.ColorFondo)
    
    Me.TipoLetraLabels = PropBag.ReadProperty("TipoLetraLabels", Me.TipoLetraLabels)
    Me.TipoLetraDireccion = PropBag.ReadProperty("TipoLetraDireccion", Me.TipoLetraDireccion)
    
    Me.PunteroItems = PropBag.ReadProperty("PunteroItems", Me.PunteroItems)
    Me.PunteroImagenItems = PropBag.ReadProperty("PunteroImagenItems", Me.PunteroImagenItems)
End Sub

Private Sub UserControl_Resize()
    If NecesitaScroll Then
        picScroll.Width = UserControl.ScaleWidth - VScroll.Width
    Else
        picScroll.Width = UserControl.ScaleWidth
    End If
End Sub

Public Function Resetear()
Attribute Resetear.VB_Description = "Resetea toda la lista"
    Dim i As Integer
    For i = 1 To servers.Count
        servers.Remove 1
    Next i
    ReDim lstServers(0)
    
    MostrarServers True
    NecesitaScroll
End Function

Private Function NecesitaScroll()
    If picScroll.Height > UserControl.ScaleHeight Then
        NecesitaScroll = True
        If VScroll.Visible = False Then VScroll.Visible = True
        VScroll.Height = UserControl.ScaleHeight
        VScroll.Left = UserControl.ScaleWidth - VScroll.Width
        VScroll.Max = picScroll.Height - UserControl.ScaleHeight
        VScroll.Min = 0
        VScroll.Value = 0
        picScroll.Width = UserControl.ScaleWidth - VScroll.Width
    Else
        NecesitaScroll = False
        If VScroll.Visible Then VScroll.Visible = False
        picScroll.Width = UserControl.ScaleWidth
    End If
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ColorSombra", ColorSombra
    PropBag.WriteProperty "ColorLabel", ColorLabel
    PropBag.WriteProperty "ColorDireccion", ColorDireccion
    PropBag.WriteProperty "ColorFondo", ColorFondo
    
    PropBag.WriteProperty "TipoLetraLabels", TipoLetraLabels
    PropBag.WriteProperty "TipoLetraDireccion", TipoLetraDireccion
    
    PropBag.WriteProperty "PunteroItems", PunteroItems
    PropBag.WriteProperty "PunteroImagenItems", PunteroImagenItems
End Sub

Private Sub VScroll_Change()
    picScroll.Top = -VScroll.Value
End Sub

Public Sub Remover(ByVal Index As Integer)
Attribute Remover.VB_Description = "Remueve determinado item de la lista"
    lstServers(servers(Index)).item = ""
    lstServers(servers(Index)).server = ""
    lstServers(servers(Index)).puerto = 0
    
    servers.Remove Index
    
    MostrarServers True
    NecesitaScroll
End Sub

Public Function Contar()
Attribute Contar.VB_Description = "Cuenta la cantidad de registros en la lista"
    Contar = servers.Count
End Function

Public Property Get ColorSombra() As OLE_COLOR
Attribute ColorSombra.VB_Description = "Modifica el color de las sombras"
Attribute ColorSombra.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ColorSombra = LabelSombra(0).ForeColor
End Property

Public Property Let ColorSombra(ByVal color As OLE_COLOR)
    Dim i As Integer
    For i = 0 To LabelSombra.Count - 1
        LabelSombra(i).ForeColor = color
    Next i
    PropertyChanged "ColorSombra"
End Property

Public Property Get ColorLabel() As OLE_COLOR
Attribute ColorLabel.VB_Description = "Modifica el color de los nombres de los servidores"
Attribute ColorLabel.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ColorLabel = Label(0).ForeColor
End Property

Public Property Let ColorLabel(ByVal color As OLE_COLOR)
    Dim i As Integer
    For i = 0 To Label.Count - 1
        Label(i).ForeColor = color
    Next i
    PropertyChanged "ColorLabel"
End Property

Public Property Get ColorDireccion() As OLE_COLOR
Attribute ColorDireccion.VB_Description = "Modifica el color de la direcci�n IP y puertos"
Attribute ColorDireccion.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ColorDireccion = LabelDireccion(0).ForeColor
End Property

Public Property Let ColorDireccion(ByVal color As OLE_COLOR)
    Dim i As Integer
    For i = 0 To LabelDireccion.Count - 1
        LabelDireccion(i).ForeColor = color
    Next i
    PropertyChanged "ColorDireccion"
End Property

Public Property Get ColorFondo() As OLE_COLOR
Attribute ColorFondo.VB_Description = "Modifica el color de fondo del control"
Attribute ColorFondo.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ColorFondo = UserControl.BackColor
End Property

Public Property Let ColorFondo(ByVal color As OLE_COLOR)
    UserControl.BackColor = color
    picScroll.BackColor = color
    PropertyChanged "ColorFondo"
End Property

Public Property Get TipoLetraLabels() As StdFont
Attribute TipoLetraLabels.VB_Description = "Tipo de letra y demas opciones de los nombres de servidor"
Attribute TipoLetraLabels.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set TipoLetraLabels = Label(0).Font
End Property

Public Property Let TipoLetraLabels(Font As StdFont)
    Dim i As Integer
    For i = 0 To LabelDireccion.Count - 1
        Set Label(i).Font = Font
        Set LabelSombra(i).Font = Font
    Next i
    PropertyChanged "TipoLetraLabels"
End Property

Public Property Get TipoLetraDireccion() As StdFont
Attribute TipoLetraDireccion.VB_Description = "Tipo de letra y demas opciones de la direccion IP y el puerto"
Attribute TipoLetraDireccion.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set TipoLetraDireccion = LabelDireccion(0).Font
End Property

Public Property Let TipoLetraDireccion(Font As StdFont)
    Dim i As Integer
    For i = 0 To LabelDireccion.Count - 1
        Set LabelDireccion(i).Font = Font
    Next i
    PropertyChanged "TipoLetraDireccion"
End Property

Public Property Get PunteroItems() As MousePointerConstants
Attribute PunteroItems.VB_Description = "Permite seleccionar el puntero que se mostrar� al pararse encima de los items"
Attribute PunteroItems.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PunteroItems = ImagenClick(0).MousePointer
End Property

Public Property Let PunteroItems(puntero As MousePointerConstants)
    Dim i As Integer
    For i = 0 To ImagenClick.Count - 1
        ImagenClick(i).MousePointer = puntero
    Next i
    PropertyChanged "PunteroItems"
End Property

Public Property Get PunteroImagenItems() As StdPicture
Attribute PunteroImagenItems.VB_Description = "Permite utilizar un puntero personalizado para los items"
Attribute PunteroImagenItems.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set PunteroImagenItems = ImagenClick(0).MouseIcon
End Property

Public Property Let PunteroImagenItems(iconopuntero As StdPicture)
    Dim i As Integer
    For i = 0 To ImagenClick.Count - 1
        Set ImagenClick(i).MouseIcon = iconopuntero
    Next i
    PropertyChanged "PunteroImagenItems"
End Property
