VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmVivMantenimiento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Modulo de Administración Garatías Hipotecarias"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVivMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   14040
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4092
      Left            =   3360
      TabIndex        =   4
      Top             =   1080
      Width           =   3372
      _Version        =   1310723
      _ExtentX        =   5948
      _ExtentY        =   7218
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   120
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6852
            Key             =   "Parametros"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6952
            Key             =   "Garantias"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6A56
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6B76
            Key             =   "ProfesionalesMain"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6D87
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6EA5
            Key             =   "Desembolsos"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":6FC9
            Key             =   "OpTramite"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":70EF
            Key             =   "OpEjecutadas"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":720D
            Key             =   "Tramites"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":732C
            Key             =   "ProPersona"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7433
            Key             =   "ProEmpresa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7527
            Key             =   "ZonasDetalle"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7657
            Key             =   "Detalle"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7769
            Key             =   "Carpeta"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7885
            Key             =   "Zonas"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7C98
            Key             =   "Abogado"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivMantenimiento.frx":7DAA
            Key             =   "Ingeniero"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   6840
      ScaleHeight     =   5012.637
      ScaleMode       =   0  'User
      ScaleWidth      =   130
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.TreeView trvArbol 
      Height          =   4620
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   3288
      _ExtentX        =   5794
      _ExtentY        =   8149
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Zonas"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin MSComctlLib.Toolbar TlbAccion 
      Height          =   336
      Left            =   960
      TabIndex        =   7
      Top             =   -240
      Visible         =   0   'False
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   582
      ButtonWidth     =   1958
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Acción"
            Key             =   "Accion"
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   1
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Operaciones"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Desembolsos"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   4
      Left            =   7320
      TabIndex        =   10
      Top             =   120
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Suspensión"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   5
      Left            =   9000
      TabIndex        =   11
      Top             =   120
      Width           =   1092
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Asignación"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   6
      Left            =   10080
      TabIndex        =   12
      Top             =   120
      Width           =   1092
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Bancos"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   7
      Left            =   11160
      TabIndex        =   13
      Top             =   120
      Width           =   1092
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Garantía"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton Menu_Barra 
      Height          =   372
      Index           =   8
      Left            =   12480
      TabIndex        =   14
      Top             =   120
      Width           =   1092
      _Version        =   1310723
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informes"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton tbMain 
      Height          =   372
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   6
      EnableMarkup    =   -1  'True
      Picture         =   "frmVivMantenimiento.frx":7ED1
   End
   Begin XtremeSuiteControls.PushButton tbMain 
      Height          =   372
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   6
      EnableMarkup    =   -1  'True
      Picture         =   "frmVivMantenimiento.frx":8503
   End
   Begin XtremeSuiteControls.PushButton tbMain 
      Height          =   372
      Index           =   2
      Left            =   1920
      TabIndex        =   17
      Top             =   120
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   6
      EnableMarkup    =   -1  'True
      Picture         =   "frmVivMantenimiento.frx":8AFE
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   624
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12732
      _Version        =   1310723
      _ExtentX        =   22458
      _ExtentY        =   1101
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitulo 
      Height          =   336
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3372
      _Version        =   1310723
      _ExtentX        =   5948
      _ExtentY        =   582
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTituloListView 
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   3372
      _Version        =   1310723
      _ExtentX        =   5948
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "   Mantenimiento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Image imgSplitter 
      Appearance      =   0  'Flat
      Height          =   5865
      Left            =   3255
      MouseIcon       =   "frmVivMantenimiento.frx":9205
      MousePointer    =   99  'Custom
      Top             =   810
      Width           =   120
   End
End
Attribute VB_Name = "frmVivMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Item_Lista_Seleccionado As XtremeSuiteControls.ListViewItem

Public NodoSeleccionado As MSComctlLib.Node
Public Tnodo As MSComctlLib.Node

Private mbMoving As Boolean
Private vKey As String
Const sglSplitLimit = 500

Private vNodoText As String
Private vNodoTag As String
Private vNodoKey As String
Private vNodoIndex As Integer


Private Sub inicializaStatus()
'    frmMenu.StatusBar.Panels(2).Text = "A: " & Format(0, "###,##0")
'    frmMenu.StatusBar.Panels(3).Text = "B: " & Format(0, "###,##0")
'    frmMenu.StatusBar.Panels(4).Text = "C: " & Format(0, "###,##0")
'    frmMenu.StatusBar.Panels(5).Text = "D: " & Format(0, "###,##0")
'    frmMenu.StatusBar.Panels(6).Text = "F: " & Format(0, "###,##0")
'    frmMenu.StatusBar.Panels(7).Text = "E: " & Format(0, "###,##0")
End Sub


Sub SizeControls(x As Single)
  On Error Resume Next
  
  If x < 1500 Then x = 1500
  If x > (Me.Width - 1500) Then x = Me.Width - 1500
  trvArbol.Width = x + 30 '60 disminuir 35
  
  imgSplitter.Left = x
  lsw.Left = x + 75
  'lsw.Width = Me.Width - (trvArbol.Width + 140)
  ShortcutCaptionTitle.Width = Me.Width
  
  lblTituloListView.Width = Me.Width - (trvArbol.Width + 160)
  lsw.Width = lblTituloListView.Width
  
  lblTitulo.Top = ShortcutCaptionTitle.Top + ShortcutCaptionTitle.Height + 10
  lblTituloListView.Top = lblTitulo.Top
  
  trvArbol.Top = lblTitulo.Top + lblTitulo.Height + 10

  trvArbol.Height = Me.Height - (trvArbol.Top + 650)
  
  
  lsw.Top = trvArbol.Top
  lsw.Height = trvArbol.Height
  
  imgSplitter.Top = trvArbol.Top
  imgSplitter.Height = trvArbol.Height
  lblTituloListView.Left = x + 80
End Sub

Private Sub Selecciona_NodoArbol()
On Error GoTo error

Dim PkeyNodoPadre As String
Dim NodoClick As MSComctlLib.Node

If Len(NodoSeleccionado.Key) > 0 Then
    Set NodoClick = Buscar_ArbolKey(trvArbol, NodoSeleccionado.Key) 'Selecciona el servicio
    If Not NodoClick Is Nothing Then
        NodoClick.Parent.Expanded = True
        trvArbol.Nodes(NodoClick.Index).Selected = True
        NodoClick.Selected = True
        NodoClick.Expanded = True
        Call trvArbol_NodeClick(NodoClick)
    Else
        Call trvArbol_NodeClick(NodoSeleccionado)
        NodoSeleccionado.Selected = True
        NodoSeleccionado.Expanded = True
    End If
Else
    Call trvArbol_NodeClick(NodoSeleccionado)
    NodoSeleccionado.Selected = True
    NodoSeleccionado.Expanded = True
End If

salir:
Exit Sub
error:
If Err.Number = 91 Then Resume salir
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
Private Function Buscar_ArbolKey(pArbol As TreeView, pKey As String) As MSComctlLib.Node
On Error GoTo error
Dim i As Long
For i = 1 To pArbol.Nodes.Count
  If pArbol.Nodes(i).Key = pKey Then
    Set Buscar_ArbolKey = pArbol.Nodes(i)
    Exit For
  End If
Next i
Exit Function
error:
MsgBox "Ocurrió un error en la busqueda de nodos en el árbol de médicos.", vbCritical
End Function
Private Sub BuscarNodoenArbol(ByVal codigo As String)
    Dim j As Long
    For j = 1 To trvArbol.Nodes.Count
        If trvArbol.Nodes(j).Key = codigo Then
           trvArbol.Nodes(j).Expanded = True
           trvArbol.Nodes(j).Selected = True
           Call trvArbol_NodeClick(trvArbol.Nodes(j))
           Exit For
        End If
    Next j
End Sub
Private Sub sbCreaNodoRaiz(ByRef pNodo As MSComctlLib.Node, ByRef pTree As Object)
On Error GoTo error

    pTree.Nodes.Clear
    
    Set pNodo = pTree.Nodes.Add()
        pNodo.Text = "Crédito Hipotecario"
        pNodo.Tag = "ModuloDeVivienda"
        pNodo.Image = "Root"
        pNodo.Key = "Vv-" & Str(pNodo.Index)
        

salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Resume salir
End Sub

Private Sub CreaNodoParametros(ByRef pNodo As MSComctlLib.Node, ByRef pTree As Object)
On Error GoTo error
    Dim wIndice As Integer

    Set pNodo = pTree.Nodes.Add(2, tvwChild)
        pNodo.Text = "Mantenimiento de Parámetros Generales"
        pNodo.Tag = "NodoTituloMantenimientoParametros"
        pNodo.Image = 3
        wIndice = pNodo.Index
        
    

salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, ""
    Resume salir
End Sub



Public Sub sbCargaArbol(ByRef pNodo As MSComctlLib.Node, ByRef pTree As Object)
On Error GoTo error
 
    Call sbCreaNodoRaiz(pNodo, pTree)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Parámetros", "NodoParametrosGenerales", "Parametros", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Zonas", "NodoZonas", "Zonas", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Profesionales", "NodoAsignacionProfesionales", "ProfesionalesMain", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, pNodo.Index, "Empresas", "NodoEmpresas", "ProEmpresa", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, pNodo.Index - 2, "Personas Físicas", "NodoPersonasFísicas", "ProPersona", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, pNodo.Index, "Ingenieros", "NodoAsigIngPF", "Ingeniero", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, pNodo.Index - 2, "Abogados", "NodoAsigAbogPF", "Abogado", "")
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Conceptos Desembolsos", "NodoTiposDesembolsos", "Desembolsos", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Garantías en Tramite", "NodoTramiteGarantia", "Garantias", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Operaciones en tramite", "NodoOperacionesTramite", "Tramites", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Tiempos de Seguimiento", "NodoTiemposSeguimiento", "Tramites", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Control Desembolso", "NodoControlDesembolso", "Tramites", "", False)
    Call ObjMantenimiento.sbCreaNodosHijos(pNodo, pTree, 1, "Operaciones Canceladas", "NodoOperacionesCanceladas", "Desembolsos", "", False)
    
    

    
    
    
    pTree.Nodes(1).Expanded = True

salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, "Vivienda"
    Resume salir
End Sub



Private Sub Inicializa_Ventana()
On Error GoTo Errores

Call sbCargaArbol(Tnodo, trvArbol)
salir:
    Exit Sub
Errores:
    
    Resume salir
End Sub

Private Sub sbEditar()
On Error GoTo error

Dim vKey As String, frm As Form

    Select Case vNodoTag
    
    Case "AsignacionAbogados"
    
    Case "NodoParametrosGenerales"
        Call sbFormsCall("frmVivParametros", vbModal, , , False, Me)
        
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
        
    Case "NodoZonasHijo"
        If Not Item_Lista_Seleccionado Is Nothing Then
            
            Call sbFormsCall("frmVivZonas", 1, , , False)
           
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        End If
    Case "NodoAbogadosEmpresaHijo", "NodoIngenierosEmpresaHijo", "NodoEmpresaHijo"
        'Edita información de empresas, abogados e ingenieros
         If Not Item_Lista_Seleccionado Is Nothing Then 'Selecciono un valor de la lista
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Ie)")
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Id)")
        Else
            vKey = fxDeCodePK(vNodoKey, 5, "(Ic)")
            vKey = fxDeCodePK(vNodoKey, gPosIni, "(Ie)")
        End If
        
        Call sbFormsCall("frmVivProfesionales", 1, , , False)
        Call sbFormActivo("frmVivProfesionales", frm)
        Call frm.sbConsulta_Externa_IdPersona(vKey)
        
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
        
    Case "NodoAbogadosPFHijo", "NodoIngenierosPFHijo"
    
        'Edita información de Personas fisicas, abogados e ingenieros
        
        vKey = fxDeCodePK(vNodoKey, 5, "(Ic)")
        vKey = fxDeCodePK(vNodoKey, gPosIni, "(Ie)")
        vKey = fxDeCodePK(vNodoKey, gPosIni, "(Id)")
        
        Call sbFormsCall("frmVivProfesionales", 1, , , False)
        Call sbFormActivo("frmVivProfesionales", frm)
        Call frm.sbConsulta_Externa_IdPersona(vKey)
        
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
        
    Case "NodoIngenierosZonaHijo", "NodoAbogadosZonaHijo"
    'Edita información de Profesionales por zona, abogados e ingenieros
        
        vKey = fxDeCodePK(vNodoKey, 5, "(Iz)")
        vKey = fxDeCodePK(vNodoKey, gPosIni, "(Ie)")
        
        Call sbFormsCall("frmVivProfesionales", 1, , , False)
        Call sbFormActivo("frmVivProfesionales", frm)
        Call frm.sbConsulta_Externa_IdPersona(vKey)
        
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
        
    Case "NodoAbogZanasHijo", "NodoIngZanasHijo"
    
        If Not Item_Lista_Seleccionado Is Nothing Then
        
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Iz)")
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Ie)")
            
            
            Call sbFormsCall("frmVivProfesionales", 1, , , False)
            Call sbFormActivo("frmVivProfesionales", frm)
            Call frm.sbConsulta_Externa_IdPersona(vKey)
            
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        End If
        
    Case "NodoEmpresas", "NodoEmpresaHijo", "NodoAsigIngEmpresa", _
        "NodoAsigAbogEmpresa", "NodoPersonasFísicas", "NodoAsigIngPF", "NodoAsigAbogPF"
        
        If Not Item_Lista_Seleccionado Is Nothing Then
        
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Dc)")
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Ie)")
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, gPosIni, "(Id)")
            
            Call sbFormsCall("frmVivProfesionales", 1, , , False)
            Call sbFormActivo("frmVivProfesionales", frm)
            Call frm.sbConsulta_Externa_IdPersona(vKey)
            
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        End If
        
    Case "NodoTiposDesembolsos"
        Call sbFormsCall("frmVivTiposDesembolsos", 1, , , False)
        
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
        
    'Case "NodoProfesionalesxZona"
    
'------Tramite de Operaciones -------------------------------------
    Case "NodoTramiteGarantia"
        If Not Item_Lista_Seleccionado Is Nothing Then
        
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
            gOperacion = vKey
            Call sbFormsCall("frmVivGarantia", , , , False)
            
            DoEvents
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        End If
'------Operaciones en tramite-------------------------------------
    Case "NodoOperacionesTramite"
    
        If Not Item_Lista_Seleccionado Is Nothing Then
        
            vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
            gOperacion = vKey
            Call sbFormsCall("frmVivGarantia", , , , False)
            DoEvents
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
        End If
        
'----Nodo Tramite de operaciones-----------------------------------
    Case "NodoTramiteGarantia"
    
'----Nodo Operaciones en tramite-----------------------------------
    Case "NodoOperacionesTramite"
          
'----Operaciones por Profesional en Zonas-----------------------------------

    Case "NodoOperaAbogZonaTram", "NodoOperaAbogZonaEje"
    
    Case "NodoOperaIngZonaTram", "NodoOperaIngEmpresaTram" '"NodoOperaIngZonaEje"
        vKey = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
        gOperacion = vKey
            Call sbFormsCall("frmVivGarantia", , , , False)
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing


'----Operaciones por Profesional en Empresas-----------------------------------
    Case "NodoOperaAbogEmpresaTram", "NodoOperaAbogEmpresaEje"
    
    Case "NodoOperaIngEmpresaEje"
        
'----Operaciones por Profesional en Personas Fisicas-----------------------------------

    Case "NodoOperaAbogPFTram", "NodoOperaAbogPFEje"

    Case "NodoOperaIngPFTram", "NodoOperaIngPFEje"
    
'---Configuración de Tiempos de seguimiento para las garantias
    Case "NodoTiemposSeguimiento"
        Call sbFormsCall("frmVivTiemposSeguimiento", 1, , , False)
        DoEvents
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        Set Item_Lista_Seleccionado = Nothing
    End Select

  Exit Sub
error:
  If Err.Number <> 91 Then
    MsgBox fxSys_Error_Handler(Err.Description)
  End If
End Sub

Private Sub Form_Activate()
gIconoLista = "Detalle" 'Carga el icono a mostrar el la lista detalle
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF5) Then
        Call RefrescarTodo
    End If
End Sub

Private Sub Form_Load()
On Error GoTo vError

'' Carga nombre de la ternimal
If Len(glogon.Maquina) = 0 Then
    Call sbMaquina
End If

vModulo = 3 'Modulo de Credito

'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

'Call ObjMantenimiento.sbToolBarIconos(TlbBarraHerramientas, True)
Call sbCargaArbol(Tnodo, trvArbol)

salir:
    Exit Sub
vError:
    If Err.Number = 13 Then Resume Next
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Resume salir
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .Top, .Width - 20, .Height - 20
    'MsgBox lblTituloListView.Left
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = x + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
  
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
  
End Sub



Private Sub Form_Resize()
On Error GoTo Errores
  If Me.WindowState <> 1 Then
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
  End If
salir:
    Exit Sub
Errores:
    If Err.Number = 384 Then
        Resume Next
    Else
      
    End If
End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vIdcontacto As Long
Dim vIdZona As Long

On Error GoTo vError


If (NodoSeleccionado.Tag = "NodoIngenierosEmpresaHijo") Or (NodoSeleccionado.Tag = "NodoAbogadosEmpresaHijo") _
    Or (NodoSeleccionado.Tag = "NodoAbogadosPFHijo") Or (NodoSeleccionado.Tag = "NodoIngenierosPFHijo") Then
    If lsw.ListItems.Count > 0 Then
        If (Item.Checked) And (Item.ListSubItems(3).Text = "N") Then
            If (NodoSeleccionado.Tag = "NodoAbogadosPFHijo") Or (NodoSeleccionado.Tag = "NodoIngenierosPFHijo") Then
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, 5, "(Ic)")
            Else
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, 5, "(Em)")
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, gPosIni, "(ic)")
            End If
            
            vIdZona = fxDeCodePK(Item.Key, 5, "(id)")
            If ObjAgregar.fxViviendaContatosxZonas(vIdZona, vIdcontacto, glogon.Usuario, "1900/01/01") Then
                Item.ListSubItems(3).Text = "M"
                Item.ForeColor = vbBlue
                Item.ListSubItems(1).ForeColor = vbBlue
            End If
            
        ElseIf (Item.Checked = False) And (Item.ListSubItems(3).Text = "M") Then
            If (NodoSeleccionado.Tag = "NodoAbogadosPFHijo") Or (NodoSeleccionado.Tag = "NodoIngenierosPFHijo") Then
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, 5, "(Ic)")
            Else
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, 5, "(Em)")
                vIdcontacto = fxDeCodePK(NodoSeleccionado.Key, gPosIni, "(ic)")
            End If
            
            vIdZona = fxDeCodePK(Item.Key, 5, "(id)")
            If ObjBorrar.fxViviendaContactosxZona(vIdZona, vIdcontacto) Then
                Item.ListSubItems(3).Text = "N"
                Item.ForeColor = vbBlack
                Item.ListSubItems(1).ForeColor = vbBlack
            End If
        End If
    End If
End If

    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    
End Sub


Private Sub RefrescarTodo()
    'Set NodoSeleccionadoTemp = NodoSeleccionado
    Call sbCargaArbol(Tnodo, trvArbol)
    'Set NodoSeleccionado = NodoSeleccionadoTemp
    'Call Selecciona_NodoArbol
End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Set Item_Lista_Seleccionado = Item
End Sub



Private Sub tbMain_Click(Index As Integer)

On Error GoTo error
    Select Case Index
            
        Case 0 'nuevo"
            Call Crear
            
        Case 1 'editar"
            Call sbEditar
            
        Case 2 'Imprimir"
            Call sbImprimir
    
    End Select
    
salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    Me.MousePointer = vbDefault
    Resume salir
End Sub

Private Sub TlbAccion_ButtonClick(ByVal Button As MSComctlLib.Button)
If TlbAccion.Buttons.Item(1).Caption = "&Acción" Then
    TlbAccion.Buttons.Item(1).Value = tbrUnpressed
Else
    If Button.Value = tbrUnpressed Then
        TlbAccion.Buttons.Item(1).Caption = "&Acción"
        TlbAccion.Buttons.Item(1).Image = 0
    End If
End If
   Call Selecciona_NodoArbol
    
End Sub

Private Sub Menu_Barra_Click(Index As Integer)

On Error GoTo vError

    TlbAccion.Buttons.Item(1).Caption = "&Acción"
    TlbAccion.Buttons.Item(1).Value = tbrUnpressed
    
Select Case Index
    Case 0 'Zonas
        TlbAccion.Buttons.Item(1).Caption = "&Zonas"
        TlbAccion.Buttons.Item(1).Value = tbrPressed

        Call Selecciona_NodoArbol
    Case 1 'Operaciones"
        TlbAccion.Buttons.Item(1).Caption = "&Operaciones"
        TlbAccion.Buttons.Item(1).Value = tbrPressed

    Case 2 'Desembolsos"
        TlbAccion.Buttons.Item(1).Caption = "&Desembolsos"
        TlbAccion.Buttons.Item(1).Value = tbrPressed

        If Item_Lista_Seleccionado Is Nothing Then Exit Sub
        gOperacion = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
        
        Call sbFormsCall("frmVivDesembolsos", 1, , , False)
    
    Case 2 'Desembolsos
        TlbAccion.Buttons.Item(1).Caption = "&Garantias"
        TlbAccion.Buttons.Item(1).Value = tbrPressed

        If Not Item_Lista_Seleccionado Is Nothing Then
            gOperacion = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
        Else
            gOperacion = 0
        End If
 
        Call sbFormsCall("frmVivDesembolsos ", 1, , , False)
        Call trvArbol_NodeClick(trvArbol.SelectedItem)

    Case 4 'Suspensión
        Call sbFormsCall("frmVivEstadoProfesionales", , , , False)
            
   Case 5 'Asignación de la Garantia a Profesionales
        Call sbFormsCall("frmVivControlAsignacionGarantia", , , , False)
   
   Case 6 'Bancos
        Call sbFormsCall("frmVivRemesasTesoreria", , , , False)
    
   Case 7 'Garantias
        If Not Item_Lista_Seleccionado Is Nothing Then
            gOperacion = fxDeCodePK(Item_Lista_Seleccionado.Key, 5, "(Op)")
        Else
            gOperacion = 0
        End If
        Call sbFormsCall("frmVivGarantia", 1, , , False)
        Call trvArbol_NodeClick(trvArbol.SelectedItem)
        
   Case 8 'Informes
       Call sbFormsCall("frmVivReportesGarantias", 1, , , False)
End Select

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbImprimir()
  Me.MousePointer = vbHourglass
  Dim vTemp As String
  
Select Case vNodoTag

    
    Case "NodoIngenierosPFHijo"
        Call ObjMantenimiento.sbImprimirZonasXContacto(fxDeCodePK(vNodoKey, 5, "(Ic)"), "I", False)
        
    Case "NodoAbogadosPFHijo"
        Call ObjMantenimiento.sbImprimirZonasXContacto(fxDeCodePK(vNodoKey, 5, "(Ic)"), "A", False)
        
    Case "NodoAsigIngPF"
        Call ObjMantenimiento.sbImprimirZonasXContactoPF(-1, "I")
    
    Case "NodoAsigAbogPF"
        Call ObjMantenimiento.sbImprimirZonasXContactoPF(-1, "A")
        
    Case "NodoEmpresas" 'Reporte de todas Empresas
        Call ObjMantenimiento.sbImprimirContactoxEmpresa(1, True)

    Case "NodoEmpresaHijo" 'Reporte de Contactos por empresa
        vTemp = fxDeCodePK(vNodoKey, 5, "(Em)")
        Call ObjMantenimiento.sbImprimirContactoxEmpresa(vTemp, False)

    Case "NodoIngenierosEmpresaHijo"
        vTemp = fxDeCodePK(vNodoKey, 5, "(Em)")
        Call ObjMantenimiento.sbImprimirZonasXContacto(fxDeCodePK(vNodoKey, gPosIni, "(Ic)"), "I", False)
        
    Case "NodoAbogadosEmpresaHijo"
        vTemp = fxDeCodePK(vNodoKey, 5, "(Em)")
        Call ObjMantenimiento.sbImprimirZonasXContacto(fxDeCodePK(vNodoKey, gPosIni, "(Ic)"), "A", False)
    
    Case "NodoAsigAbogEmpresa"
        vTemp = fxDeCodePK(NodoSeleccionado.Parent.Key, 5, "(Em)")
        Call ObjMantenimiento.sbImprimirZonasXContacto(vTemp, "A", True)
        
    Case "NodoAsigIngEmpresa"
        vTemp = fxDeCodePK(NodoSeleccionado.Parent.Key, 5, "(Em)")
        Call ObjMantenimiento.sbImprimirZonasXContacto(vTemp, "", True)
    
End Select

Me.MousePointer = vbDefault
End Sub
Private Sub Crear()
  'Dim wPrimaryKey As Variant
  
    Select Case vNodoTag
    
        Case "NodoAsignacionProfesionales", "NodoAsigAbogEmpresa", "NodoAsigIngEmpresa", "NodoIngenierosEmpresaHijo", _
        "NodoEmpresaHijo", "NodoAsigAbogPF", "NodoAsigIngPF", "NodoPersonasFísicas", "NodoEmpresas"

            Call sbFormsCall("frmVivProfesionales", 1, , , False)
            
            DoEvents
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        Case "NodoZonas", "NodoZonasHijo", "NodoIngenierosPFHijo", "NodoAbogadosPFHijo", _
              "NodoAbogadosEmpresaHijo"
            Call sbFormsCall("frmVivZonas", 1, , , False)
            
            DoEvents
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        Case "NodoTiposDesembolsos"
            
            Call sbFormsCall("frmVivTiposDesembolsos", 1, , , False)
            
            DoEvents
            Call trvArbol_NodeClick(trvArbol.SelectedItem)
            Set Item_Lista_Seleccionado = Nothing
            
        Case "NodoTramiteGarantia", "NodoOperacionesTramite"
            Call sbFormsCall("frmVivControlAsignacionGarantia", , , , False)
    End Select
End Sub

Private Sub trvArbol_Expand(ByVal Node As MSComctlLib.Node)

On Error GoTo vError

Me.MousePointer = vbHourglass

Set ObjMantenimiento.TreeView = Me.trvArbol
Set ObjMantenimiento.ListView = Me.lsw

Call ObjMantenimiento.Expand(Node)
        
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbArbolClick(ByVal Node As MSComctlLib.Node)

On Error GoTo error

Me.MousePointer = vbHourglass


Call sbHabilitarBotones(Node)

lsw.ColumnHeaders.Clear
lsw.ListItems.Clear
lsw.Checkboxes = False

Set ObjMantenimiento.TreeView = Me.trvArbol
Set ObjMantenimiento.ListView = Me.lsw
    
    
Select Case Node.Tag


    Case "NodoParametrosGenerales"
        Call ObjMantenimiento.sbListaParametros
        
    Case "AsignacionAbogados"
    
    Case "NodoIngenierosZonaHijo"
    
    Case "NodoAbogadosZonaHijo"
    
    Case "NodoZonas"
        Call ObjMantenimiento.sbListaZanas
        
    Case "NodoIngZanasHijo"
        Call ObjMantenimiento.sbListaProfesionalesxZona(Node.Parent.Index, "I")
        
    Case "NodoAbogZanasHijo"
        Call ObjMantenimiento.sbListaProfesionalesxZona(Node.Parent.Index, "A")

    Case "NodoZonasHijo"
        Call ObjMantenimiento.sbListaUbicaciondeZanas(Node.Index)
        
    Case "NodoCantonHijo"
        Call ObjMantenimiento.sbListaZanasXCanton(Node.Index)
        
    Case "NodoEmpresaHijo"
        Call ObjMantenimiento.sbListaContactosEmpresa(Node.Index)
        
    Case "NodoEmpresas"
        Call ObjMantenimiento.sbListaEmpresas
        
    Case "NodoAsigIngEmpresa"
        Call ObjMantenimiento.sbListaContactosxTipoProfesional(Node.Parent.Index, "I")
        
    Case "NodoAsigAbogEmpresa"
        Call ObjMantenimiento.sbListaContactosxTipoProfesional(Node.Parent.Index, "A")
    
    Case "NodoPersonasFísicas"
        Call ObjMantenimiento.sbListaPersonasFisicasxTipoProfesional("")
        
    Case "NodoAsigIngPF"
        Call ObjMantenimiento.sbListaPersonasFisicasxTipoProfesional("I")
        
    Case "NodoAsigAbogPF"
        Call ObjMantenimiento.sbListaPersonasFisicasxTipoProfesional("A")
        
    Case "NodoTiposDesembolsos"
        Call ObjMantenimiento.sbListaTiposDesembolsos
         
    Case "NodoProfesionalesxZona"
    
    Case "NodoIngenierosEmpresaHijo"
        
        If TlbAccion.Buttons(1).Value = tbrPressed Then
            Me.lsw.Checkboxes = True
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas(Node.Index)
        Else
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas
        End If
        
     
    Case "NodoAbogadosEmpresaHijo"
        If TlbAccion.Buttons(1).Value = tbrPressed Then
            Me.lsw.Checkboxes = True
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas(Node.Index)
        Else
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas
        End If
        
    Case "NodoAbogadosEmpresaHijo"
        
'------------------Consulta de Operaciones Personas fisicas-----------------------------

    
    Case "NodoIngenierosPFHijo", "NodoAbogadosPFHijo"
       
        If TlbAccion.Buttons(1).Value = tbrPressed Then
            Me.lsw.Checkboxes = True
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas(Node.Index, False)
        Else
            Set ObjMantenimiento.ListView = Me.lsw
            Call ObjMantenimiento.sbListaZanas
        End If
    
'----Nodo Tramite de operaciones-----------------------------------
    Case "NodoTramiteGarantia"
        ObjMantenimiento.sbListaTramiteDeOperaciones
    
'----Nodo Operaciones en tramite-----------------------------------
    Case "NodoOperacionesTramite"
        ObjMantenimiento.sbListaOperacionesEnTramite
          
'----Operaciones por Profesional en Zonas-----------------------------------

    Case "NodoOperaAbogZonaTram", "NodoOperaAbogZonaEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Iz)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ie)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ic)")
        
        If Node.Tag = "NodoOperaAbogZonaTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "T")
            
        ElseIf Node.Tag = "NodoOperaIngZonaEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "E")
        End If
        
    Case "NodoOperaIngZonaTram", "NodoOperaIngZonaEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Iz)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ie)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ic)")
        
        If Node.Tag = "NodoOperaIngZonaTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "T")
            
        ElseIf Node.Tag = "NodoOperaIngZonaEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "E")
        End If
    
'----Operaciones por Profesional en Empresas-----------------------------------
    Case "NodoOperaAbogEmpresaTram", "NodoOperaAbogEmpresaEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Em)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ic)")
        
        If Node.Tag = "NodoOperaAbogEmpresaTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "T")
            
        ElseIf Node.Tag = "NodoOperaAbogEmpresaEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "E")
        End If
    
    Case "NodoOperaIngEmpresaTram", "NodoOperaIngEmpresaEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Em)")
        vKey = fxDeCodePK(Node.Parent.Key, gPosIni, "(Ic)")
        
        If Node.Tag = "NodoOperaIngEmpresaTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "T")
            
        ElseIf Node.Tag = "NodoOperaIngEmpresaEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "E")
        End If
    
        
'----Operaciones por Profesional en Personas Fisicas-----------------------------------

    Case "NodoOperaAbogPFTram", "NodoOperaAbogPFEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Ic)")
        If Node.Tag = "NodoOperaAbogPFTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "T")
            
        ElseIf Node.Tag = "NodoOperaAbogPFEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "A", "E")
        End If

    Case "NodoOperaIngPFTram", "NodoOperaIngPFEje"
        vKey = fxDeCodePK(Node.Parent.Key, 5, "(Ic)")
        
        If Node.Tag = "NodoOperaIngPFTram" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "T")
            
        ElseIf Node.Tag = "NodoOperaIngPFEje" Then
            Call ObjMantenimiento.sbListaOperacionesXProsional(vKey, "I", "E")
        End If
    
    Case "NodoTiemposSeguimiento"
        Call ObjMantenimiento.sbListaTiemposSgte
    
    Case "NodoControlDesembolso"
        Call ObjMantenimiento.sbListaControlDesembolso(False)
        
    Case "NodoOperacionesCanceladas"
        
        Call ObjMantenimiento.sbListaControlDesembolso(True)
        
End Select
 
Me.MousePointer = vbDefault
 
salir:
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

    
Private Sub sbHabilitarBotones(ByVal Node As MSComctlLib.Node)

'    Select Case Node.Tag
'
'            '-------------------------Paramentros ----------------------
'        Case "NodoParametrosGenerales", "ModuloDeVivienda"
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False  'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            tlbPrincipal.Buttons.Item(1).Enabled = False 'Nuevo
'        Case "AsignacionAbogados"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'
'            '--------------------------Zonas---------------------
'
'        Case "NodoZonas"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'
'        Case "NodoAbogadosZonaHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
''        Case "NodoOperaIngZona"
''
'        Case "NodoIngenierosZonaHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
'        Case "NodoEmpresaHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
''            '--------------------------Profesionales----------------------
''
''        Case "NodoProfesionalesxZona"
'        Case "NodoOperaIngPFTram", "NodoOperaIngPFEje"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = True 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'        Case "NodoOperaAbogPFTram", "NodoOperaAbogPFEje"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = True 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
'            '--------------------------Profesionales.Empresas----------------------
'        Case "NodoIngenierosEmpresaHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
'        Case "NodoAbogadosEmpresaHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
'
'        Case "NodoOperaAbogEmpresaTram"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = True 'Desembolosos
'
'        Case "NodoOperaIngEmpresaTram"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'
'        Case "NodoIngenierosPFHijo", "NodoAbogadosPFHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = True 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'            TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'            TlbBarraHerramientas.Buttons.Item(5).Enabled = True 'Suspensión Contacto
''            '--------------------------Profesionales.Personas Fisica---------------------
'        Case "NodoIngenierosPFHijo"
'            tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'            TlbBarraHerramientas.Buttons.Item(1).Enabled = True 'Zonas
'            TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'            TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'
'''----------------------------Tipos Desembolsos---------------------
''        Case "NodoTiposDesembolsos"
'
''------Tramite de Operaciones -------------------------------------
'    Case "NodoTramiteGarantia"
'        tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'        TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'        TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'        TlbBarraHerramientas.Buttons.Item(3).Enabled = True 'Desembolosos
'        TlbBarraHerramientas.Buttons.Item(4).Enabled = True 'Tramite de Garantia
'        TlbBarraHerramientas.Buttons.Item(5).Enabled = False 'Suspensión Contacto
'
''-----------------Operaciones en tramite---------------------------------------------
'    Case "NodoOperacionesTramite", "NodoOperacionesCanceladas", "NodoControlDesembolso"
'
'        tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'        TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'        TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'        TlbBarraHerramientas.Buttons.Item(3).Enabled = True 'Desembolosos
'        TlbBarraHerramientas.Buttons.Item(4).Enabled = True 'Tramite de Garantia
'        TlbBarraHerramientas.Buttons.Item(5).Enabled = False 'Suspensión Contacto
'
''---Configuración de Tiempos de seguimiento para las garantias
'    Case "NodoTiemposSeguimiento"
'        tlbPrincipal.Buttons.Item(1).Enabled = False 'Nuevo
'        TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'        TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'        TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'        TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'        TlbBarraHerramientas.Buttons.Item(5).Enabled = False 'Suspensión Contacto
'
'    Case Else
'        tlbPrincipal.Buttons.Item(1).Enabled = True 'Nuevo
'        TlbBarraHerramientas.Buttons.Item(1).Enabled = False 'Zonas
'        TlbBarraHerramientas.Buttons.Item(2).Enabled = False 'Operaciones
'        TlbBarraHerramientas.Buttons.Item(3).Enabled = False 'Desembolosos
'        TlbBarraHerramientas.Buttons.Item(4).Enabled = False 'Tramite de Garantia
'        TlbBarraHerramientas.Buttons.Item(5).Enabled = False 'Suspensión Contacto
'
'    End Select
'
''Inicializa Seguridad
'    Call Formularios(Me)
'    Call RefrescaTags(Me)
    
End Sub

'Public Sub EstiloImprimir()
'On Error GoTo Error
'    If TVCentroMedico.SelectedItem.Tag = "NodoGrupo" _
'       Or TVCentroMedico.SelectedItem.Tag = "NodoUsuario" Then
'        Toolbar.Buttons(3).Style = tbrDropdown
'        Toolbar.Buttons(3).ButtonMenus.Item(1).Visible = True
'        Toolbar.Buttons(3).ButtonMenus.Item(2).Visible = True
'        If TVCentroMedico.SelectedItem.Tag = "NodoUsuario" Then
'            Toolbar.Buttons(3).ButtonMenus.Item(3).Visible = True
'        End If
'    Else
'        Toolbar.Buttons(3).Style = tbrDefault
'        Toolbar.Buttons(3).ButtonMenus.Item(1).Visible = False
'        Toolbar.Buttons(3).ButtonMenus.Item(2).Visible = False
'        Toolbar.Buttons(3).ButtonMenus.Item(3).Visible = False
'    End If
'    Exit Sub
'salir:
'    Exit Sub
'Error:
'    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
'    Resume salir
'End Sub

Public Sub trvArbol_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo vError

If Node Is Nothing Then Exit Sub
    
    vNodoKey = Node.Key
    vNodoTag = Node.Tag
    vNodoText = Node.Text
    vNodoIndex = Node.Index
    
    If Node.Tag <> "ModuloDeVivienda" Then
        Set NodoSeleccionado = Node
    End If
    
  lblTituloListView.Caption = Node.FullPath
  lblTitulo.Caption = Node.Text
  
    Call sbArbolClick(Node)
        
Exit Sub
    
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



