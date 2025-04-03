VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmUS_DerechosXOpcion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permisos x Opción"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgExp01 
      Left            =   4800
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":0000
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6862
            Key             =   "imgFormularios"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":D0C4
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":13926
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":1A188
            Key             =   "x2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":209EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":2724C
            Key             =   "imgOpcionDetalle"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":2DAAE
            Key             =   "x1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":34310
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":3AB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":413D4
            Key             =   "imgFrmOpcion"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":47C36
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":4E498
            Key             =   "imgGrupoDetalle"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":54CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":5B55C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerInicia 
      Interval        =   10
      Left            =   4800
      Top             =   1080
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   688
      _CBWidth        =   10950
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "cboPermiso"
      MinHeight1      =   330
      Width1          =   2055
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   525
      NewRow2         =   0   'False
      Caption3        =   "Seleccione una opción"
      MinHeight3      =   330
      Width3          =   420
      Key3            =   "0"
      NewRow3         =   0   'False
      Begin VB.ComboBox cboPermiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipos de Permisos"
         Top             =   30
         Width           =   1860
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   2250
         TabIndex        =   1
         Top             =   30
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgSecurity"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Expander/Recoger"
               ImageIndex      =   14
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":61DBE
            Key             =   "imgFrm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":62698
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":62F72
            Key             =   "imgModulo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView vTree 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9975
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgExp01"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSecurity 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6384C
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":63B66
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":64440
            Key             =   "Cuestion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":64D1A
            Key             =   "CheckList"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65034
            Key             =   "User"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6534E
            Key             =   "UserGroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65668
            Key             =   "Keys"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65F42
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6681C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":670F6
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":679D0
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":67CEA
            Key             =   "SearchFolder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":685C4
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":68E9E
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   5175
      Left            =   5520
      TabIndex        =   5
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5655
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Usuarios"
      TabPicture(0)   =   "frmUS_DerechosXOpcion.frx":691B8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblX"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Grupos"
      TabPicture(1)   =   "frmUS_DerechosXOpcion.frx":6FA1A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Label lblX 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   30
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmUS_DerechosXOpcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node


Private Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = vTree.Nodes.Add(vPadre, tvwChild)
    nodX.Image = vImagen
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & vTree.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
Set vNode = nodX

End Sub


Private Sub sbCargaInicial()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xNode As Node, lng As Long


Me.MousePointer = vbHourglass

With vTree
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "US", "Root")
  xNode.Bold = True
  
  strSQL = "select * from modulos order by modulo"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("US", Trim(rs!Nombre), "imgModulo", True, "0x0" & rs!modulo & "M")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close

  strSQL = "select * from formularios order by formulario"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!modulo & "M", Trim(rs!Descripcion), "imgFrm", True, "0x0" & rs!modulo & "-" & rs!frmID & "F")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close
  
  strSQL = "select O.*,F.frmID" _
         & " from opciones O inner join formularios F on O.formulario = F.formulario"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!modulo & "-" & rs!frmID & "F", Trim(rs!Opcion_descripcion), "imgOpcion", False, "0x0" & rs!frmID & "-" & rs!id_opt & "O")
   
     .Nodes.Item(.Nodes.Count).Tag = 1
   
   rs.MoveNext
  Loop
  rs.Close


   xNode.Expanded = True

End With


Me.MousePointer = vbDefault

Me.Show

End Sub


Private Sub cboPermiso_Click()
   Call sbCargaInicial
End Sub


Private Sub Form_Load()
vModulo = 13

cboPermiso.AddItem "Autorizaciones"
cboPermiso.AddItem "Restricciones"
cboPermiso.Text = "Autorizaciones"
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub





Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, vTipo As String

If lblX.Tag = "" Or lblX.Tag = 0 Then Exit Sub

vTipo = Mid(cboPermiso.Text, 1, 1)

If ssTab.Tab = 0 Then
 'Usuarios
 If Item.Checked Then
    strSQL = "insert permisos(id_opt,nombre,tipo,estado) values(" _
           & lblX.Tag & ",'" & Item.Tag & "','U','" & vTipo & "')"
 Else
    strSQL = "delete permisos where id_opt = " & lblX.Tag _
           & " and estado = '" & vTipo & "' and nombre = '" & Item.Tag _
           & "' and tipo = 'U'"
 
 End If
 glogon.Conection.Execute strSQL
 
 Call sbSEGCuentaLog("16", cboPermiso.Text, glogon.Usuario, Item.Tag)
 
Else
 'Grupo
 If Item.Checked Then
    strSQL = "insert permisos(id_opt,nombre,tipo,estado) values(" _
           & lblX.Tag & ",'" & Item.Tag & "','G','" & vTipo & "')"
 Else
    strSQL = "delete permisos where id_opt = " & lblX.Tag _
           & " and estado = '" & vTipo & "' and nombre = '" & Item.Tag _
           & "' and tipo = 'G'"
 
 End If
 glogon.Conection.Execute strSQL
 
 Call sbSEGCuentaLog("14", cboPermiso.Text & "...:" & Item.Tag, glogon.Usuario)
 
End If


End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Call sbCargaDatosLsw
End Sub


Private Sub TimerInicia_Timer()
    TimerInicia.Interval = 0
    Call sbCargaInicial
End Sub



Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lng As Long

With vTree.Nodes
 For lng = 1 To .Count
  If Right(.Item(lng).Key, 1) = "M" Then
    .Item(lng).Expanded = IIf(.Item(lng).Expanded, False, True)
  End If
 Next lng
End With

End Sub

Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Function fxIndiceMultiple(xkey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xkey = fxIndiceCodigo(xkey)

blnPaso = True

If xkey = "" Then
  fxIndiceMultiple = ""
  Exit Function
End If

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xkey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xkey, i, 50) '50 es un default ningun asiento es tan largo

End If

fxIndiceMultiple = strResultado

End Function

Private Sub sbCargaDatosLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vTipo As String

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

vTipo = Mid(cboPermiso.Text, 1, 1)

If ssTab.Tab = 0 Then
 'Usuarios
 strSQL = "select nombre,descripcion,userID from usuarios" _
        & " where UserID in(select nombre from Permisos where tipo = 'U' and id_opt = " & lblX.Tag _
        & " and estado = '" & vTipo & "')"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Nombre)
       itmX.SubItems(1) = rs!Descripcion
       itmX.Tag = rs!UserID
       itmX.Checked = True
       If vTipo = "A" Then
           itmX.ForeColor = vbBlue
       Else
           itmX.ForeColor = vbRed
       End If
   rs.MoveNext
 Loop
 rs.Close

 strSQL = "select nombre,descripcion,userID from usuarios" _
        & " where userID not in(select nombre from Permisos where tipo = 'U' and id_opt = " & lblX.Tag _
        & " and estado = '" & vTipo & "') and Estado = 'A' order by nombre"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Nombre)
       itmX.SubItems(1) = rs!Descripcion
       itmX.Tag = rs!UserID
   rs.MoveNext
 Loop
 rs.Close


Else
 'Grupos
 strSQL = "select id_grupo,nombre from Grupos" _
        & " where id_grupo in(select nombre from Permisos where tipo = 'G' and id_opt = " & lblX.Tag _
        & " and estado = '" & vTipo & "')"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!id_grupo)
       itmX.SubItems(1) = rs!Nombre
       itmX.Tag = rs!id_grupo
       itmX.Checked = True
       If vTipo = "A" Then
           itmX.ForeColor = vbBlue
       Else
           itmX.ForeColor = vbRed
       End If
   rs.MoveNext
 Loop
 rs.Close
 
 strSQL = "select id_grupo,nombre from Grupos" _
        & " where id_grupo not in(select nombre from Permisos where tipo = 'G' and id_opt = " & lblX.Tag _
        & " and estado = '" & vTipo & "') order by nombre"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!id_grupo)
       itmX.SubItems(1) = rs!Nombre
       itmX.Tag = rs!id_grupo
   rs.MoveNext
 Loop
 rs.Close
 
 
 
End If


Me.MousePointer = vbDefault

End Sub

Private Sub vTree_NodeClick(ByVal Node As MSComctlLib.Node)
 
 If Node.Image = "imgOpcion" Then
    lblX.Tag = fxIndiceMultiple(Node.Key, "N")
    CoolBarX.Bands.Item(3).Caption = Node.FullPath
    lblX.Caption = cboPermiso.Text & " / " & Node.Text
    Call sbCargaDatosLsw
 Else
    lblX.Tag = 0
    CoolBarX.Bands.Item(3).Caption = ""
    lblX.Caption = ""
 End If

End Sub
