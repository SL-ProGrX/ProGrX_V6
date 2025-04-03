VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmUS_DerechosNew 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de permisos a los módulos"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgToolBarPermisos 
      Left            =   6720
      Top             =   1080
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
            Picture         =   "frmUS_DerechosNew.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":00F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":0230
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDenegado 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3960
      TabIndex        =   13
      Text            =   "Denegado"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAutorizado 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   3960
      TabIndex        =   12
      Text            =   "Autorizado"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   688
      BandCount       =   4
      _CBWidth        =   7635
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "cboTipo"
      MinHeight1      =   330
      Width1          =   2340
      NewRow1         =   0   'False
      Child2          =   "cboPermiso"
      MinHeight2      =   330
      Width2          =   2055
      NewRow2         =   0   'False
      Child3          =   "tlbAux"
      MinHeight3      =   330
      Width3          =   525
      NewRow3         =   0   'False
      Child4          =   "tlb"
      MinHeight4      =   330
      Width4          =   420
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   5175
         TabIndex        =   11
         Top             =   30
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   1984
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgToolBarPermisos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Cambios"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Deshacer"
               Key             =   "Deshacer"
               Object.ToolTipText     =   "DesHacer Cambios"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   4620
         TabIndex        =   10
         Top             =   30
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgToolBarPermisos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Expander/Recoger"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
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
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Tipos de Permisos"
         Top             =   30
         Width           =   1860
      End
      Begin VB.ComboBox cboTipo 
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
         TabIndex        =   8
         ToolTipText     =   "Grupo o Usuarios"
         Top             =   30
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdAplicar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Aplicar"
      Height          =   315
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   6
      Top             =   6390
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   2760
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
            Picture         =   "frmUS_DerechosNew.frx":0324
            Key             =   "imgFrm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":0BFE
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":14D8
            Key             =   "imgModulo"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSujeto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmUS_DerechosNew.frx":1DB2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   6255
   End
   Begin VB.Timer TimerInicia 
      Interval        =   10
      Left            =   6360
      Top             =   1800
   End
   Begin VB.CommandButton cmdDeshacer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Deshacer"
      Height          =   315
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.TreeView vTree 
      Height          =   5535
      Left            =   20
      TabIndex        =   0
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
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
      Left            =   6240
      Top             =   1800
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
            Picture         =   "frmUS_DerechosNew.frx":1DCD
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":20E7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":29C1
            Key             =   "Cuestion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":329B
            Key             =   "CheckList"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":35B5
            Key             =   "User"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":38CF
            Key             =   "UserGroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":3BE9
            Key             =   "Keys"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":44C3
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":4D9D
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":5677
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":5F51
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":626B
            Key             =   "SearchFolder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":6B45
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":741F
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmUS_DerechosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vScroll As Boolean


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

If Len(txtSujeto.Tag) = 0 Then
  vTree.Nodes.Clear
  Exit Sub
End If

Me.MousePointer = vbHourglass

With vTree
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "US", "Root")
  xNode.Bold = True
  
  strSQL = "select * from modulos order by modulo"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("US", rs!Nombre, "imgModulo", True, "0x0" & rs!modulo & "M")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close

  strSQL = "select * from formularios order by formulario"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!modulo & "M", rs!Descripcion, "imgFrm", True, "0x0" & rs!modulo & "-" & rs!frmID & "F")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close
  
  strSQL = "select DISTINCT O.*,F.frmID,P.nombre" _
         & " from opciones O inner join formularios F on O.formulario = F.formulario" _
         & " left join permisos P on O.id_opt = P.id_opt and P.nombre = '" & txtSujeto.Tag & "'" _
         & " and P.tipo = '" & gEntidad.Tipo & "' and P.estado = '" & Mid(cboPermiso.Text, 1, 1) & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!modulo & "-" & rs!frmID & "F", rs!Opcion_descripcion, "imgOpcion", False, "0x0" & rs!frmID & "-" & rs!id_opt & "O")
   
   If Not IsNull(rs!Nombre) Then
     .Nodes.Item(.Nodes.Count).ForeColor = IIf((Mid(cboPermiso.Text, 1, 1) = "A"), vbBlue, vbRed)
     .Nodes.Item(.Nodes.Count).Checked = True
     .Nodes.Item(.Nodes.Count).Bold = True
     .Nodes.Item(.Nodes.Count).Tag = 1
   Else
     .Nodes.Item(.Nodes.Count).Tag = 0
   End If
   
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

Private Sub cboTipo_Click()
If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
   gEntidad.Tipo = "G"
   txtSujeto.Text = ""
   txtSujeto.Tag = ""
   lblID.Caption = ""
   Call sbCargaInicial
Else
   gEntidad.Tipo = "U"
   txtSujeto.Text = ""
   txtSujeto.Tag = ""
   lblID.Caption = ""
   Call sbCargaInicial
End If

End Sub

Private Sub sbAplicar()
Dim strSQL As String, lng As Long, vTipo As String


vTipo = Mid(cboPermiso.Text, 1, 1)

If vTipo = "R" And gEntidad.Tipo = "G" Then
   MsgBox "Los Grupos no se les puede restringir opciones por razones de orden...", vbInformation
   Exit Sub
End If

Me.MousePointer = vbHourglass


With vTree.Nodes

 PrgBar.Max = .Count + 1
 PrgBar.Value = 1
 
 For lng = 1 To .Count
   If Right(.Item(lng).Key, 1) = "O" Then
     'Si no estaba marcada y ahora si (Incluye)
     If .Item(lng).Tag = 0 And .Item(lng).Checked Then
        strSQL = "insert permisos(id_opt,nombre,tipo,estado) values(" _
               & fxIndiceMultiple(.Item(lng).Key, "N") & ",'" & lblID.Caption _
               & "','" & gEntidad.Tipo & "','" & vTipo & "')"
               
        glogon.Conection.Execute strSQL
     End If
     'Si estaba marcada y ahora no (Excluye)
     If .Item(lng).Tag = 1 And Not .Item(lng).Checked Then
        strSQL = "delete permisos where id_opt = " & fxIndiceMultiple(.Item(lng).Key, "N") _
               & " and estado = '" & vTipo & "' and nombre = '" & lblID.Caption _
               & "' and tipo = '" & gEntidad.Tipo & "'"
        glogon.Conection.Execute strSQL
     End If
     
   End If
 
  PrgBar.Value = PrgBar.Value + 1
 Next lng

End With

If gEntidad.Tipo = "U" Then
  Call sbSEGCuentaLog("16", cboPermiso.Text, glogon.Usuario, txtSujeto.Text)
Else
  Call sbSEGCuentaLog("14", cboPermiso.Text & "...:" & txtSujeto.Text, glogon.Usuario)
End If

Me.MousePointer = vbDefault
MsgBox "Cambios Aplicados Satisfactoriamente...", vbInformation


End Sub



Public Sub cmdDeshacer_Click()
   Call sbCargaInicial
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    
    If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
     'Grupos
        strSQL = "select Top 1 id_grupo as LLave,nombre from grupos"
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where nombre > '" & txtSujeto & "' order by nombre asc"
        Else
           strSQL = strSQL & " where nombre < '" & txtSujeto & "' order by nombre desc"
        End If
    
    Else
     'Usuarios
        strSQL = "select Top 1 UserID as LLave,nombre,descripcion from usuarios"
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where nombre > '" & txtSujeto & "' order by nombre asc"
        Else
           strSQL = strSQL & " where nombre < '" & txtSujeto & "' order by nombre desc"
        End If
    
    End If
    
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF And Not rs.BOF Then
      txtSujeto = rs!Nombre
      txtSujeto.Tag = rs!llave
      lblID.Caption = rs!llave
      Call sbCargaInicial
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
    
  MsgBox Err.Description, vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub

Private Sub Form_Load()
vModulo = 13

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True


'Inicializa Tipos
cboTipo.AddItem "Listado de Grupos Disponibles"
cboTipo.ItemData(cboTipo.NewIndex) = 1
cboTipo.AddItem "Listado de Usuarios Disponibles"
cboTipo.ItemData(cboTipo.NewIndex) = 2

cboPermiso.AddItem "Autorizaciones"
cboPermiso.AddItem "Restricciones"
cboPermiso.Text = "Autorizaciones"
 
Select Case gEntidad.Tipo
 Case "G"
    cboTipo.Text = "Listado de Grupos Disponibles"
    txtSujeto = gEntidad.Grupo
    txtSujeto.Tag = gEntidad.GrpID
    lblID.Caption = gEntidad.GrpID
 Case "U"
    cboTipo.Text = "Listado de Usuarios Disponibles"
    txtSujeto = gEntidad.Usuario
    txtSujeto.Tag = gEntidad.UserID
    lblID.Caption = gEntidad.UserID
 Case Else
    cboTipo.Text = "Listado de Usuarios Disponibles"
    txtSujeto = gEntidad.Usuario
    txtSujeto.Tag = gEntidad.UserID
    lblID.Caption = gEntidad.UserID
End Select

Call Formularios(Me)
Call RefrescaTags(Me)

tlb.Enabled = cmdAplicar.Enabled

End Sub

Private Sub TimerInicia_Timer()
TimerInicia.Interval = 0
Call sbCargaInicial
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Aplicar"
   Call sbAplicar
  Case "Deshacer"
   Call sbCargaInicial
End Select
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

Private Sub txtSujeto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
  If gEntidad.Tipo = "G" Then
    gBusquedas.Consulta = "Select id_grupo as Codigo,nombre from grupos"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
  Else
    gBusquedas.Consulta = "Select UserId as Codigo,nombre,descripcion from Usuarios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
  End If
  frmBusquedas.Show vbModal
  lblID.Caption = gBusquedas.Resultado
  txtSujeto.Tag = gBusquedas.Resultado
  txtSujeto = gBusquedas.Resultado2
  Call sbCargaInicial
End If
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


Private Sub vTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim lng As Long, vKey As String, vKey2 As String
Dim lng2 As Long

Me.MousePointer = vbHourglass

Select Case Right(Node.Key, 1)
  Case "F" 'Formulario
    vKey = fxIndiceMultiple(Node.Key, "N")
    For lng = Node.Index + 1 To vTree.Nodes.Count
      If vTree.Nodes(lng).Key <> "" Then
            If vKey = fxIndiceMultiple(vTree.Nodes(lng).Key, "T") And Right(vTree.Nodes(lng).Key, 1) = "O" Then
                  vTree.Nodes(lng).Checked = Node.Checked
                  If Mid(cboPermiso.Text, 1, 1) = "A" Then
                       vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, txtAutorizado.ForeColor, vbBlack)
                       vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, txtAutorizado.BackColor, vbWhite)
                       vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                  Else
                       vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, txtDenegado.ForeColor, vbBlack)
                       vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, txtDenegado.BackColor, vbWhite)
                       vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                  End If
            End If
      End If
    Next lng
  
  Case "M" 'Modulo
    'En orden logico se localizan primero todos los formularios y debajo de este las opciones
    vKey = fxIndiceCodigo(Node.Key)
    
    
    For lng = Node.Index + 1 To vTree.Nodes.Count
     If Right(vTree.Nodes(lng).Key, 1) = "F" Then
          If vKey = fxIndiceMultiple(vTree.Nodes(lng).Key, "T") Then
                'Marca el Formulario
                vTree.Nodes(lng).Checked = Node.Checked
                
                If Mid(cboPermiso.Text, 1, 1) = "A" Then
                     vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, txtAutorizado.ForeColor, vbBlack)
                     vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, txtAutorizado.BackColor, vbWhite)
                     vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                Else
                     vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, txtDenegado.ForeColor, vbBlack)
                     vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, txtDenegado.BackColor, vbWhite)
                     vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                End If
                
                'Marcar las Opciones del Formulario
                vKey2 = fxIndiceMultiple(vTree.Nodes(lng).Key, "N")
                For lng2 = vTree.Nodes(lng).Index + 1 To vTree.Nodes.Count
                 If vTree.Nodes(lng2).Key <> "" Then
                  If vKey2 = fxIndiceMultiple(vTree.Nodes(lng2).Key, "T") And Right(vTree.Nodes(lng2).Key, 1) = "O" Then
                        vTree.Nodes(lng2).Checked = Node.Checked
                        If Mid(cboPermiso.Text, 1, 1) = "A" Then
                             vTree.Nodes(lng2).ForeColor = IIf(vTree.Nodes(lng2).Checked, txtAutorizado.ForeColor, vbBlack)
                             vTree.Nodes(lng2).BackColor = IIf(vTree.Nodes(lng2).Checked, txtAutorizado.BackColor, vbWhite)
                             vTree.Nodes(lng2).Bold = IIf(vTree.Nodes(lng2).Checked, True, False)
                        Else
                             vTree.Nodes(lng2).ForeColor = IIf(vTree.Nodes(lng2).Checked, txtDenegado.ForeColor, vbBlack)
                             vTree.Nodes(lng2).BackColor = IIf(vTree.Nodes(lng2).Checked, txtDenegado.BackColor, vbWhite)
                             vTree.Nodes(lng2).Bold = IIf(vTree.Nodes(lng2).Checked, True, False)
                        End If
                    End If
                  End If
                Next lng2
                
          End If
        
         End If ' right F
        Next lng
   
End Select

If Node.Index > 1 Then
    If Mid(cboPermiso.Text, 1, 1) = "A" Then
         Node.ForeColor = IIf(Node.Checked, txtAutorizado.ForeColor, vbBlack)
         Node.BackColor = IIf(Node.Checked, txtAutorizado.BackColor, vbWhite)
         Node.Bold = IIf(Node.Checked, True, False)
    Else
         Node.ForeColor = IIf(Node.Checked, txtDenegado.ForeColor, vbBlack)
         Node.BackColor = IIf(Node.Checked, txtDenegado.BackColor, vbWhite)
         Node.Bold = IIf(Node.Checked, True, False)
    End If
End If

Me.MousePointer = vbDefault

End Sub
