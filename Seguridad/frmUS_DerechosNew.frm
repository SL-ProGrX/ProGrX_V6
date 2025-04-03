VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmUS_DerechosNew 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de permisos a los módulos"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   340
      Index           =   0
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   600
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_DerechosNew.frx":0000
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   2
      Top             =   9345
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   630
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   1560
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
            Picture         =   "frmUS_DerechosNew.frx":0727
            Key             =   "imgFrm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":1001
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":18DB
            Key             =   "imgModulo"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerInicia 
      Interval        =   10
      Left            =   7440
      Top             =   1080
   End
   Begin MSComctlLib.TreeView vTree 
      Height          =   8055
      Left            =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14208
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
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
            Picture         =   "frmUS_DerechosNew.frx":21B5
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":24CF
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":2DA9
            Key             =   "Cuestion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":3683
            Key             =   "CheckList"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":399D
            Key             =   "User"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":3CB7
            Key             =   "UserGroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":3FD1
            Key             =   "Keys"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":48AB
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":5185
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":5A5F
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":6339
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":6653
            Key             =   "SearchFolder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":6F2D
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosNew.frx":7807
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtId 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   615
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSujeto 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   615
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10398
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "(Presione F4 para Consultar)"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   345
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Extender"
      Top             =   120
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   600
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_DerechosNew.frx":7B21
   End
   Begin XtremeSuiteControls.ComboBox cboPermiso 
      Height          =   330
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   340
      Index           =   1
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   600
      _StockProps     =   79
      Caption         =   "Deshacer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_DerechosNew.frx":8229
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   1085
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
      SubItemCaption  =   -1  'True
      VisualTheme     =   7
      Alignment       =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   661
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
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmUS_DerechosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNode As Node, vScroll As Boolean


Private Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xKey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next


Set nodX = vTree.Nodes.Add(vPadre, tvwChild)
    nodX.Image = vImagen
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    If xKey = "N" Then
        nodX.Key = vTexto & "0x0" & vTree.Nodes.Count & "ID"
    Else
        nodX.Key = xKey
    End If
    
Set vNode = nodX

End Sub


Private Sub sbCargaInicial()
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
  
  strSQL = "select * from US_modulos order by modulo"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos("US", Trim(rs!Nombre), "imgModulo", True, "0x0" & rs!Modulo & "M")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close

  strSQL = "select * from US_formularios order by descripcion"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!Modulo & "M", Trim(rs!Descripcion), "imgFrm", True, "0x0" & rs!Modulo & "-" & Trim(rs!Formulario) & "F")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close
  
  strSQL = "select DISTINCT O.*,isnull(P.ESTADO,'Z') as 'PermisoEstado'" _
         & " from US_opciones O inner join US_formularios F on O.formulario = F.formulario" _
         & " left join US_ROL_Permisos P on O.cod_Opcion = P.cod_Opcion and P.cod_Rol = '" & txtSujeto.Tag & "'" _
         & " and P.estado = '" & Mid(cboPermiso.Text, 1, 1) _
         & "' order by O.Opcion_descripcion"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!Modulo & "-" & Trim(rs!Formulario) & "F", Trim(rs!Opcion_descripcion), "imgOpcion", False, "0x0" & Trim(rs!Formulario) & "-" & rs!cod_Opcion & "O")
   
   If rs!PermisoEstado <> "Z" Then
'     .Nodes.Item(.Nodes.Count).ForeColor = IIf((Mid(cboPermiso.Text, 1, 1) = "A"), vbBlue, vbRed)
'     .Nodes.Item(.Nodes.Count).Checked = True
'     .Nodes.Item(.Nodes.Count).Bold = True
'     .Nodes.Item(.Nodes.Count).Tag = 1
     
     vNode.ForeColor = IIf((Mid(cboPermiso.Text, 1, 1) = "A"), vbBlue, vbRed)
     vNode.Checked = True
     vNode.Bold = True
     vNode.Tag = 1
   
   Else
'     .Nodes.Item(.Nodes.Count).Tag = 0
     vNode.Tag = 0
   End If
   
   rs.MoveNext
  Loop
  rs.Close

 
   xNode.Expanded = True

End With


Me.MousePointer = vbDefault


End Sub



Private Sub btnAccion_Click(Index As Integer)
Select Case Index
  Case 0 '"Aplicar"
   Call sbAplicar
  Case 1 '"Refrescar"
   Call sbCargaInicial
End Select
End Sub

Private Sub btnRefresh_Click()
Dim lng As Long

With vTree.Nodes
 For lng = 1 To .Count
  If Right(.Item(lng).Key, 1) = "M" Then
    .Item(lng).Expanded = IIf(.Item(lng).Expanded, False, True)
  End If
 Next lng
End With
End Sub

Private Sub cboPermiso_Click()
   Call sbCargaInicial
End Sub

Private Sub cboTipo_Click()
If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
   gEntidad.Tipo = "R"
   txtSujeto.Text = ""
   txtSujeto.Tag = ""
   txtId.Text = ""
   
   Call sbCargaInicial
Else
   gEntidad.Tipo = "U"
   txtSujeto.Text = ""
   txtSujeto.Tag = ""
   txtId.Text = ""
   Call sbCargaInicial
End If

End Sub

Private Sub sbAplicar()
Dim lng As Long, vTipo As String


vTipo = Mid(cboPermiso.Text, 1, 1)

If vTipo = "R" And gEntidad.Tipo = "R" Then
   MsgBox "A los ROLES no se les puede restringir opciones por razones de orden...", vbInformation
   Exit Sub
End If

Me.MousePointer = vbHourglass


With vTree.Nodes

 prgBar.Max = .Count + 1
 prgBar.Value = 1
 
 strSQL = ""
 
 For lng = 1 To .Count
   If Right(.Item(lng).Key, 1) = "O" Then
     'Si no estaba marcada y ahora si (Incluye)
     If .Item(lng).Tag = 0 And .Item(lng).Checked Then
        strSQL = strSQL & Space(10) & "insert US_ROL_permisos(cod_Opcion,cod_Rol,estado,registro_Fecha,registro_usuario) values(" _
               & fxIndiceMultiple(.Item(lng).Key, "N") & ",'" & txtId.Text _
               & "','" & vTipo & "',getdate(),'" & glogon.Usuario & "')"
               
'        Call ConectionExecute(strSQL, 1)
        .Item(lng).Tag = 1
        .Item(lng).ForeColor = vbBlue
        .Item(lng).Bold = True
        
     End If
     'Si estaba marcada y ahora no (Excluye)
     If .Item(lng).Tag = 1 And Not .Item(lng).Checked Then
        strSQL = strSQL & Space(10) & "delete US_ROL_Permisos where cod_Opcion = " & fxIndiceMultiple(.Item(lng).Key, "N") _
               & " and estado = '" & vTipo & "' and cod_ROL = '" & txtId.Text & "'"
'        Call ConectionExecute(strSQL, 1)
        
        .Item(lng).Tag = 0
        .Item(lng).ForeColor = vbBlack
        .Item(lng).Bold = False

     End If
     
   End If
 
  prgBar.Value = prgBar.Value + 1
  
  'Ejecuta el Lote
  If Len(strSQL) > 20000 Then
    Call ConectionExecute(strSQL, 1)
    strSQL = ""
  End If
 Next lng

End With


'Ejecuta el Lote
If Len(strSQL) > 0 Then
  Call ConectionExecute(strSQL, 1)
  strSQL = ""
End If


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

On Error GoTo vError

If vScroll Then
    
'    If cboTipo.ItemData(cboTipo.ListIndex) = 1 Then
'     'Grupos
'        strSQL = "select Top 1 id_grupo as LLave,nombre from grupos"
'        If FlatScrollBar.Value = 1 Then
'           strSQL = strSQL & " where nombre > '" & txtSujeto & "' order by nombre asc"
'        Else
'           strSQL = strSQL & " where nombre < '" & txtSujeto & "' order by nombre desc"
'        End If
'
'    Else
'     'Usuarios
'        strSQL = "select Top 1 UserID as LLave,nombre,descripcion from usuarios"
'        If FlatScrollBar.Value = 1 Then
'           strSQL = strSQL & " where nombre > '" & txtSujeto & "' order by nombre asc"
'        Else
'           strSQL = strSQL & " where nombre < '" & txtSujeto & "' order by nombre desc"
'        End If
'
'    End If
    
    
        strSQL = "select Top 1 cod_Rol as LLave,Descripcion as 'Nombre' from US_ROLES"
        If FlatScrollBar.Value = 1 Then
           strSQL = strSQL & " where Descripcion > '" & txtSujeto & "' order by Descripcion asc"
        Else
           strSQL = strSQL & " where Descripcion < '" & txtSujeto & "' order by Descripcion desc"
        End If
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtSujeto = rs!Nombre
      txtSujeto.Tag = rs!llave
      txtId.Text = rs!llave
      Call sbCargaInicial
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
    
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
'vModulo = 13

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True


'Inicializa Tipos
cboTipo.AddItem "Listado de Roles Disponibles"
cboTipo.ItemData(cboTipo.ListCount - 1) = "1"
cboTipo.AddItem "Listado de Usuarios Disponibles"
cboTipo.ItemData(cboTipo.ListCount - 1) = "2"

cboPermiso.AddItem "Autorizaciones"
cboPermiso.AddItem "Restricciones"
cboPermiso.Text = "Autorizaciones"
 
Select Case gEntidad.Tipo
 Case "G", "R"
    cboTipo.Text = "Listado de Roles Disponibles"
    txtSujeto = gEntidad.Rol_Name
    txtSujeto.Tag = gEntidad.Rol_Id
    txtId.Text = gEntidad.Rol_Id
 Case "U"
    cboTipo.Text = "Listado de Usuarios Disponibles"
    txtSujeto = gEntidad.Usuario
    txtSujeto.Tag = gEntidad.UserID
    txtId.Text = gEntidad.UserID
 Case Else
    cboTipo.Text = "Listado de Usuarios Disponibles"
    txtSujeto = gEntidad.Usuario
    txtSujeto.Tag = gEntidad.UserID
    txtId.Text = gEntidad.UserID
End Select

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub TimerInicia_Timer()
TimerInicia.Interval = 0
Call sbCargaInicial
End Sub



Private Sub txtSujeto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Consulta = "Select Cod_Rol as Codigo,Descripcion from US_ROLES"
    gBusquedas.Columna = "Descripcion"
    gBusquedas.Orden = "Descripcion"
  
  frmBusquedas.Show vbModal
  txtId.Text = gBusquedas.Resultado
  txtSujeto.Tag = gBusquedas.Resultado
  txtSujeto = gBusquedas.Resultado2
  Call sbCargaInicial
End If
End Sub

Private Function fxIndiceCodigo(xKey As String) As String
xKey = Mid(xKey, 4, Len(xKey))
xKey = Mid(xKey, 1, Len(xKey) - 1)
fxIndiceCodigo = xKey
End Function

Private Function fxIndiceMultiple(xKey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xKey = fxIndiceCodigo(xKey)

blnPaso = True

If xKey = "" Then
  fxIndiceMultiple = ""
  Exit Function
End If

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xKey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xKey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xKey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xKey, i, 50) '50 es un default ningun asiento es tan largo

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
                       vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, vbGreen, vbBlack)
                       vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, vbWhite, vbWhite)
                       vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                  Else
                       vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, vbRed, vbBlack)
                       vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, vbWhite, vbWhite)
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
                     vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, vbGreen, vbBlack)
                     vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, vbWhite, vbWhite)
                     vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                Else
                     vTree.Nodes(lng).ForeColor = IIf(vTree.Nodes(lng).Checked, vbRed, vbBlack)
                     vTree.Nodes(lng).BackColor = IIf(vTree.Nodes(lng).Checked, vbWhite, vbWhite)
                     vTree.Nodes(lng).Bold = IIf(vTree.Nodes(lng).Checked, True, False)
                End If
                
                'Marcar las Opciones del Formulario
                vKey2 = fxIndiceMultiple(vTree.Nodes(lng).Key, "N")
                For lng2 = vTree.Nodes(lng).Index + 1 To vTree.Nodes.Count
                 If vTree.Nodes(lng2).Key <> "" Then
                  If vKey2 = fxIndiceMultiple(vTree.Nodes(lng2).Key, "T") And Right(vTree.Nodes(lng2).Key, 1) = "O" Then
                        vTree.Nodes(lng2).Checked = Node.Checked
                        If Mid(cboPermiso.Text, 1, 1) = "A" Then
                             vTree.Nodes(lng2).ForeColor = IIf(vTree.Nodes(lng2).Checked, vbGreen, vbBlack)
                             vTree.Nodes(lng2).BackColor = IIf(vTree.Nodes(lng2).Checked, vbWhite, vbWhite)
                             vTree.Nodes(lng2).Bold = IIf(vTree.Nodes(lng2).Checked, True, False)
                        Else
                             vTree.Nodes(lng2).ForeColor = IIf(vTree.Nodes(lng2).Checked, vbRed, vbBlack)
                             vTree.Nodes(lng2).BackColor = IIf(vTree.Nodes(lng2).Checked, vbWhite, vbWhite)
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
         Node.ForeColor = IIf(Node.Checked, vbGreen, vbBlack)
         Node.BackColor = IIf(Node.Checked, vbWhite, vbWhite)
         Node.Bold = IIf(Node.Checked, True, False)
    Else
         Node.ForeColor = IIf(Node.Checked, vbRed, vbBlack)
         Node.BackColor = IIf(Node.Checked, vbWhite, vbWhite)
         Node.Bold = IIf(Node.Checked, True, False)
    End If
End If

Me.MousePointer = vbDefault

End Sub
