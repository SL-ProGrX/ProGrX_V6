VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmUS_DerechosXOpcion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Permisos x Opción"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Extender"
      Top             =   120
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_DerechosXOpcion.frx":0000
   End
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
            Picture         =   "frmUS_DerechosXOpcion.frx":0708
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6F6A
            Key             =   "imgFormularios"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":D7CC
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":1402E
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":1A890
            Key             =   "x2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":210F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":27954
            Key             =   "imgOpcionDetalle"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":2E1B6
            Key             =   "x1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":34A18
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":3B27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":41ADC
            Key             =   "imgFrmOpcion"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":4833E
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":4EBA0
            Key             =   "imgGrupoDetalle"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":55402
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":5BC64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerInicia 
      Interval        =   10
      Left            =   4800
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
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
            Picture         =   "frmUS_DerechosXOpcion.frx":624C6
            Key             =   "imgFrm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":62DA0
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6367A
            Key             =   "imgModulo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView vTree 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13573
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgExp01"
      BorderStyle     =   1
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
            Picture         =   "frmUS_DerechosXOpcion.frx":63F54
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6426E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":64B48
            Key             =   "Cuestion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65422
            Key             =   "CheckList"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6573C
            Key             =   "User"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65A56
            Key             =   "UserGroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":65D70
            Key             =   "Keys"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":6664A
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":66F24
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":677FE
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":680D8
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":683F2
            Key             =   "SearchFolder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":68CCC
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_DerechosXOpcion.frx":695A6
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      CausesValidation=   0   'False
      Height          =   7695
      Left            =   5520
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13573
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2188
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   5362
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboPermiso 
      Height          =   330
      Left            =   120
      TabIndex        =   2
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   615
      Left            =   0
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   720
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
Attribute VB_Name = "frmUS_DerechosXOpcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vPaso As Boolean


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
Dim strSQL As String, rs As New ADODB.Recordset
Dim xNode As Node, lng As Long


Me.MousePointer = vbHourglass

With vTree
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "US", "Root")
  xNode.Bold = True
  
  strSQL = "select * from us_modulos order by modulo"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos("US", Trim(rs!Nombre), "imgModulo", True, "0x0" & rs!Modulo & "M")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close

  strSQL = "select * from us_formularios order by modulo,Descripcion"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!Modulo & "M", Trim(rs!Descripcion), "imgFrm", True, "0x0" & rs!Modulo & "-" & Trim(rs!Formulario) & "F")
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close
  
  strSQL = "select * from US_Opciones"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos("0x0" & rs!Modulo & "-" & Trim(rs!Formulario) & "F", Trim(rs!Opcion_descripcion), "imgOpcion", False, "0x0" & Trim(rs!Formulario) & "-" & rs!cod_Opcion & "O")
   
     .Nodes.Item(.Nodes.Count).Tag = 1
   
   rs.MoveNext
  Loop
  rs.Close


   xNode.Expanded = True

End With


Me.MousePointer = vbDefault

Me.Show

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

 If Item.Checked Then
    strSQL = "insert Us_Rol_Permisos(cod_Opcion,Cod_Rol,estado,registro_Fecha,registro_usuario) values(" _
           & lblX.Tag & ",'" & Item.Tag & "','" & vTipo & "',getdate(),'" & glogon.Usuario & "')"
 Else
    strSQL = "delete Us_Rol_Permisos where cod_Opcion = " & lblX.Tag _
           & " and estado = '" & vTipo & "' and cod_Rol = '" & Item.Tag & "'"
 End If
 Call ConectionExecute(strSQL, 1)
 
 Call sbSEGCuentaLog("14", cboPermiso.Text & "...:" & Item.Text, glogon.Usuario)
 
End Sub


Private Sub TimerInicia_Timer()
    TimerInicia.Interval = 0
    Call sbCargaInicial
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

Private Sub sbCargaDatosLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vTipo As String

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

vTipo = Mid(cboPermiso.Text, 1, 1)


 strSQL = "Select R.cod_rol,R.descripcion, isnull(P.Estado,'Z') as 'Estado'" _
        & " from US_ROLES R left join US_ROL_PERMISOS P on R.cod_Rol = P.cod_Rol and P.cod_Opcion = " & lblX.Tag _
        & " and P.Estado = '" & vTipo & "'" _
        & " where R.Activo = 1" _
        & " order by isnull(P.Estado,'Z'), R.Descripcion"
  
 Call OpenRecordSet(rs, strSQL, 1)
 Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!cod_rol)
       itmX.SubItems(1) = rs!Descripcion
       itmX.Tag = rs!cod_rol
       
       If rs!ESTADO <> "Z" Then
            itmX.Checked = vbChecked
            
            If vTipo = "A" Then
                itmX.ForeColor = vbBlue
            Else
                itmX.ForeColor = vbRed
            End If
       End If
   rs.MoveNext
 Loop
 rs.Close
 

Me.MousePointer = vbDefault

End Sub

Private Sub vTree_NodeClick(ByVal Node As MSComctlLib.Node)
 If Node.Image = "imgOpcion" Then
    lblX.Tag = fxIndiceMultiple(Node.Key, "N")
    scTitulo.Caption = Node.FullPath
    lblX.Caption = cboPermiso.Text & " ¦ " & Node.Text
    Call sbCargaDatosLsw
    
     
 Else
    lblX.Tag = 0
    scTitulo.Caption = ""
    lblX.Caption = ""
 End If

End Sub
