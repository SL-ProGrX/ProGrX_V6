VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCC_ConsultaCuentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Cuentas"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Explorador"
      TabPicture(0)   =   "frmSIF_CuentasContables.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ArbolExp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "frmSIF_CuentasContables.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).Control(1)=   "txtCriterio"
      Tab(1).Control(2)=   "cboTipo"
      Tab(1).Control(3)=   "cmdBuscar"
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ListView lsw 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8916
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7126
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Movimientos"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtCriterio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73440
         TabIndex        =   4
         Top             =   360
         Width           =   5415
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   -67935
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   5400
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9525
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgExplorer"
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
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   7200
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":0914
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":11F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":150C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":1828
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":1B44
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":1E60
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":217C
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":2A58
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":3334
            Key             =   "imgAsientos"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_CuentasContables.frx":3650
            Key             =   "imgCuentas"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCC_ConsultaCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node

Private Sub ArbolExp_DblClick()
If vNode.Bold = False Then
    If vNode.Text <> "Cuentas" And Right(vNode.Key, 1) <> "T" Then
        gCuenta = fxIndiceCodigo(vNode.Key)
        Unload frmCC_ConsultaCuentas
    End If
Else
  MsgBox "Esta Cuenta No Acepta Movimientos, Verificar...", vbExclamation
End If
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

strSQL = "select cod_cuenta,descripcion,acepta_movimientos from cuentas " _
       & " where cod_empresa = " & GLOBALES.gEnlace
Select Case cboTipo.Text
  Case "Por - Cuenta"
     strSQL = strSQL & " and cod_cuenta like '" & txtCriterio & "%'"
  Case "Por - Descripción"
     strSQL = strSQL & " and descripcion like '%" & txtCriterio & "%'"
End Select
strSQL = strSQL & " order by cod_cuenta"
rs.Open strSQL, glogon.Conection, adOpenStatic
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , fxgCntCuentaFormato(True, rs!cod_cuenta))
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = IIf((rs!Acepta_movimientos = "S"), "SI", "NO")
      
      If itmX.SubItems(2) = "NO" Then itmX.ForeColor = vbRed
  
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
cboTipo.AddItem "Por - Cuenta"
cboTipo.AddItem "Por - Descripción"

cboTipo.Text = "Por - Cuenta"
gCuenta = ""
Call sbRefrescaArbol

ssTab.Tab = 0

End Sub


Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Cuentas", "Cuentas", "imgRoot")
  'Crear Arbol Inicial
  
  strSQL = "select tipo_cuenta,Descripcion from tipos_cuentas where cod_empresa = " & GLOBALES.gEnlace
  rs.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rs.EOF
    Call sbCreaNodos("Cuentas", rs!Descripcion, "imgCuentas", True, "S", "0x0" & rs!tipo_cuenta & "T")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With


End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Cuentas" Then

Select Case Right(Node.Key, 1)
        
    Case "T" 'Tipos de Cuentas
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from cuentas where cuenta_madre = ''" _
               & " and cod_empresa = " & GLOBALES.gEnlace _
               & " and tipo_cuenta = '" & fxIndiceCodigo(Node.Key) & "'"
        rs.Open strSQL, glogon.Conection, adOpenStatic
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxgCntCuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", True, rs!Acepta_movimientos, "0x0" & fxgCntCuentaFormato(False, rs!cod_cuenta) & "C")
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from cuentas where cuenta_madre = '" & fxgCntCuentaFormato(False, fxIndiceCodigo(Node.Key)) _
               & "' and cod_empresa = " & GLOBALES.gEnlace
        rs.Open strSQL, glogon.Conection, adOpenStatic
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxgCntCuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", True, rs!Acepta_movimientos, "0x0" & fxgCntCuentaFormato(False, rs!cod_cuenta) & "C")
          rs.MoveNext
        Loop
        rs.Close
End Select

End If

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
               , vAcepta As String, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    If vAcepta = "N" Then nodX.Bold = True
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
    
End Sub

Private Sub lsw_DblClick()
If lsw.ListItems.Count > 0 Then
  If lsw.SelectedItem.SubItems(2) = "SI" Then
    gCuenta = fxgCntCuentaFormato(False, lsw.SelectedItem)
    Unload frmCC_ConsultaCuentas
  Else
    MsgBox "Esta Cuenta No Recibe Movimientos, Verificar...", vbExclamation
  End If
End If
End Sub

Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdBuscar_Click
End Sub
