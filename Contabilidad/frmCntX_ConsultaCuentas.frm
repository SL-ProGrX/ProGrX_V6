VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCntX_ConsultaCuentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Cuentas"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Explorador"
      TabPicture(0)   =   "frmCntX_ConsultaCuentas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ArbolExp"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "frmCntX_ConsultaCuentas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FlatScrollBarX"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBuscar"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCriterio"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lsw"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFiltroCuenta(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtFiltroCuenta(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtNivel"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtNivel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "5"
         Top             =   720
         Width           =   750
      End
      Begin VB.TextBox txtFiltroCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtFiltroCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4575
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8070
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
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   7065
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   5400
         Left            =   -74880
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
      Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
         Height          =   285
         Left            =   7200
         TabIndex        =   8
         Top             =   720
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         _Version        =   393216
         Arrows          =   65536
         Min             =   1
         Max             =   5
         Orientation     =   1572865
         Value           =   5
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Niveles"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   6360
         TabIndex        =   12
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Filtro Cuentas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConsultaCuentas.frx":0038
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConsultaCuentas.frx":0145
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConsultaCuentas.frx":0261
            Key             =   "imgCuentas"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConsultaCuentas.frx":037F
            Key             =   "imgAsientos"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCntX_ConsultaCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node

Private Sub ArbolExp_DblClick()
If vNode.Text <> "Cuentas" And Right(vNode.Key, 1) <> "T" Then
    gCuenta = fxIndiceCodigo(vNode.Key)
    Unload Me
End If
End Sub


Private Sub sbLlenaLswCtas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas " _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and nivel <= " & txtNivel

If Trim(txtFiltroCuenta(0).Text) <> "" Then strSQL = strSQL & " and cod_cuenta like '" & Trim(txtFiltroCuenta(0).Text) & "%'"
If Trim(txtFiltroCuenta(1).Text) <> "" Then strSQL = strSQL & " and descripcion like '%" & Trim(txtFiltroCuenta(1).Text) & "%'"

strSQL = strSQL & " order by cod_cuenta"
Call OpenRecordSet(rs, strSQL, 0)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , fxCntX_CuentaFormato(True, rs!cod_cuenta))
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = IIf((rs!Acepta_movimientos = 1), "Sí", "No")
      
      If itmX.SubItems(2) = "No" Then itmX.ForeColor = vbRed
  
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
       

End Sub

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Text <> "Cuentas" And Right(Node.Key, 1) <> "T" Then
    gCuenta = fxIndiceCodigo(Node.Key)
    Unload Me
End If
End Sub

Private Sub FlatScrollBarX_Change()
txtNivel = FlatScrollBarX.Value
Call sbLlenaLswCtas
End Sub

Private Sub Form_Load()

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
  
  strSQL = "select tipo_cuenta,Descripcion from CntX_Tipos_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    Call sbCreaNodos("Cuentas", rs!Descripcion, "imgCuentas", True, "0x0" & rs!tipo_cuenta & "T")
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
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = ''" _
               & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and tipo_cuenta = '" & fxIndiceCodigo(Node.Key) & "'"
               
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", IIf((rs!Acepta_movimientos = 0), True, False) _
                , "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", rs!Acepta_movimientos)
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
    
        strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = '" & fxCntX_CuentaFormato(False, fxIndiceCodigo(Node.Key)) _
               & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", IIf((rs!Acepta_movimientos = 0), True, False) _
                , "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C", rs!Acepta_movimientos)
          rs.MoveNext
        Loop
        rs.Close
End Select

End If

End Sub


Private Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
                       , Optional xkey As String = "N", Optional pAceptaMov As Integer = 1)
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
    
    If pAceptaMov = 0 Then
       nodX.Bold = True
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub

Private Sub lsw_DblClick()
If lsw.ListItems.Count > 0 Then
  gCuenta = fxCntX_CuentaFormato(False, lsw.SelectedItem)
  Unload Me
End If
End Sub

Private Sub txtFiltroCuenta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbLlenaLswCtas
End Sub
