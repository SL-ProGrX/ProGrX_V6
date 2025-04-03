VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCntX_ConsultaCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Cuentas"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7095
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   10515
      _Version        =   1441793
      _ExtentX        =   18547
      _ExtentY        =   12515
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnEstilo 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   8640
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Consulta de Cuentas"
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_ConsultaCuentas.frx":0000
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
      Height          =   285
      Left            =   9840
      TabIndex        =   0
      Top             =   960
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      Min             =   1
      Max             =   8
      Orientation     =   1638401
      Value           =   5
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   5400
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9525
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   Begin XtremeSuiteControls.FlatEdit txtFiltroCuenta 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltroCuenta 
      Height          =   330
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNivel 
      Height          =   330
      Left            =   8880
      TabIndex        =   8
      Top             =   960
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
      Text            =   "8"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   330
      Left            =   7200
      TabIndex        =   9
      Top             =   960
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.PushButton btnEstilo 
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   8640
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Explorador de Cuentas"
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_ConsultaCuentas.frx":0708
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7200
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Niveles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   8880
      TabIndex        =   1
      Top             =   720
      Width           =   750
   End
   Begin VB.Image imgBanner 
      Height          =   1350
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmCntX_ConsultaCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vNode As Node, vContabilidad As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub ArbolExp_DblClick()
If vNode.Text <> "Cuentas" And Right(vNode.Key, 1) <> "T" Then
    gCuenta = fxIndiceCodigo(vNode.Key)
    UnLoad Me
End If
End Sub


Private Sub sbLlenaLswCtas()

On Error GoTo vError

Me.MousePointer = vbHourglass


Call fxSysCleanTxtInject(txtFiltroCuenta(0))
Call fxSysCleanTxtInject(txtFiltroCuenta(1))

strSQL = "exec spCntX_Consulta_Cuentas " & vContabilidad _
      & ",'" & txtFiltroCuenta(0).Text _
      & "','" & txtFiltroCuenta(1).Text _
      & "','" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & txtNivel.Text _
      & ", " & vModulo & ", '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Cod_Cuenta_Mask)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = IIf((rs!Acepta_movimientos = 1), "Sí", "No")
      
      If itmX.SubItems(2) = "No" Then
            itmX.Bold = True
            itmX.TextBackColor = RGB(214, 234, 248)
      End If
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
       
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
       

End Sub

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Text <> "Cuentas" And Right(Node.Key, 1) <> "T" Then
    gCuenta = fxIndiceCodigo(Node.Key)
    UnLoad Me
End If
End Sub

Private Sub btnEstilo_Click(Index As Integer)

Select Case Index
  Case 0 '"Consulta
    lsw.Visible = True
    ArbolExp.Visible = False
    
  Case 1 'Explorer
    ArbolExp.Visible = True
    ArbolExp.Left = lsw.Left
    ArbolExp.Top = lsw.Top
    ArbolExp.Width = lsw.Width
    ArbolExp.Height = lsw.Height
    
    lsw.Visible = False
End Select
  
txtFiltroCuenta.Item(0).Visible = lsw.Visible
txtFiltroCuenta.Item(1).Visible = lsw.Visible
cboDivisa.Visible = lsw.Visible

txtNivel.Visible = lsw.Visible

FlatScrollBarX.Visible = lsw.Visible

lblTitulo.Item(0).Visible = lsw.Visible
lblTitulo.Item(1).Visible = lsw.Visible
lblTitulo.Item(2).Visible = lsw.Visible
lblTitulo.Item(3).Visible = lsw.Visible


End Sub

Private Sub FlatScrollBarX_Change()
txtNivel = FlatScrollBarX.Value
Call sbLlenaLswCtas
End Sub

Private Sub Form_Load()

'vContabilidad = GLOBALES.gEnlace
'

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

If vModulo = 20 Then
    vContabilidad = gCntX_Parametros.CodigoConta
Else
    vContabilidad = GLOBALES.gEnlace
End If

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Cuenta", 2500
lsw.ColumnHeaders.Add , , "Nombre", 6000
lsw.ColumnHeaders.Add , , "Movimientos", 1500, vbCenter


strSQL = "select rtrim(cod_divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from CntX_Divisas where cod_contabilidad = " _
       & vContabilidad & " order by divisa_local desc"

Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

Call sbRefrescaArbol

End Sub


Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Cuentas", "Cuentas") 'imgRoot
  'Crear Arbol Inicial
  
  strSQL = "select TIPO_CUENTA,Descripcion from CntX_Tipos_Cuentas where cod_contabilidad = " & vContabilidad & " order by Prioridad,Tipo_cuenta"
  Call OpenRecordSet(rs, strSQL)
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

On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Cuentas" Then

Select Case Right(Node.Key, 1)
        
    Case "T" 'Tipos de Cuentas
    
        strSQL = "select cod_cuenta,cod_cuenta_Mask, descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = ''" _
               & " and cod_contabilidad = " & vContabilidad _
               & " and TIPO_CUENTA = '" & fxIndiceCodigo(Node.Key) & "' order by cod_cuenta"
               
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, rs!Cod_Cuenta_Mask & " - " & rs!Descripcion, "imgFolder", IIf((rs!Acepta_movimientos = 0), True, False) _
                , "0x0" & fxgCntCuentaFormato(False, rs!cod_cuenta) & "C", rs!Acepta_movimientos)
          rs.MoveNext
        Loop
        rs.Close
    
    Case Else 'SubCuentas
    
        strSQL = "select cod_cuenta,cod_cuenta_Mask,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = '" & fxgCntCuentaFormato(False, fxIndiceCodigo(Node.Key)) _
               & "' and cod_contabilidad = " & vContabilidad & " order by cod_cuenta"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Call sbCreaNodos(Node.Key, rs!Cod_Cuenta_Mask & " - " & rs!Descripcion, "imgFolder", IIf((rs!Acepta_movimientos = 0), True, False) _
                , "0x0" & fxgCntCuentaFormato(False, rs!cod_cuenta) & "C", rs!Acepta_movimientos)
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
'    nodX.Image = vImagen
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

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  gCuenta = fxgCntCuentaFormato(False, Item.Text)
  UnLoad Me
End Sub


Private Sub txtFiltroCuenta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbLlenaLswCtas
End Sub

