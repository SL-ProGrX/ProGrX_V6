VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.Controls.v18.6.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.ShortcutBar.v18.6.0.ocx"
Begin VB.Form frmTES_Explorer 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3456
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11484
   HelpContextID   =   1006
   Icon            =   "frmTES_Explorer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3456
   ScaleWidth      =   11484
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2172
      Left            =   2760
      TabIndex        =   8
      Top             =   720
      Width           =   3012
      _Version        =   1179654
      _ExtentX        =   5313
      _ExtentY        =   3831
      _StockProps     =   77
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.PushButton btnRefrescar 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1812
      _Version        =   1179654
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Refrescar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   7
      Picture         =   "frmTES_Explorer.frx":030A
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2160
      Left            =   5400
      ScaleHeight     =   940.557
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   6480
      Top             =   360
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":0A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":12E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":1BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":1EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":21FA
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":2516
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":2832
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":2B4E
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":342A
            Key             =   "imgGrupo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":3D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":45E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":4EBE
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":579A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Explorer.frx":5ABE
            Key             =   "imgNotas"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   2610
      _ExtentX        =   4614
      _ExtentY        =   3810
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   2052
      _Version        =   1179654
      _ExtentX        =   3620
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Width           =   4572
      _Version        =   1179654
      _ExtentX        =   8065
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   1332
      _Version        =   1179654
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Width           =   1332
      _Version        =   1179654
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   336
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   360
      Width           =   3012
      _Version        =   1179654
      _ExtentX        =   5313
      _ExtentY        =   593
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   336
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   2652
      _Version        =   1179654
      _ExtentX        =   4678
      _ExtentY        =   593
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   2565
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
End
Attribute VB_Name = "frmTES_Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vNode As Node
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub sbCargaLsw(vTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, dbMonto As Double


Me.MousePointer = vbHourglass

strSQL = "select C.nsolicitud,C.ndocumento,C.tipo,C.codigo,C.beneficiario,C.monto,C.fecha_solicitud" _
       & ",C.fecha_anula,C.fecha_emision,C.fecha_autorizacion,B.descripcion" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_banco" _
       & " where C.tipo = '" & vTipo & "' and C.id_banco = " & cbo.ItemData(cbo.ListIndex)


Select Case cboTipo.Text
 Case "Solicitados"
    strSQL = strSQL & " and C.FECHA_SOLICITUD between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and C.estado in('P')"
           
 Case "Emitidos"
    
    strSQL = strSQL & " and C.FECHA_EMISION between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'  and C.estado in('I','T')"
    
 Case "Anulados"
    
    strSQL = strSQL & " and C.FECHA_ANULA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "' and C.estado in('A')"
    
 Case "Autorizados"
    strSQL = strSQL & " and C.FECHA_AUTORIZACION between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and C.estado in('P')"

End Select

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Tipo", 540
lsw.ColumnHeaders.Add , , "Solicitud", 1240
lsw.ColumnHeaders.Add , , "Documento", 1840
lsw.ColumnHeaders.Add , , "Código", 1540
lsw.ColumnHeaders.Add , , "Beneficiario", 3540
lsw.ColumnHeaders.Add , , "Monto", 1840, vbRightJustify
lsw.ColumnHeaders.Add , , "Banco", 3140
lsw.ColumnHeaders.Add , , "Solicitud", 2140, vbCenter
lsw.ColumnHeaders.Add , , "Emision", 2140, vbCenter
lsw.ColumnHeaders.Add , , "Anulacion", 2140, vbCenter
lsw.ColumnHeaders.Add , , "Autorizacion", 2140, vbCenter
    
Call OpenRecordSet(rs, strSQL, 0)
dbMonto = 0
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Tipo)
     itmX.SubItems(1) = rs!NSolicitud
     itmX.SubItems(2) = rs!nDocumento & ""
     itmX.SubItems(3) = rs!Codigo
     itmX.SubItems(4) = rs!Beneficiario
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = rs!Descripcion
     itmX.SubItems(7) = Format(rs!fecha_solicitud, "dd/mm/yyyy")
     itmX.SubItems(8) = Format(rs!Fecha_Emision, "dd/mm/yyyy")
     itmX.SubItems(9) = Format(rs!Fecha_Anula, "dd/mm/yyyy")
     itmX.SubItems(10) = Format(rs!fecha_autorizacion, "dd/mm/yyyy")
     
     dbMonto = dbMonto + rs!Monto
 rs.MoveNext
Loop
rs.Close

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = "TOTAL"
    itmX.SubItems(5) = "____________"

Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(1) = ""
    itmX.SubItems(2) = ""
    itmX.SubItems(5) = Format(dbMonto, "Standard")
    
Me.MousePointer = vbDefault


End Sub


Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo vError

Set vNode = Node

lblTitle(0).Caption = vNode.FullPath
lblTitle(1).Caption = UCase(vNode.Parent) & " : " & UCase(vNode.Text)

Select Case Node.Text
  Case "Transferencias"
     Call sbCargaLsw("TE")
  Case "Cheques"
     Call sbCargaLsw("CK")
  Case "Depositos"
     Call sbCargaLsw("DP")
  Case "Notas de Debito"
     Call sbCargaLsw("ND")
  Case "Notas de Credito"
     Call sbCargaLsw("NC")
End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault

End Sub


Private Sub btnRefrescar_Click()
 Call sbRefrescaArbol
End Sub

Private Sub cboTipo_Click()
 Call sbRefrescaArbol
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 9

cboTipo.Clear
cboTipo.AddItem "Solicitados"
cboTipo.AddItem "Emitidos"
cboTipo.AddItem "Anulados"
cboTipo.AddItem "Autorizados"



strSQL = "select id_banco as 'Idx',rtrim(descripcion) as 'ItmX' from Tes_Bancos where estado = 'A'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

 Call sbRefrescaArbol

 
 dtpCorte.Value = fxFechaServidor
 dtpInicio = dtpCorte.Value
 cboTipo.Text = "Solicitados"
 
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If UnloadMode = 0 Then
'      Cancel = True
'      Me.WindowState = 1
'   End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
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


Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    ArbolExp.Width = x
    imgSplitter.Left = x
    lsw.Left = x + 40
    lsw.Width = Me.Width - (ArbolExp.Width + 140)
    lblTitle(0).Width = ArbolExp.Width
    lblTitle(1).Left = lsw.Left + 20
    lblTitle(1).Width = lsw.Width - 40


    'set the top
    lsw.Top = ArbolExp.Top
    imgSplitter.Top = ArbolExp.Top
    ArbolExp.Height = Me.Height - 1100

    lsw.Height = ArbolExp.Height
    imgSplitter.Height = ArbolExp.Height
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim itmX As ListViewItem

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Tesoreria", "Tesoreria", "imgRoot")
  'Crear Arbol Inicial
  Call sbCreaNodos(vNode.Key, "Cheques", "imgDetalle", False)
  Call sbCreaNodos(vNode.Key, "Transferencias", "imgNotas", False)
  Call sbCreaNodos(vNode.Key, "Notas de Credito", "imgOpcion", False)
  Call sbCreaNodos(vNode.Key, "Notas de Debito", "imgOpcion", False)
  Call sbCreaNodos(vNode.Key, "Depositos", "imgNotas", False)
  
  .Nodes(1).Expanded = True
  
     With lsw
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Empresa", 4450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
    
       Set itmX = lsw.ListItems.Add(, , GLOBALES.gstrNombreEmpresa)
           itmX.SubItems(3) = glogon.Usuario
           itmX.SubItems(4) = Format(fxFechaServidor, "dd/mm/yyyy")
     End With
  
End With

End Sub

Function fxIndice(str As String) As String
Dim nodX As Node, lng As Long
On Error Resume Next
With ArbolExp
  For lng = 2 To .Nodes.Count
    Set nodX = .Nodes.Item("0x0" & lng)
    If nodX.Text = str Then
     fxIndice = "0x0" & lng
     Exit Function
    End If
  Next lng
End With
fxIndice = "0"
End Function


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, Optional vExpande As Boolean)
Dim nodX As Node, vKey As String

On Error GoTo vError

vKey = "0x0" & ArbolExp.Nodes.Count + 1
    Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
        nodX.Key = vKey
        nodX.Text = vTexto
        nodX.Tag = nodX.Index
        nodX.Image = vImagen
                
If vExpande = True Then
    vKey = "0x000" & ArbolExp.Nodes.Count + 1
        Set nodX = ArbolExp.Nodes.Add(nodX.Key, tvwChild)
            nodX.Key = vKey
            nodX.Tag = nodX.Index
    Exit Sub
End If

vError:
End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

