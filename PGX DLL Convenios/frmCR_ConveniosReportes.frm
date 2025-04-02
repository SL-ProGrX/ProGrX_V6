VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_ConveniosReportes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes"
   ClientHeight    =   6540
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9432
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLineas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5400
      TabIndex        =   22
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox cboUsuarios 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CheckBox chkFechas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5400
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cboConvenios 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox cboEOperacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   240
      Top             =   5640
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6862
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6980
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6AAA
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6BD0
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6CDE
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6DEB
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":6F04
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":7032
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConveniosReportes.frx":713F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   312
      Left            =   8040
      TabIndex        =   0
      Top             =   5976
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   550
      ButtonWidth     =   1693
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgArbol"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   5760
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   4215
      _ExtentX        =   7430
      _ExtentY        =   10160
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   191561731
      CurrentDate     =   36278
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   191561731
      CurrentDate     =   36278
   End
   Begin VB.Label Label1 
      Caption         =   "Usuarios"
      Height          =   255
      Index           =   14
      Left            =   4320
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5400
      TabIndex        =   17
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Convenio"
      Height          =   255
      Index           =   17
      Left            =   4320
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Convenios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblReporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Reporte"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fechas"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   7320
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Estados"
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solicitud"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmCR_ConveniosReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkLineas_Click()
  If chkLineas.Value = vbChecked Then
    txtCodigo.Enabled = False
  Else
    txtCodigo.Enabled = True
  End If
End Sub

Private Sub Form_Activate()
  vModulo = 16
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 16

Call sbRefrescaArbol

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"
txtCodigo.Enabled = True
  
cboUsuarios.Clear
strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
       & " from  crd_grupos"
Call sbLlenaCbo(cboUsuarios, strSQL)

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
  
End Sub

Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

If Right(Node.Key, 1) = "Z" Then
  lblReporte.Caption = Node.Text
  lblReporte.Tag = fxIndiceCodigo(Node.Key)
  
  Select Case Node.Text
     Case "Reporte General"
       cboEOperacion.Visible = True
         
  End Select
  
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
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
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub

Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPadre As String

With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  Call sbCreaNodos("Reportes", "Convenios", "imgCV", False, "0x0OPR")
     Call sbCreaNodos("0x0OPR", "Reporte General", "imgDetalle", False, "0x0" & "GEN" & "Z")

  .Nodes(1).Expanded = True
End With

End Sub

Private Function fxReporteFile() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select prefijo from crd_reportes where id = " & lblReporte.Tag
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxReporteFile = ""
Else
  fxReporteFile = Trim(rs!prefijo)
End If
rs.Close

End Function

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

On Error GoTo vError


Select Case UCase(Button.Key)
Case "REPORTE"
         
    
    With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .Connect = glogon.ConectRPT
    
    If lblReporte.Caption = Empty Then
       MsgBox "Debe de seleccionar un reporte"
       Exit Sub
    Else
           Select Case lblReporte.Caption
                'Operaciones
                 Case "Reporte General"
                   .WindowTitle = "Reporte General de Convenios"
                   .ReportFileName = SIFGlobal.fxPathReportes("APAOperacionesGeneral.rpt")
                   .Formulas(0) = "fxTitulo= 'Reporte General'"
'                 Case "Reporte x Acreedor"
'                   If cboAcreedores.Text = Empty Then Exit Sub
'
'                   .WindowTitle = "Operaciones Según Acreedor"
'                   .ReportFileName = SIFGlobal.fxPathReportes("APAOperacionesAcreedor.rpt")
'                   .Formulas(0) = "fxTitulo='Operaciones según Acreedor'"
'                   .SelectionFormula = " {CRD_APA_OPERACIONES.COD_ACREEDOR}=" & "'" & DeCodificaPrimaryKey(cboAcreedores.SelectedItem.Key, 1, "(id)") & "'"
'                   .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.ESTADO}=" & "'" & Mid(cboEstado, 1, 1) & "'"
                    
'                 'Reporte por saldos
'                 Case "Reporte x Saldo"
'                    .WindowTitle = "Reporte x Saldo"
'                    .ReportFileName = SIFGlobal.fxPathReportes("APAOperacionesSaldos.rpt")
'
'                 If cboCondicionSaldo.Text <> "Todos" Then
'                    Select Case cboCondicionSaldo.Text
'                         Case "Mayor que"
'                            .Formulas(0) = "fxTitulo =' Operaciones con saldo mayor a  " & Format(txtSaldoMayor.Text, "Currency") & "'"
'                            .SelectionFormula = "{CRD_APA_OPERACIONES.SALDO} > " & txtSaldoMayor.Text & ""
'                         Case "Menor que"
'                            .Formulas(0) = "fxTitulo =' Operaciones con saldo menor a  " & Format(txtSaldoMayor.Text, "Currency") & "'"
'                            .SelectionFormula = "{CRD_APA_OPERACIONES.SALDO} < " & txtSaldoMayor.Text & ""
'                         Case "Entre"
'                            .Formulas(0) = "fxTitulo =' Operaciones con saldo entre  " & Format(txtSaldoMayor.Text, "Currency") & "  y " & Format(txtSaldoMenor.Text, "Currency") & " '"
'                            .SelectionFormula = "{CRD_APA_OPERACIONES.SALDO} <= " & txtSaldoMayor.Text & ""
'                            .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.SALDO} >= " & txtSaldoMenor.Text & ""
'                         Case "Igual a"
'                            .Formulas(0) = "fxTitulo =' Operaciones con saldo igual a  " & Format(txtSaldoMayor.Text, "Currency") & "  '"
'                            .SelectionFormula = "{CRD_APA_OPERACIONES.SALDO} = " & txtSaldoMayor.Text & ""
'                    End Select
'
'                    If cboAcreedores.Text <> Empty Then
'                      .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.COD_ACREEDOR}=" & "'" & DeCodificaPrimaryKey(cboAcreedores.SelectedItem.Key, 1, "(id)") & "'"
'                      .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.ESTADO}=" & "'" & Mid(cboEstado, 1, 1) & "'"
'                    End If
'                 Else
'                  If cboAcreedores.Text <> Empty Then
'                     .SelectionFormula = " {CRD_APA_OPERACIONES.COD_ACREEDOR}=" & "'" & DeCodificaPrimaryKey(cboAcreedores.SelectedItem.Key, 1, "(id)") & "'"
'                     .SelectionFormula = .SelectionFormula & " and {CRD_APA_OPERACIONES.ESTADO}=" & "'" & Mid(cboEstado, 1, 1) & "'"
'                  End If
'                 End If
                 

                      
           End Select
            
             .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
             .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
             .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
             
             .Action = 1
             '.PrintReport
             dtpInicio.Enabled = True
             dtpCorte.Enabled = True
        End If
       End With
    End Select
    
Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  lblDescripcion.Caption = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
End Sub
