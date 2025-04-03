VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFSL_Reportes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPlan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CheckBox ckFechasTodas 
      Caption         =   "Todas"
      Height          =   255
      Left            =   7920
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cboOficina 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2520
      Width           =   3735
   End
   Begin VB.ComboBox cboEstado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2160
      Width           =   3735
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   3360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5927
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   169213955
      CurrentDate     =   36278
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   169213955
      CurrentDate     =   36278
   End
   Begin MSComctlLib.Toolbar tlbImprime 
      Height          =   330
      Left            =   8280
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      ButtonWidth     =   1799
      ButtonHeight    =   582
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
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   120
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6862
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6980
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6AAA
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6BD0
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6CDE
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6DEB
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6F04
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":7032
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":713F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Plan "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblReporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   4560
      TabIndex        =   12
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Estados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes FOSOL"
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
      TabIndex        =   8
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
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
      Index           =   5
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
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
      Index           =   6
      Left            =   5640
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmFSL_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As ADODB.Recordset

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
       Case "Expediente por Oficina"
         cboTipo.Clear
         cboTipo.AddItem "Detalle"
         cboTipo.AddItem "Resumen"
         cboTipo.Text = "Detalle"
       
       Case "Expediente por Estado"
         cboTipo.Clear
         cboTipo.AddItem "Detalle"
         cboTipo.AddItem "Resumen"
         cboTipo.Text = "Detalle"
         
       Case "Listado por Plan"
         cboTipo.Clear
         cboTipo.AddItem "Detalle"
         cboTipo.AddItem "Resumen"
         cboTipo.Text = "Detalle"

     End Select
  End If

Exit Sub

vError:
  MsgBox Err.Description

End Sub

Private Sub ckFechasTodas_Click()
   If ckFechasTodas.Value = vbChecked Then
      dtpInicio.Enabled = False
      dtpCorte.Enabled = False
   Else
      dtpInicio.Enabled = True
      dtpCorte.Enabled = True
   End If
End Sub

Private Sub Form_Activate()
  vModulo = 3
End Sub

Private Sub Form_Load()
 vModulo = 3
  
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value
 Call sbRefrescaArbol
 Call sbCargaCombo
End Sub

Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String
Dim vPadre As String

With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  
  Call sbCreaNodos("Reportes", "Expedientes", "imgExp", False, "0x0OEx")
     Call sbCreaNodos("0x0OEx", "Expediente por Oficina", "imgDetalle", False, "0x0" & "OFI" & "Z")
     Call sbCreaNodos("0x0OEx", "Expediente por Estado", "imgDetalle", False, "0x0" & "EST" & "Z")
     Call sbCreaNodos("0x0OEx", "Listado por Plan", "imgDetalle", False, "0x0" & "PLN" & "Z")

  .Nodes(1).Expanded = True
End With


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

Private Sub tlbImprime_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
Dim strSQL As String

strSQL = Empty

Select Case UCase(Button.Key)

Case "REPORTE"
    With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .Connect = "pwd =" & glogon.RootKey
    
    If lblReporte.Caption = Empty Then
       MsgBox "Debe de seleccionar un reporte"
       Exit Sub
    Else
       
     Select Case lblReporte.Caption
     
         Case "Expediente por Oficina"
           .WindowTitle = "Reporte de los expedientes registrados por Oficina"
           
           If cboTipo.Text = "Resumen" Then
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesOficinaResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesOficinaDetalle.rpt")
           End If
           
           .Formulas(0) = "fxTitulo= 'Expedientes por Oficina'"
                      
           If cboOficina.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "{FSL_EXPEDIENTES.REGISTRA_OFICINA}='" & SIFGlobal.fxSIFCodText(cboOficina.Text) & "'"
           End If
                                            
         Case "Expediente por Estado"
           
           .WindowTitle = "Reporte de los expedientes registrados por Estado"
           
           If cboTipo.Text = "Resumen" Then
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesEstadoResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesEstadoDetalle.rpt")
           End If

           .Formulas(0) = "fxTitulo= 'Expedientes por Estado'"
                      
           If cboEstado.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "({FSL_EXPEDIENTES.Estado}='" & SIFGlobal.fxSIFCodText(cboEstado.Text) & "')"
           End If
           
         Case "Listado por Plan"
         
           .WindowTitle = "Reporte de los expedientes registrados por Plan"
         
           If cboTipo.Text = "Resumen" Then
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesPlanResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedientesPlanDetalle.rpt")
           End If
           
           .Formulas(0) = "fxTitulo= 'Expedientes según Plan de Aplicación'"
                      
           If cboPlan.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "({FSL_EXPEDIENTES.COD_PLAN}=" & SIFGlobal.fxSIFCodText(cboPlan.Text) & ")"
           End If
          
                   
    End Select 'Fin de Case lblReporte.Caption
    
    
     If ckFechasTodas.Value = vbUnchecked Then
        If Len(strSQL) > 0 Then strSQL = strSQL & " and "
        strSQL = strSQL & "cdate({FSL_EXPEDIENTES.REGISTRO_FECHA}) in Datetime(" & Format(dtpInicio.Value, "yyyy,mm,dd,00,00,00") _
               & ") to Datetime(" & Format(dtpCorte.Value, "yyyy,mm,dd,23,59,59") & ")"
               
     End If
     
     .SelectionFormula = strSQL
     
     .Formulas(1) = "fxSubTitulo='Desde  " & Format(dtpInicio.Value, "dd/mm/yyyy") & "  Hasta  " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
     .Formulas(2) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(4) = "fxUsuario='Usuario: " & glogon.Usuario & "'"
     
     '.Action = 1
     .PrintReport
     
   End If
   
   End With
End Select

Exit Sub

vError:
      MsgBox Err.Description, vbCritical
End Sub

Private Sub sbCargaCombo()

    strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as Itmx" _
           & " from SIF_Oficinas order by cod_oficina"
    Call sbLlenaCbo(cboOficina, strSQL, True, False)
    
    
    strSQL = "select rtrim(cod_plan) + ' - ' + rtrim(descripcion) as Itmx" _
           & " from FSL_PLANES_APLICACION where ACTIVO = 1 order by COD_PLAN"
    Call sbLlenaCbo(cboPlan, strSQL, True, False)

    
    cboEstado.Clear
    cboEstado.AddItem ("TODOS")
    cboEstado.AddItem ("APR" & " - " & "APROBADO")
    cboEstado.AddItem ("APL" & " - " & "APELACION")
    cboEstado.AddItem ("PEN" & " - " & "PENDIENTE")
    cboEstado.AddItem ("REC" & " - " & "RECHAZADO")
    cboEstado.Text = "TODOS"
    
End Sub
