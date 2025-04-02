VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox ckFechasTodas 
      Height          =   255
      Left            =   8400
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnImprime 
         Height          =   615
         Left            =   8640
         TabIndex        =   11
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmFSL_Reportes.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   3360
      Left            =   120
      TabIndex        =   0
      Top             =   1320
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   120
      Top             =   4440
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
            Picture         =   "frmFSL_Reportes.frx":0707
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":6F69
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":7087
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":71B1
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":72D7
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":73E5
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":74F2
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":760B
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":7739
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Reportes.frx":7846
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   360
      Left            =   5640
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   635
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   360
      Left            =   6960
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   635
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   5640
      TabIndex        =   13
      Top             =   3240
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   330
      Left            =   5640
      TabIndex        =   14
      Top             =   2640
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboOficina 
      Height          =   330
      Left            =   5640
      TabIndex        =   15
      Top             =   3840
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      Left            =   5640
      TabIndex        =   16
      Top             =   1440
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de FOSOL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   2
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   3840
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
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Reporte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
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
  MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub btnImprime_Click()
On Error GoTo vError
Dim strSQL As String

strSQL = Empty


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
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesOficinaResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesOficinaDetalle.rpt")
           End If
           
           .Formulas(0) = "fxTitulo= 'Expedientes por Oficina'"
                      
           If cboOficina.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "{FSL_EXPEDIENTES.REGISTRA_OFICINA}='" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
           End If
                                            
         Case "Expediente por Estado"
           
           .WindowTitle = "Reporte de los expedientes registrados por Estado"
           
           If cboTipo.Text = "Resumen" Then
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesEstadoResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesEstadoDetalle.rpt")
           End If

           .Formulas(0) = "fxTitulo= 'Expedientes por Estado'"
                      
           If cboEstado.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "({FSL_EXPEDIENTES.Estado}='" & SIFGlobal.fxCodText(cboEstado.Text) & "')"
           End If
           
         Case "Listado por Plan"
         
           .WindowTitle = "Reporte de los expedientes registrados por Plan"
         
           If cboTipo.Text = "Resumen" Then
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesPlanResumen.rpt")
           Else
             .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedientesPlanDetalle.rpt")
           End If
           
           .Formulas(0) = "fxTitulo= 'Expedientes según Plan de Aplicación'"
                      
           If cboPlan.Text <> "TODOS" Then
             If Len(strSQL) > 0 Then strSQL = strSQL & " and "
             strSQL = "({FSL_EXPEDIENTES.COD_PLAN}=" & cboPlan.ItemData(cboPlan.ListIndex) & ")"
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


Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
  vModulo = 7
End Sub

Private Sub Form_Load()
 vModulo = 7
  
  
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture
  
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


Private Sub sbCargaCombo()

  cboTipo.Clear
  cboTipo.AddItem "Detallado"
  cboTipo.AddItem "Resumen"
  cboTipo.Text = "Detallado"


    strSQL = "select rtrim(cod_oficina) as 'IdX',  rtrim(descripcion) as 'Itmx'" _
           & " from SIF_Oficinas order by cod_oficina"
    Call sbCbo_Llena_New(cboOficina, strSQL, True, True)
    
    
    strSQL = "select rtrim(cod_plan) as 'IdX' , rtrim(descripcion) as 'Itmx'" _
           & " from FSL_PLANES where ACTIVO = 1 order by COD_PLAN"
    Call sbCbo_Llena_New(cboPlan, strSQL, True, True)

    
    cboEstado.Clear
    cboEstado.AddItem ("TODOS")
    cboEstado.AddItem ("APR" & " - " & "APROBADO")
    cboEstado.AddItem ("APL" & " - " & "APELACION")
    cboEstado.AddItem ("PEN" & " - " & "PENDIENTE")
    cboEstado.AddItem ("REC" & " - " & "RECHAZADO")
    cboEstado.Text = "TODOS"
    
End Sub
