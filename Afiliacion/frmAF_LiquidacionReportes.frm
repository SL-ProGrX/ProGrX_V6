VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_LiquidacionReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Liquidaciones"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   HelpContextID   =   1014
   Icon            =   "frmAF_LiquidacionReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   7335
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   7575
      _Version        =   1441793
      _ExtentX        =   13361
      _ExtentY        =   2143
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   615
         Left            =   5400
         TabIndex        =   19
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
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
         Picture         =   "frmAF_LiquidacionReportes.frx":030A
      End
      Begin XtremeSuiteControls.CheckBox chkResumen 
         Height          =   615
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe Tipo Resumen?"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton cmdRepLiq 
      Height          =   495
      Left            =   6375
      TabIndex        =   1
      Top             =   360
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   868
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_LiquidacionReportes.frx":0AC6
   End
   Begin XtremeSuiteControls.FlatEdit txtLiq 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkReversadas 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Excluir Liquidaciones Reversadas"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.RadioButton optFechas 
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.RadioButton optFechas 
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   3960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Rango"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   1440
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboFiltro 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboDesembolso 
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   2160
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboInstituciones 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   2520
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.CheckBox chkUnidadProgramatica 
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   3240
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9546
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Informe por Departamento / Unidad"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   1440
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Filtros"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Empresa"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Inicio"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Liquidación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   1880
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmAF_LiquidacionReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnInforme_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite reporte de devoluciones de Aportes.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, vUnidad As String

Me.MousePointer = vbHourglass

vUnidad = ""

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Personas"
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

    .Connect = glogon.ConectRPT
    
    Select Case Mid(cbo.Text, 1, 2)
       Case "00" 'Informativo Sin Montos
            If chkResumen.Value = vbChecked Then
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTotales.rpt")
              .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
            
            Else
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionesInfo.rpt")
              .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
            End If
            
       Case "01" 'Listado de Liquidaciones
            If chkResumen.Value = vbChecked Then
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTotales.rpt")
              .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
            
            Else
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_DetalleLiquidaciones.rpt")
              .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
            End If
       
       Case "02" 'Listado x Causa de Renuncia
            If chkResumen.Value = vbChecked Then
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionXCausasRsm.rpt")
            Else
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionXCausasDet.rpt")
            End If
    
       Case "03" 'Listado Especial
            If chkResumen.Value = vbChecked Then
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionEspecialDet.rpt")
            Else
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionEspecialDet.rpt")
            End If
    
       Case "04" 'Liquidaciones con Saldos
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionSaldos.rpt")
              strSQL = "({LIQUIDA_DETALLE.LIQ_SALDO} - {LIQUIDA_DETALLE.LIQ_AMORTIZA} <> 0) AND {CATALOGO.RETENCION} = 'N'"
       
       Case "05" 'Liquidaciones x Rubros
              .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionRubros.rpt")
    End Select
    

    If optFechas(1).Value Then
        .Formulas(2) = "SubTitulo='Del  " & Format(dtpInicio, "dd/mm/yyyy") & "  Al  " _
              & Format(dtpCorte, "dd/mm/yyyy") & " - FILTRO : " & UCase(cboFiltro.Text) & "'"
        
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.FECLIQ} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    Else
       .Formulas(2) = "SubTitulo='HISTORICO - FILTRO : " & UCase(cboFiltro.Text) & "'"
    End If
    
  If chkReversadas.Value = vbChecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{LIQUIDACION.ESTADO} = 'P'"
  End If
    
  .Formulas(2) = .Formulas(2) & "' [Inst: " & cboInstituciones.Text & "]'"
  If cboInstituciones.Text <> "TODOS" Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstituciones.ItemData(cboInstituciones.ListIndex)
  End If
    
  Select Case Mid(cboFiltro.Text, 1, 2)
    Case "01" 'Socio -> Ex.Int
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.ESTADOACTUAL} = 'S' AND {LIQUIDACION.ESTADOACTLIQ} = 'A'"
    Case "02" 'Socio -> Ex.Empleado
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.ESTADOACTUAL} = 'S' AND {LIQUIDACION.ESTADOACTLIQ} = 'P'"
    Case "03" 'Ex.Int -> Ex.Empleado
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.ESTADOACTUAL} = 'A' AND {LIQUIDACION.ESTADOACTLIQ} = 'P'"
    Case "00" 'Nada Todos
  End Select
  
  
  Select Case Mid(cboDesembolso.Text, 1, 2)
    Case "01" 'Con Desembolso
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.UBICACION} = 'T'"
    Case "02" 'Sin Desembolso
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{LIQUIDACION.UBICACION} = 'C'"
    Case "00" 'Nada Todos
  End Select
  
    
  If chkUnidadProgramatica.Value = vbChecked Then
    vUnidad = InputBox("Digite la Unidad Programatica a Consultar : ", "Listado Liquidaciones x Unidad")
    If Len(Trim(vUnidad)) > 0 Then
        strSQL = strSQL & " AND {SOCIOS.UP} = '" & vUnidad & "'"
        .Formulas(2) = Mid(.Formulas(2), 1, Len(.Formulas(2)) - 1) & " UP : " & vUnidad & "'"
    End If
  End If
  
  .SelectionFormula = strSQL
  .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub cmdRepLiq_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If txtLiq.Text = "" Or Not IsNumeric(txtLiq.Text) Then Exit Sub

Call sbgAFIBoletaLiquidacion(txtLiq.Text)

End Sub

Private Sub sbInicializaCbo()
Dim strSQL As String

cbo.AddItem "00 - Liquidaciones Informativo"
cbo.AddItem "01 - Listados Liquidaciones"
cbo.AddItem "02 - Liquidaciones x Causas"
cbo.AddItem "03 - Listado especial Recursos"
cbo.AddItem "04 - Listado Liq. con Saldos"
cbo.AddItem "05 - Listado Liq. x Rubros"

cbo.Text = "00 - Liquidaciones Informativo"

cboFiltro.AddItem "01 - Asociado a Ex-Interno"
cboFiltro.AddItem "02 - Asociado a Ex-Empleado"
cboFiltro.AddItem "03 - Ex-Interno a Ex-Empleado"
cboFiltro.AddItem "00 - Todos"

cboFiltro.Text = "00 - Todos"

cboDesembolso.AddItem "00 - Todos"
cboDesembolso.AddItem "01 - Con Desembolso"
cboDesembolso.AddItem "02 - Sin Desembolso"
cboDesembolso.Text = "00 - Todos"

strSQL = "select cod_institucion as Idx, descripcion as ItmX from instituciones"
Call sbCbo_Llena_New(cboInstituciones, strSQL, True, True)

dtpInicio.Enabled = False
dtpCorte.Enabled = dtpInicio.Enabled

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value



End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call sbInicializaCbo

End Sub

Private Sub optFechas_Click(Index As Integer)

If Index = 0 Then
 dtpInicio.Enabled = False
Else
 dtpInicio.Enabled = True
End If
dtpCorte.Enabled = dtpInicio.Enabled


End Sub
