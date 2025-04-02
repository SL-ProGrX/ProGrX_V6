VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCC_PlanillaReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informes de deducciones de Planillas"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Left            =   240
      TabIndex        =   26
      Top             =   6120
      Width           =   8532
      _Version        =   1310722
      _ExtentX        =   15049
      _ExtentY        =   1926
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   660
         Index           =   0
         Left            =   6720
         TabIndex        =   27
         Top             =   240
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   1164
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
         Appearance      =   16
         Picture         =   "frmCC_PlanillaReportes.frx":0000
      End
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   2520
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Generación de Deducciones"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkInstituciones 
      Height          =   372
      Left            =   4560
      TabIndex        =   9
      Top             =   3480
      Width           =   3972
      _Version        =   1310722
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Rastrear todas las deductoras   "
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
      TextAlignment   =   1
      Alignment       =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   6852
      _Version        =   1310722
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.CheckBox chkLineas 
      Height          =   372
      Left            =   4560
      TabIndex        =   10
      Top             =   3840
      Width           =   3972
      _Version        =   1310722
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Rastrear todas las Líneas de Crédito   "
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
      TextAlignment   =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   312
      Left            =   5640
      TabIndex        =   11
      Top             =   2520
      Width           =   2892
      _Version        =   1310722
      _ExtentX        =   5106
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
   Begin XtremeSuiteControls.ComboBox cboLineas 
      Height          =   312
      Left            =   5640
      TabIndex        =   12
      Top             =   4560
      Width           =   2892
      _Version        =   1310722
      _ExtentX        =   5106
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipos 
      Height          =   312
      Left            =   5640
      TabIndex        =   13
      Top             =   4920
      Width           =   2892
      _Version        =   1310722
      _ExtentX        =   5106
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   5640
      TabIndex        =   14
      Top             =   2880
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
      _ExtentY        =   582
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
      Height          =   330
      Left            =   7080
      TabIndex        =   15
      Top             =   2880
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   312
      Left            =   1680
      TabIndex        =   16
      Top             =   1680
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   2880
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Carga Deducciones"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   2
      Left            =   600
      TabIndex        =   19
      Top             =   3240
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Desglose"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   3
      Left            =   600
      TabIndex        =   20
      Top             =   3720
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aplicación de Patrimonio"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   4080
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aplicación de Abonos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   5
      Left            =   600
      TabIndex        =   22
      Top             =   4560
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Envío a Fondos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   6
      Left            =   600
      TabIndex        =   23
      Top             =   4920
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Bitácora de Aplicaciones"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   7
      Left            =   600
      TabIndex        =   24
      Top             =   5280
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Análisis de Efectividad de Cobros "
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   8
      Left            =   600
      TabIndex        =   25
      Top             =   5640
      Width           =   3492
      _Version        =   1310722
      _ExtentX        =   6159
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Especial Cliente Corporativo"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin VB.Label lblEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Caso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Label lblEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea Deduc."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Width           =   1212
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   4800
      X2              =   8400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "   Salida"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   360
      X2              =   3960
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "   Reportes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Deducciones de Planillas"
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
      Height          =   612
      Index           =   0
      Left            =   1884
      TabIndex        =   3
      Top             =   360
      Width           =   6852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmCC_PlanillaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vScroll As Boolean
Dim mFrecuencPago As String


Private Sub btnReporte_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

Dim pProcesoFormat As String

On Error GoTo vError

Me.MousePointer = vbHourglass

pProcesoFormat = fxFechaProcesoFormat(txtProceso.Text)

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Módulo de Deducciones"
     
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "Usuario = '" & glogon.Usuario & "'"
    .Formulas(3) = "Institucion = '" & cboInstitucion.Text & "'"
    .Formulas(4) = "Fecha='" & pProcesoFormat & "'"
    
'    .Formulas(2) = "SubTitulo = 'Fecha Proceso : " & Format(txtProceso.Text, "####-##") & "'"
    
  
  Select Case True
    Case optX.Item(0).Value 'Genera Deducciones
        Select Case cboReporte.Text
          Case "Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Generada.rpt")
            .SelectionFormula = "{PRM_PLANILLA.PROCESO} = " & txtProceso.Text _
                              & " AND {PRM_PLANILLA.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
          
          Case "Detalle"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_GeneradaDetalle.rpt")
            .SelectionFormula = "{PRM_PLANILLA.PROCESO} = " & txtProceso.Text _
                              & " AND {PRM_PLANILLA.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
        
          Case "Línea"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_GeneradaxLinea.rpt")
            .SelectionFormula = "{CUOTAS_ENVIADAS.FECPRO} = " & txtProceso.Text _
                              & " AND {CUOTAS_ENVIADAS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
          
          Case "Línea Detalle"
        
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_GeneradaxLineaDetalle.rpt")
            .SelectionFormula = "{CUOTAS_ENVIADAS.FECPRO} = " & txtProceso.Text _
                              & " AND {CUOTAS_ENVIADAS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
        
          Case "Base Actual"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Generada_BaseActual.rpt")
            .StoredProcParam(0) = cboInstitucion.ItemData(cboInstitucion.ListIndex)
            .StoredProcParam(1) = txtProceso.Text
            
        End Select
    
    Case optX.Item(1).Value 'Carga Deducciones
        Select Case cboReporte.Text
          Case "Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Cargada.rpt")
            .SelectionFormula = "{PRM_CARGADO.FECHA_PROCESO} = " & txtProceso.Text _
                      & " AND {PRM_CARGADO.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
          Case "Detalle"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CargadaDetalle.rpt")
            .SelectionFormula = "{PRM_CARGADO.FECHA_PROCESO} = " & txtProceso.Text _
                      & " AND {PRM_CARGADO.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
        
          Case "No Localizados"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Cargada_NoLocalizado.rpt")
            .SelectionFormula = "{vPrmCargadoPersonasNoEncontradas.FECHA_PROCESO} = " & txtProceso.Text _
                   & " AND {vPrmCargadoPersonasNoEncontradas.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
        
        End Select
    
    Case optX.Item(2).Value 'Desgloce
        Select Case cboReporte.Text
          Case "Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdCarga.rpt")
          Case "Detalle"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdCargaDetalle.rpt")
          Case "Agrupado: Línea"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdCargaDetalleAgrupado.rpt")
          Case "Agrupado: Persona"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdCargaDetalleAgrupadoPersona.rpt")
    
         End Select
        
            .SelectionFormula = "{PRM_CREDITOS.FECHA_PROCESO} = " & txtProceso.Text _
                              & " AND {PRM_CREDITOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
                              
    Case optX.Item(3).Value 'Aplicacion a Patrimonio
            
            strSQL = "Select porc_aporte,porc_ahorro from instituciones where cod_institucion = " & GLOBALES.gInstitucion
            Call OpenRecordSet(rs, strSQL)
                .Formulas(6) = "Porcentaje=" & IIf(IsNull(rs!PORC_APORTE), 0, rs!PORC_APORTE) / 100
                .Formulas(7) = "PorcAhorro=" & IIf(IsNull(rs!porc_ahorro), 0, rs!porc_ahorro) / 100
            rs.Close
            
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_PatAplicados.rpt")
            .SelectionFormula = "{SOCIOSTEMP.EXISTE} = 'S' AND {SOCIOSTEMP.FECHAPROC} = " & txtProceso.Text _
                              & " AND {SOCIOSTEMP.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    
    Case optX.Item(4).Value 'Aplicacion a Creditos
            .Formulas(6) = "Titulo = 'ABONOS APLICADOS'"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdAbonosAplicados.rpt")
            .SelectionFormula = "{APLICACIONCR.APL_FECHAP}=" & txtProceso.Text
            
            
    Case optX.Item(5).Value 'Envio a Fondos
            .Formulas(6) = "Titulo='PLANILLAS'"
            .Formulas(7) = "SubTitulo='ENVIO AL FONDO DE INCONSISTENCIAS              PROCESO : " & Format(txtProceso.Text, "####-##") & "'"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Fondo.rpt")
            .SelectionFormula = "{PRM_FONDO.PROCESO} = " & txtProceso.Text & " AND {PRM_FONDO.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  
  
    Case optX.Item(6).Value 'Bitacora
        
        Select Case cboReporte.Text
            Case "Fechas"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_BitacoraRsm.rpt")
            Case "Proceso + Fechas"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_BitacoraRsmProceso.rpt")
            Case "Proceso + Institución"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_BitacoraRsmProcesoInst.rpt")
          End Select
        
        .Formulas(6) = "Titulo='PLANILLAS'"
        .Formulas(7) = "SubTitulo='BITACORA DE APLICACIONES              FECHAS : " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                & "-" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{vSIFPlanillaBitacoraRsm.FECHA} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
  
    Case optX.Item(7).Value 'Analisis de Efectividad de Cobro x Planillas
'        If chkInstituciones.Value = vbChecked Then 'Todas las instituciones
'              Select Case cboReporte.Text
'                Case "Resumen"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasRsm.rpt")
'                Case "Detalle"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasDet.rpt")
'                Case "Estadística"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasEst.rpt")
'                Case "Tipo"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasTipo.rpt")
'                Case "Línea"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasLinea.rpt")
'                Case "Línea Resumen"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasLineaRsm.rpt")
'
'
'                Case "Persona"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasPersona.rpt")
'                Case "Persona Resumen"
'                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadInstTodasPersonaRsm.rpt")
'
'               End Select
'
'             .StoredProcParam(0) = txtProceso.Text
'
'        Else 'Por Institucion
'
'
'
'       End If
       
       
            Select Case cboReporte.Text
              Case "Resumen"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadRsm.rpt")
              
              Case "Detalle"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadDet.rpt")

              Case "Estadística"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadEst.rpt")

              Case "Tipo"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadTipo.rpt")

              
              Case "Línea"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadLinea.rpt")
              Case "Línea Resumen"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadLineaRsm.rpt")
             
              Case "Persona"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadPersonaDet.rpt")
              Case "Persona Resumen"
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_EfectividadPersonaRsm.rpt")
             
             End Select
            
            If chkInstituciones.Value = vbChecked Then
                .StoredProcParam(0) = 0
            Else
                .StoredProcParam(0) = cboInstitucion.ItemData(cboInstitucion.ListIndex)
            End If
             .StoredProcParam(1) = txtProceso.Text
       
       'Finalmente los paremetros FILTROS
         If chkLineas.Value = vbUnchecked Then
          .SelectionFormula = "{spSIFPlanillaCompara;1.Linea} = '" & cboLineas.ItemData(cboLineas.ListIndex) & "'"
         End If
    
         If cboTipos.Text <> "Todos" Then
            If Len(.SelectionFormula) > 0 Then .SelectionFormula = .SelectionFormula & " AND "
            .SelectionFormula = .SelectionFormula & "{spSIFPlanillaCompara;1.Tipo} = " & cboTipos.ItemData(cboTipos.ListIndex)
         End If
  
  
    Case optX.Item(8).Value 'Cliente Corporativo
        
        Select Case cboReporte.Text
          Case "Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_ClienteCorRsm.rpt")
          
          Case "Línea Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_ClienteCorDetalle.rpt")
        
          Case "Persona Resumen"
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_ClienteCorPersonaRsm.rpt")
          
          Case Else
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_ClienteCorDetalle.rpt")
          
          
        End Select
            .StoredProcParam(0) = cboInstitucion.ItemData(cboInstitucion.ListIndex)
            .StoredProcParam(1) = txtProceso.Text
  
  End Select
  
    .Action = 1

End With

Me.MousePointer = vbDefault

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub


strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
Call OpenRecordSet(rs, strSQL)
'    cboInstitucion.Text = rs!Descripcion
    mFrecuencPago = rs!Frecuencia_Id
rs.Close

'Refresca la Lista de Lineas si la opción esta en Analisis de Efectividad
If optX.Item(7).Value Then
  Call sbLineas
End If


End Sub

Private Sub cboReporte_Click()
If optX.Item(6).Value Then
  dtpInicio.Visible = True
Else
  dtpInicio.Visible = False
End If

dtpCorte.Visible = dtpInicio.Visible

End Sub

Private Sub chkLineas_Click()
If chkLineas.Value = vbChecked Then
   cboLineas.Enabled = False
Else
   cboLineas.Enabled = True
End If
Call sbLineas
End Sub

Private Sub FlatScrollBar_Change()
Dim vFecha As Currency

On Error GoTo vError

vFecha = txtProceso.Text


If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       vFecha = fxFechaProcesoSiguiente(vFecha)
    Else
       vFecha = fxFechaProcesoAnterior(vFecha)
    End If
    
    txtProceso.Text = vFecha
      
End If



vScroll = False
FlatScrollBar.Value = 0
vScroll = True


'Refresca la Lista de Lineas si la opción esta en Analisis de Efectividad
If optX.Item(7).Value Then
  Call sbLineas
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


vPaso = True

mFrecuencPago = "M"

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

txtProceso.Text = GLOBALES.glngFechaCR

strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
    cboInstitucion.Text = rs!Descripcion
    mFrecuencPago = rs!Frecuencia_Id
rs.Close


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

vPaso = False

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


With cboTipos
  .Clear
  .AddItem "Todos"
  .ItemData(.ListCount - 1) = CStr(1000)
 
  .AddItem "Cobro Registrado"
  .ItemData(.ListCount - 1) = CStr(1)
  .AddItem "Cobro No Registrado"
  .ItemData(.ListCount - 1) = CStr(2)
  .AddItem "Cobro Registrado / No Enviado"
  .ItemData(.ListCount - 1) = CStr(3)
  .AddItem "Cobro Apl. NC."
  .ItemData(.ListCount - 1) = CStr(4)
  .AddItem "Sobrante Enviado a Fondo"
  .ItemData(.ListCount - 1) = CStr(5)
  
  .Text = "Todos"
End With

Call optX_Click(0)
End Sub


Private Sub sbLineas()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

    strSQL = "select Rtrim(P.CODIGO) as 'IdX', rtrim(C.DESCRIPCION) + '   [' +  Rtrim(P.CODIGO) +  ']'  as 'ItmX'" _
           & " from PRM_CREDITOS P INNER JOIN CATALOGO C ON P.CODIGO = C.CODIGO" _
           & " Where P.FECHA_PROCESO = " & txtProceso.Text
    If chkInstituciones.Value = vbUnchecked Then
        strSQL = strSQL & " And P.COD_INSTITUCION = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
    strSQL = strSQL & " GROUP BY C.DESCRIPCION, P.CODIGO order by C.DESCRIPCION"
    Call sbCbo_Llena_New(cboLineas, strSQL, False, True)

Me.MousePointer = vbDefault

Exit Sub


vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub optX_Click(Index As Integer)


cboReporte.Clear
cboReporte.AddItem "Resumen"
cboReporte.Text = "Resumen"

chkInstituciones.Visible = False
chkLineas.Visible = False
cboLineas.Visible = False

chkInstituciones.Value = vbUnchecked
chkLineas.Value = vbChecked
cboLineas.Enabled = False
cboTipos.Visible = False

lblEtiqueta.Item(0).Visible = False
lblEtiqueta.Item(1).Visible = False

Select Case Index
  Case 0 'Generado
    cboReporte.AddItem "Detalle"
    cboReporte.AddItem "Línea"
    cboReporte.AddItem "Línea Detalle"
    cboReporte.AddItem "Base Actual"
  Case 1  'Cargado
    cboReporte.AddItem "Detalle"
    cboReporte.AddItem "No Localizados"
  Case 2 'Desgloce
    cboReporte.AddItem "Detalle"
    cboReporte.AddItem "Agrupado: Línea"
    cboReporte.AddItem "Agrupado: Persona"
  Case 6 'Bitacora
    cboReporte.Clear
    cboReporte.AddItem "Fechas"
    cboReporte.AddItem "Proceso + Fechas"
    cboReporte.AddItem "Proceso + Institución"
    cboReporte.Text = "Fechas"
  
  Case 7 'Análisis de Efectividad de Cobros
    cboReporte.AddItem "Detalle"
    cboReporte.AddItem "Tipo"
    cboReporte.AddItem "Estadística"
    cboReporte.AddItem "Línea"
    cboReporte.AddItem "Línea Resumen"
    
    cboReporte.AddItem "Persona"
    cboReporte.AddItem "Persona Resumen"
    
    chkInstituciones.Visible = True
    chkLineas.Visible = True
    cboLineas.Visible = True
    cboTipos.Visible = True
    
    lblEtiqueta.Item(0).Visible = True
    lblEtiqueta.Item(1).Visible = True
    
    Call sbLineas
    Call chkLineas_Click
    
  Case 8 'Cliente Corporativo
    cboReporte.AddItem "Detalle"
    cboReporte.AddItem "Línea Resumen"
    cboReporte.AddItem "Persona Resumen"
  Case Else
End Select

End Sub


Private Sub txtProceso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    'Refresca la Lista de Lineas si la opción esta en Analisis de Efectividad
    If optX.Item(7).Value Then
      Call sbLineas
    End If
End If
End Sub
