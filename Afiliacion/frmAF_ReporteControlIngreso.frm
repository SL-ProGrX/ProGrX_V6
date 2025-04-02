VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_ReporteControlIngreso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Control de Ingresos vrs Aportes"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   Icon            =   "frmAF_ReporteControlIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   8940
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4332
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8772
      _Version        =   1441793
      _ExtentX        =   15473
      _ExtentY        =   7641
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Personas sin Deducción"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "chkTodasFechas"
      Item(0).Control(1)=   "dtpInicio"
      Item(0).Control(2)=   "dtpCorte"
      Item(0).Control(3)=   "Label1(7)"
      Item(0).Control(4)=   "cbo"
      Item(0).Control(5)=   "cboInstitucion"
      Item(0).Control(6)=   "Label1(8)"
      Item(0).Control(7)=   "Label1(9)"
      Item(0).Control(8)=   "Label1(10)"
      Item(0).Control(9)=   "txtPriDeduc"
      Item(0).Control(10)=   "Label1(11)"
      Item(0).Control(11)=   "chkTodasProceso"
      Item(0).Control(12)=   "Label1(12)"
      Item(0).Control(13)=   "opt(0)"
      Item(0).Control(14)=   "opt(1)"
      Item(0).Control(15)=   "gbInforme"
      Item(1).Caption =   "Nombres Duplicados"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "chkDuplicacionAprox"
      Item(1).Control(1)=   "cmdDuplicados"
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   264
         Index           =   0
         Left            =   2040
         TabIndex        =   16
         Top             =   2520
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Sí"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkTodasFechas 
         Height          =   252
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   1332
         _Version        =   1441793
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
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2040
         TabIndex        =   7
         Top             =   960
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   312
         Left            =   2040
         TabIndex        =   8
         Top             =   1320
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtPriDeduc 
         Height          =   312
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Text            =   "202005"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodasProceso 
         Height          =   252
         Left            =   3480
         TabIndex        =   14
         Top             =   1800
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   264
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   2520
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "No"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox gbInforme 
         Height          =   1092
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Informe: "
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdReporte 
            Height          =   612
            Left            =   6960
            TabIndex        =   20
            Top             =   360
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_ReporteControlIngreso.frx":030A
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   264
            Index           =   0
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   466
            _StockProps     =   79
            Caption         =   "General"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   264
            Index           =   1
            Left            =   1920
            TabIndex        =   22
            Top             =   480
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   466
            _StockProps     =   79
            Caption         =   "Agrupado por Primer Deducción"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   264
            Index           =   2
            Left            =   1920
            TabIndex        =   23
            Top             =   720
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   466
            _StockProps     =   79
            Caption         =   "Agrupado por Promotor"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDuplicacionAprox 
         Height          =   492
         Left            =   -67960
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1441793
         _ExtentX        =   7006
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Buscar Casos con grado de 90% de Similitud"
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
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton cmdDuplicados 
         Height          =   852
         Left            =   -67960
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1441793
         _ExtentX        =   7006
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Reporte de Personas con Duplicación (con Número de Identificación diferente)"
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
         Picture         =   "frmAF_ReporteControlIngreso.frx":0AC6
         TextImageRelation=   4
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicó la deducción?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(aaaa-mm)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   11
         Left            =   2040
         TabIndex        =   13
         Top             =   2160
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Primer Deducción"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2652
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado de la Persona"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2652
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Ingreso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2652
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Ingreso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   6
      Left            =   2004
      TabIndex        =   0
      Top             =   360
      Width           =   5412
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_ReporteControlIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkTodasFechas_Click()
If chkTodasFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkTodasProceso_Click()
If chkTodasProceso.Value = vbChecked Then
   txtPrideduc.Enabled = False
Else
   txtPrideduc.Enabled = True
End If
End Sub

Private Sub cmdDuplicados_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Personas"
  .Connect = glogon.ConectRPT
  
  If chkDuplicacionAprox.Value = vbChecked Then
    .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoControlDupGrado.rpt")
  Else
    .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoControlDup.rpt")
  End If
  .Formulas(1) = "Titulo='Listado de Nombres Duplicados'"
  .Formulas(2) = "SubTitulo='(Cédulas Diferentes)'"
'  .SelectionFormula = strSQL
  .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()
Dim vTitulo As String, vSubTitulo As String
Dim strSQL As String

Me.MousePointer = vbHourglass

Select Case True
  Case opt.Item(1).Value
     vTitulo = "No Indica Deducción"
     strSQL = "{AHORRO_CONSOLIDADO.AHORRO} = 0"
  
  Case opt.Item(1).Value
     vTitulo = "Caso con deducción aplicada"
strSQL = "{AHORRO_CONSOLIDADO.AHORRO} > 0"

End Select



If cboInstitucion.Text <> "TODOS" Then
    strSQL = strSQL & " AND {SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
vSubTitulo = vSubTitulo & " ¦ Inst.: " & cboInstitucion.Text


If cbo.Text <> "TODOS" Then
    vSubTitulo = vSubTitulo & " ¦ Estado : " & Mid(cbo.Text, 3, 30)
    strSQL = strSQL & " AND {SOCIOS.ESTADOACTUAL} = '" & cbo.ItemData(cbo.ListIndex) & "'"
Else
    vSubTitulo = vSubTitulo & " ¦ Todos los estados"
End If

If chkTodasFechas.Value = vbChecked Then
    vSubTitulo = vSubTitulo & " ¦ Todas las Fechas de Ingreso"
Else
    vSubTitulo = vSubTitulo & " ¦ Ingreso de " & Format(dtpInicio.Value, "dd¦mm¦yyyy") _
               & " a " & Format(dtpCorte.Value, "dd¦mm¦yyyy")
    strSQL = strSQL & " AND {SOCIOS.FECHAINGRESO} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
           & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
End If

If chkTodasProceso.Value = vbChecked Then
    vSubTitulo = vSubTitulo & " ¦ Todos los Procesos"
Else
    vSubTitulo = vSubTitulo & " ¦ Proceso Deducción " & txtPrideduc
    strSQL = strSQL & " AND {SOCIOS.PRIDEDUC} = " & txtPrideduc
End If


With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowShowGroupTree = True
  .WindowTitle = "Reportes del Módulo de Personas"
  .Connect = glogon.ConectRPT
  
  
  Select Case True
    Case optReporte.Item(0).Value 'General
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoControl.rpt")
    Case optReporte.Item(1).Value 'Primer Deduccion
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoControlPriDeduc.rpt")
    Case optReporte.Item(2).Value 'Promotor
      .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoControlPromotor.rpt")
  End Select
  
  .Formulas(1) = "Titulo='" & UCase(vTitulo) & "'"
  .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
  
  .SelectionFormula = strSQL
  .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

tcMain.Item(0).Selected = True


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select cod_estado as 'IdX', descripcion as 'ItmX'" _
       & " from AFI_Estados_Persona"

Call sbCbo_Llena_New(cbo, strSQL, True, True)


strSQL = "select cod_institucion as 'IdX',descripcion as 'ItmX' from instituciones"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtPrideduc.Text = Year(dtpInicio.Value) & Format(Month(dtpInicio.Value), "00")

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
