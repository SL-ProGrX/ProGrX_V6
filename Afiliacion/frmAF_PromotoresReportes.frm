VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_PromotoresReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes de Ejecutivos ¦ Promotores"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   HelpContextID   =   1008
   Icon            =   "frmAF_PromotoresReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkTodasFechas 
      Height          =   252
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
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
   Begin XtremeSuiteControls.CheckBox chkPromotor 
      Height          =   252
      Left            =   8040
      TabIndex        =   3
      Top             =   1560
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
   Begin XtremeSuiteControls.FlatEdit txtPromotorId 
      Height          =   312
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   8775
      _Version        =   1441793
      _ExtentX        =   15478
      _ExtentY        =   1931
      _StockProps     =   79
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   6960
         TabIndex        =   9
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
         Picture         =   "frmAF_PromotoresReportes.frx":030A
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPromotorName 
      Height          =   312
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Ejecutivos y Promotores"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
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
      TabIndex        =   10
      Top             =   360
      Width           =   6132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Promotor"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_PromotoresReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdReporte_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If dtpInicio.Value > dtpCorte.Value Then
   Me.MousePointer = vbDefault
   MsgBox "Verifique su entrada de Datos", vbExclamation, "Error en el Rango de Fechas"
   Exit Sub
End If

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Personas"

    .Connect = glogon.ConectRPT


    If GLOBALES.gstrReporte = "Resumen" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_ResumenAfiliaciones.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_DetalleAfiliaciones.rpt")
    End If

     strSQL = "{SOCIOS.ESTADOACTUAL} = 'S'"

     If chkTodasFechas.Value = xtpUnchecked Then
        strSQL = strSQL & " AND {SOCIOS.FECHAINGRESO} >= Date (" _
               & Format(dtpInicio.Value, "YYYY,MM,DD") & ") and {SOCIOS.FECHAINGRESO} <= Date (" _
               & Format(dtpCorte.Value, "YYYY,MM,DD") & ")"
     End If

     If chkPromotor.Value = xtpUnchecked And IsNumeric(txtPromotorId.Text) Then
         strSQL = strSQL & " AND {PROMOTORES.ID_PROMOTOR}=" & txtPromotorId.Text
     End If
     
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "FechaDe='" & Format(dtpInicio.Value, "dd/MM/yyyy") & "'"
    .Formulas(2) = "FechaHasta='" & Format(dtpCorte.Value, "dd/MM/yyyy") & "'"
    
    .SelectionFormula = strSQL
    .PrintReport
         
End With

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

End Sub


Private Sub txtPromotorId_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyF4 Then
    
    gBusquedas.Col1Name = "Id"
    gBusquedas.Col1Name = "Nombre"
    gBusquedas.Filtro = ""
    gBusquedas.Columna = "ID_Promotor"
    gBusquedas.Orden = "ID_Promotor"
    gBusquedas.Consulta = "Select ID_Promotor,Nombre From Promotores"

    frmBusquedas.Show vbModal

    txtPromotorId.Text = gBusquedas.Resultado
    txtPromotorName.Text = gBusquedas.Resultado
End If
    
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtPromotorName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyF4 Then
    
    gBusquedas.Col1Name = "Id"
    gBusquedas.Col1Name = "Nombre"
    gBusquedas.Filtro = ""
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "Select ID_Promotor,Nombre From Promotores"

    frmBusquedas.Show vbModal

    txtPromotorId.Text = gBusquedas.Resultado
    txtPromotorName.Text = gBusquedas.Resultado
End If
    
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
