VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmFNDContratosInformes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Certificados: Boletas"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   240
      Top             =   2400
   End
   Begin XtremeSuiteControls.CheckBox chkIdAlterna 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   7800
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Utiliza Identificación alterna?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtDirigidoA 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   5640
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "A quién interese"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   615
      Left            =   7080
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2984
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
      Appearance      =   16
      Picture         =   "frmFNDContratosInformes.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   615
      Left            =   8760
      TabIndex        =   3
      Top             =   7680
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   1080
      _StockProps     =   79
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
      Picture         =   "frmFNDContratosInformes.frx":07BC
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Certificado de Ahorro a Plazo"
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3720
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Gradiente (Cupones)"
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
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   4080
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ultimo Retiro/Liquidación"
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
   Begin XtremeSuiteControls.FlatEdit txtEmitidoPor 
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   6240
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Responsable"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPuesto 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6840
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Puesto"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOperadora 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   1800
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPlan 
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   2280
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   21
      Top             =   4440
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Autorización de Deducción"
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
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   22
      Top             =   4800
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Solicitud de Ahorros"
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
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   23
      Top             =   5160
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Contrato Cuenta Sinpe"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Contrato :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Plan :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Operadora :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   1320
      Width           =   7575
      _Version        =   1441793
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   12
      Top             =   8280
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Dirigido a :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Emitido por :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   6840
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Puesto :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Certificados, Gradientes y otros"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   360
      Width           =   4575
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmFNDContratosInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mCedula As String

Private Sub btnCerrar_Click()
 Unload Me
End Sub



Private Sub btnReporte_Click()
Dim vLiquidacion As Long

If OptX.Item(5).Value Then
    Call sbFnd_Contratos_Cuenta_Sinpe(scMain(0).Caption)
    Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = False
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Planes de Ahorros"

   .Connect = glogon.ConectRPT
   
   strSQL = ""

   Select Case True
       Case OptX.Item(0).Value 'Certificado de Ahorro a Plazo
            
            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Certificado_Boleta.rpt")

            .StoredProcParam(0) = txtOperadora.Tag
            .StoredProcParam(1) = txtPlan.Tag
            .StoredProcParam(2) = txtContrato.Text
            .StoredProcParam(3) = glogon.Usuario
        
       Case OptX.Item(1).Value 'Gradiente (Cupones)
       
       
        strSQL = "{FND_CONTRATOS.COD_OPERADORA} = " & txtOperadora.Tag _
                & " And {FND_CONTRATOS.COD_PLAN} = '" & txtPlan.Tag _
                & "' and {FND_CONTRATOS.COD_CONTRATO} = " & txtContrato.Text
        
          .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratoCDPCupones.rpt")
          .Formulas(0) = "Fecha='Fecha: " & Format(fxFechaServidor, "yyyy-mm-dd") & "'"
          .Formulas(1) = "Usuario='Usuario: " & Trim(glogon.Usuario) & "'"
          .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
          .Formulas(3) = "SubTitulo='GRADIENTE (CUPONES)'"
          
          
          .Formulas(4) = "fxEjecutivo='" & txtEmitidoPor.Text & "'"
          .Formulas(5) = "fxPuesto='" & txtPuesto.Text & "'"
          
          .SelectionFormula = strSQL
          

        
'        .SubreportToChange = "sbTexto"
'        .StoredProcParam(0) = mCedula
'        .StoredProcParam(1) = chkIdAlterna.Value
'        .StoredProcParam(2) = 2
'
'        .SubreportToChange = "sbPatrimonio"
'        .StoredProcParam(0) = mCedula


       Case OptX.Item(2).Value 'Ultimo Retiro/Liquidación
         
            strSQL = "Select isnull(max(Consec),0) as Consec From Fnd_liquidacion" _
                   & " where cod_plan = '" & txtPlan.Tag & "' and cod_operadora = " _
                   & txtOperadora.Tag & " and cod_contrato = " _
                   & txtContrato.Text
            Call OpenRecordSet(rs, strSQL)
              vLiquidacion = rs!consec
            rs.Close
       
            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionBoleta.rpt")
            .SelectionFormula = "{FND_LIQUIDACION.CONSEC} =" & vLiquidacion
            .Formulas(0) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
            .SubreportToChange = "sbAsiento"
            .StoredProcParam(0) = "FLIQ"
            .StoredProcParam(1) = vLiquidacion
            .StoredProcParam(2) = 1
            
       Case OptX.Item(3).Value 'Autorización de Deducción
       
       
        strSQL = "{vFnd_Contratos.CEDULA} = '" & mCedula & "'" _
                & " And {vFnd_Contratos.ESTADO} = 'A' And {vFnd_Contratos.MONTO} > 0"
        
          .ReportFileName = SIFGlobal.fxPathReportes("Fondos_AutorizacionDeduccion.rpt")
          .Formulas(0) = "Fecha='Fecha: " & Format(fxFechaServidor, "yyyy-mm-dd") & "'"
          .Formulas(1) = "Usuario='Usuario: " & Trim(glogon.Usuario) & "'"
          .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
          
         .StoredProcParam(0) = mCedula
          
'        .SubreportToChange = "sbAhorros"
'        .SelectionFormula = strSQL
          

       Case OptX.Item(4).Value 'Solicitud de Ahorros
       
       
        
          .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Solicitud_Ahorros.rpt")
          .Formulas(0) = "Fecha='Fecha: " & Format(fxFechaServidor, "yyyy-mm-dd") & "'"
          .Formulas(1) = "Usuario='Usuario: " & Trim(glogon.Usuario) & "'"
          .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
          
        
        strSQL = "{vAFI_Persona_Integral.CEDULA} = '" & mCedula & "'"
        .SelectionFormula = strSQL
          
        strSQL = "{vFnd_Contratos.CEDULA} = '" & mCedula & "'" _
                & " And {vFnd_Contratos.ESTADO} = 'A' And {vFnd_Contratos.MONTO} > 0"
          
        .SubreportToChange = "sbAhorros"
        .SelectionFormula = strSQL

        
'        .SubreportToChange = "sbTexto"
'        .StoredProcParam(1) = chkIdAlterna.Value
'        .StoredProcParam(2) = 2
'
'        .SubreportToChange = "sbPatrimonio"
'        .StoredProcParam(0) = mCedula

       
   End Select
   
   .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbInicializa()

On Error GoTo vError

strSQL = "select C.COD_OPERADORA, C.COD_PLAN, c.COD_CONTRATO, c.CEDULA, S.NOMBRE" _
       & " , O.DESCRIPCION as 'Operadora_Desc'" _
       & " , P.DESCRIPCION as 'Plan_Desc'" _
       & " , P.COD_MONEDA" _
       & " from FND_CONTRATOS C" _
       & "    inner join SOCIOS S on C.CEDULA = S.CEDULA" _
       & "    inner join FND_OPERADORAS O on C.COD_OPERADORA = C.COD_OPERADORA" _
       & "    inner join FND_PLANES P on C.COD_OPERADORA = P.COD_OPERADORA and C.COD_PLAN = P.COD_PLAN" _
       & " Where C.Cod_Operadora = " & gFondos.Operadora & " and C.Cod_Plan = '" & gFondos.Plan _
       & "' and C.Cod_Contrato = " & gFondos.Contrato

Call OpenRecordSet(rs, strSQL)

txtOperadora.Tag = rs!Cod_Operadora
txtOperadora.Text = rs!OPERADORA_DESC

txtPlan.Tag = rs!Cod_Plan
txtPlan.Text = rs!Plan_Desc

txtContrato.Text = rs!Cod_Contrato

scMain.Item(0).Caption = Trim(rs!Cedula)
scMain.Item(1).Caption = Trim(rs!Nombre)

mCedula = rs!Cedula

txtPuesto.Text = ""

strSQL = "select descripcion from Usuarios where Nombre  = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
  txtEmitidoPor.Text = rs!Descripcion
rs.Close

Exit Sub

vError:
End Sub


Private Sub Form_Load()

vModulo = 18

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Exit Sub

vError:

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub
