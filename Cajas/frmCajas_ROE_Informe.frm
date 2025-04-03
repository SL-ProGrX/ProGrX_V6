VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_ROE_Informe 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ROE: Informes"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todas"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ComboBox cboInforme 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
      _Version        =   1572864
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
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   855
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   3360
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8705
      _ExtentY        =   1508
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
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   615
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmCajas_ROE_Informe.frx":0000
      End
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   855
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   4560
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8705
      _ExtentY        =   1508
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
      Begin XtremeSuiteControls.PushButton btnBoleta 
         Height          =   615
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "ROE"
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
         Picture         =   "frmCajas_ROE_Informe.frx":07BC
      End
      Begin XtremeSuiteControls.FlatEdit txtROE_Id 
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ROE:"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo ROE:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Inicio:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corte:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Informe:"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de ROE"
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
      Height          =   492
      Index           =   2
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4332
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmCajas_ROE_Informe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBoleta_Click()

Dim vSubTitulo As String, strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Módulo de Cajas"
     
     .Connect = glogon.ConectRPT
     
      vSubTitulo = "ROE ID: " & txtROE_Id.Text
      
     .ReportFileName = SIFGlobal.fxPathReportes("Cajas_ROE_Detallado.rpt")
     
     .Formulas(0) = "fxFecha = 'Fecha/Hora : " & fxFechaServidor & "'"
     .Formulas(1) = "fxUsuario = 'Usuario : " & glogon.Usuario & "'"
     .Formulas(2) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(3) = "fxSubTitulo = '" & vSubTitulo & "'"
     
     
     strSQL = "{vCajas_ROE.ID_ROE} = " & txtROE_Id.Text
     
    .SelectionFormula = strSQL
     
     .Action = 1

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnInforme_Click()

Dim vSubTitulo As String, strSQL As String, vWhere As Boolean

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Módulo de Cajas"
     
     .Connect = glogon.ConectRPT
     
      vSubTitulo = "Tipo ROE: " & cboTipo.Text
      If chkFechas.Value = xtpChecked Then
        vSubTitulo = vSubTitulo & " ¦ Todas las Fechas"
      Else
        vSubTitulo = vSubTitulo & " ¦ Fechas: " & Format(dtpInicio.Value, "dd/mm/yyyy") & " al " & Format(dtpCorte.Value, "dd/mm/yyyy")
      End If
     
     Select Case Mid(cboInforme.Text, 1, 1)
        Case "D"
             .ReportFileName = SIFGlobal.fxPathReportes("Cajas_ROE_Detallado.rpt")
        Case "R"
            .ReportFileName = SIFGlobal.fxPathReportes("Cajas_ROE_Resumen.rpt")
        Case "A"
            .ReportFileName = SIFGlobal.fxPathReportes("Cajas_ROE_Anulaciones.rpt")
     End Select
     
     .Formulas(0) = "fxFecha = 'Fecha/Hora : " & fxFechaServidor & "'"
     .Formulas(1) = "fxUsuario = 'Usuario : " & glogon.Usuario & "'"
     .Formulas(2) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(3) = "fxSubTitulo = '" & vSubTitulo & "'"
     
     strSQL = ""
     vWhere = False
     
     If cboTipo.Text <> "Ambos" Then
        strSQL = "{vCajas_ROE.TIPOROE} = '" & cboTipo.Text & "'"
        vWhere = True
     End If
     
     If chkFechas.Value = xtpUnchecked Then
        If vWhere Then
            strSQL = strSQL & " AND "
        End If
        If Mid(cboInforme.Text, 1, 1) = "A" Then
            strSQL = strSQL & " cdate({vCajas_ROE.FECHA_ANULACION}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        Else
            strSQL = strSQL & " cdate({vCajas_ROE.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        End If
     End If
     
     
    .SelectionFormula = strSQL
     
     .Action = 1

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
    dtpInicio.Enabled = False
Else
    dtpInicio.Enabled = True
End If


dtpCorte.Enabled = dtpInicio.Enabled
End Sub

Private Sub Form_Load()

vModulo = 5
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

cboInforme.AddItem "Detallado"
cboInforme.AddItem "Resumen"
cboInforme.AddItem "Anulaciones"
cboInforme.Text = "Detallado"

cboTipo.AddItem "Ambos"
cboTipo.AddItem "Asociado"
cboTipo.AddItem "Depositante"
cboTipo.Text = "Ambos"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call chkFechas_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
