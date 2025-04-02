VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAH_ReportesRangosFecha 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Patrimonio"
   ClientHeight    =   5010
   ClientLeft      =   3045
   ClientTop       =   3330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "frmAH_ReportesRangosFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   612
      Left            =   4560
      TabIndex        =   2
      Top             =   4080
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
      Picture         =   "frmAH_ReportesRangosFecha.frx":030A
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   612
      Left            =   6240
      TabIndex        =   3
      Top             =   4080
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   1080
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
      Picture         =   "frmAH_ReportesRangosFecha.frx":0AC6
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   252
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Informe Consolidado Detallado"
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
      Height          =   252
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Resumen por Estado de la Persona"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Resumen por Institución"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   3360
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Informe de Aporte Patronal en Custodia"
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
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Patrimonio"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   360
      Width           =   4572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2052
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado de la Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10332
   End
End
Attribute VB_Name = "frmAH_ReportesRangosFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCerrar_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Dim strSQL As String


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vModulo = 2

strSQL = "select rtrim(cod_Estado) as  'IdX', rtrim(Descripcion) as 'ItmX'" _
       & " from AFI_ESTADOS_PERSONA Where Activo = 1"
Call sbCbo_Llena_New(cbo, strSQL, True, True)

strSQL = "select cod_institucion as Idx , Descripcion as ItmX" _
       & " from INSTITUCIONES Where Activa = 1"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

End Sub

Private Sub btnReporte_Click()
Dim strRuta As String, strSQL As String
Dim sSubtitulo As String

Me.MousePointer = vbHourglass

strSQL = ""
sSubtitulo = "Estados.: " & cbo.Text & "  ¦  Institución.: " & cboInstitucion.Text

If cbo.Text <> "TODOS" Then
   strSQL = "{SOCIOS.ESTADOACTUAL} = '" & cbo.ItemData(cbo.ListIndex) & "'"
End If

If cboInstitucion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If



With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Patrimonio"

   .Connect = glogon.ConectRPT
   
   .Formulas(0) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
   .Formulas(1) = "fxUsuario='" & glogon.Usuario & "'"
   .Formulas(2) = "fxFecha='" & fxFechaServidor & "'"
   .Formulas(3) = "SubTitulo='" & sSubtitulo & "'"



   Select Case True
'       Case OptX.Item(0).Value 'Informe por Provincias
'          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_ConsolidadoProvincia.rpt")
'
       Case OptX.Item(1).Value 'Informe Consolidado
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_Consolidado.rpt")
       
       Case OptX.Item(2).Value 'Informe Consolidado Estado
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_Consolidado_Estado.rpt")
       
       Case OptX.Item(3).Value 'Informe Consolidado Institucion
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_Consolidado_Institucion.rpt")
       
       Case OptX.Item(4).Value 'Informe Custodia
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_AporteEnCustodia.rpt")
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vPAT_Consolidado.CUSTODIA} > 0"
       
   End Select
   
   .SelectionFormula = strSQL
   .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


