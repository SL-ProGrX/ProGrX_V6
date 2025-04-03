VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.ShortcutBar.v20.2.0.ocx"
Begin VB.Form frmAH_Constancias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe de Constancias Patrimoniales"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.CheckBox chkIdAlterna 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   6360
      Width           =   3375
      _Version        =   1310722
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
      TabIndex        =   8
      Top             =   4200
      Width           =   6855
      _Version        =   1310722
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
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
      _Version        =   1310722
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
      Picture         =   "frmAH_Constancias.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   6240
      Width           =   855
      _Version        =   1310722
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
      Picture         =   "frmAH_Constancias.frx":07BC
   End
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
      _Version        =   1310722
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Constancia de Aporte Patronal"
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
      TabIndex        =   3
      Top             =   2520
      Width           =   4815
      _Version        =   1310722
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Constancia de Patrimonio de la Persona"
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
      TabIndex        =   4
      Top             =   2880
      Width           =   4815
      _Version        =   1310722
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Informe  Consolidado + Movimientos"
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
      TabIndex        =   10
      Top             =   4800
      Width           =   6855
      _Version        =   1310722
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
      TabIndex        =   12
      Top             =   5400
      Width           =   6855
      _Version        =   1310722
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
   Begin XtremeSuiteControls.RadioButton OptX 
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Top             =   3480
      Width           =   4815
      _Version        =   1310722
      _ExtentX        =   8488
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Constancia de Afiliación"
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
      Caption         =   "Informes y Constancias de Patrimonio"
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
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Width           =   2175
      _Version        =   1310722
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   4800
      Width           =   2175
      _Version        =   1310722
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
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
      _Version        =   1310722
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   7575
      _Version        =   1310722
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
      _Version        =   1310722
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmAH_Constancias"
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

Me.MousePointer = vbHourglass


With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = False
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Patrimonio"

   .Connect = glogon.ConectRPT
   
   strSQL = ""

   Select Case True
       Case OptX.Item(0).Value 'Constancia Patronal
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_ConstanciaPatronal.rpt")

        .Formulas(0) = "fxDirigido='" & txtDirigidoA.Text & "'"
        .Formulas(1) = "fxEmite='" & txtEmitidoPor.Text & "'"
        .Formulas(2) = "fxPuesto='" & txtPuesto.Text & "'"
        
        .SubreportToChange = "sbTexto"
        .StoredProcParam(0) = mCedula
        .StoredProcParam(1) = chkIdAlterna.Value
        .StoredProcParam(2) = 1
        
       Case OptX.Item(1).Value 'Constrancia de Patrimonio
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_ConstanciaPatrimonio.rpt")

        .Formulas(0) = "fxDirigido='" & txtDirigidoA.Text & "'"
        .Formulas(1) = "fxEmite='" & txtEmitidoPor.Text & "'"
        .Formulas(2) = "fxPuesto='" & txtPuesto.Text & "'"
        
        .SubreportToChange = "sbTexto"
        .StoredProcParam(0) = mCedula
        .StoredProcParam(1) = chkIdAlterna.Value
        .StoredProcParam(2) = 2

        .SubreportToChange = "sbPatrimonio"
        .StoredProcParam(0) = mCedula


       Case OptX.Item(2).Value 'Informe de Estado del Patrimonio
         
         .WindowShowGroupTree = True
         
         .Formulas(0) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
         .Formulas(1) = "fxUsuario='Usuario : " & glogon.Usuario & "'"
         .Formulas(2) = "fxFecha='Fecha   : " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
         
         
          strSQL = "{SOCIOS.CEDULA}='" & mCedula & "'"
        
          .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_Aportes.rpt")
  
       Case OptX.Item(3).Value 'Constancia de Afiliación
          .ReportFileName = SIFGlobal.fxPathReportes("Personas_ConstanciaAfiliacion.rpt")
        
        .Formulas(0) = "fxDirigido='" & txtDirigidoA.Text & "'"
        .Formulas(1) = "fxEmite='" & txtEmitidoPor.Text & "'"
        .Formulas(2) = "fxPuesto='" & txtPuesto.Text & "'"
        
        .SubreportToChange = "sbTexto"
        .StoredProcParam(0) = mCedula
        .StoredProcParam(1) = chkIdAlterna.Value
        .StoredProcParam(2) = 3

       
   End Select
   
   .SelectionFormula = strSQL
   .PrintReport
End With

Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

vModulo = 2

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

scMain.Item(0).Caption = GLOBALES.gTag
scMain.Item(1).Caption = GLOBALES.gTag2

mCedula = GLOBALES.gTag

txtPuesto.Text = ""

strSQL = "select descripcion from Usuarios where Nombre  = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
  txtEmitidoPor.Text = rs!Descripcion
rs.Close

Exit Sub

vError:

End Sub
