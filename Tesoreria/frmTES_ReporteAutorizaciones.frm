VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_ReporteAutorizaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte de Autorizaciones"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmTES_ReporteAutorizaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCombos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
      Begin VB.CheckBox chkFechas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkUsuarios 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkTipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkBancos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   4335
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox cbo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   312
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   312
         Left            =   4080
         TabIndex        =   15
         Top             =   1320
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Informes de Bancos"
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
         Height          =   312
         Index           =   10
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   3972
      End
      Begin VB.Label lblDe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblHasta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   732
      End
      Begin VB.Label lblBanco 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   732
      End
      Begin VB.Label lblTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   852
      End
      Begin VB.Image imgBusqueda_Rapida 
         Height          =   255
         Index           =   0
         Left            =   5280
         Picture         =   "frmTES_ReporteAutorizaciones.frx":030A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   255
      End
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   672
      Left            =   6000
      TabIndex        =   13
      Top             =   3360
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_ReporteAutorizaciones.frx":0D18
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de Autorizaciones"
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
      Height          =   312
      Index           =   0
      Left            =   2280
      TabIndex        =   17
      Top             =   360
      Width           =   3972
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmTES_ReporteAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()

On Error GoTo vError

If vPaso Then Exit Sub

Call sbTesTiposDocsCargaCbo(cboTipo, cbo.ItemData(cbo.ListIndex))

vError:

End Sub

Private Sub cmdImprimir_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If dtpDesde > dtpHasta Then
   MsgBox "Verifique el Rango de Fechas", vbExclamation, "No se Puede Imprimir"
   Me.MousePointer = vbDefault
   Exit Sub
End If

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(1) = "Subtitulo='Del  " & Format(dtpDesde, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
    .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_SolicitudesAutorizadas.rpt")
        
    If chkBancos.Value = vbUnchecked Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = "{CHEQUES.ID_BANCO} =" & cbo.ItemData(cbo.ListIndex)
    End If
        
    If chkTipos.Value = vbUnchecked Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{CHEQUES.TIPO} ='" & fxCodigoCbo(cboTipo) & "'"
    End If
        
    If chkUsuarios.Value = vbUnchecked Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{CHEQUES.USER_AUTORIZA} ='" & Trim(txtUsuario) & "'"
    End If
        
    If chkFechas.Value = vbUnchecked Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{CHEQUES.FECHA_AUTORIZACION} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
    End If
    .SelectionFormula = strSQL

    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 9

End Sub

Private Sub Form_Load()
 vModulo = 9

 
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 dtpDesde.Value = fxFechaServidor
 dtpHasta.Value = dtpDesde.Value
 
 vPaso = True
 Call sbTesBancoCargaCboAccesoGeneral(cbo)
 vPaso = False


End Sub


Private Sub imgBusqueda_Rapida_Click(Index As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Llama al formulario de busqueda, proporcionandole los parametros tales
'               como los campos a desplegar y la columna de busqueda, asi como la
'               columna por medio de la cual seran ordenados los registros.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo vError

gBusquedas.Resultado = Trim(txtUsuario)
gBusquedas.Consulta = "Select Nombre From Usuarios"
gBusquedas.Columna = "Nombre"
gBusquedas.Orden = "Nombre"

frmBusquedas.Show vbModal
txtUsuario = gBusquedas.Resultado

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


