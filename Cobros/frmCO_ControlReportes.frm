VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCO_ControlReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Control de Cobro"
   ClientHeight    =   6156
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10512
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6156
   ScaleWidth      =   10512
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3252
      Left            =   0
      TabIndex        =   21
      Top             =   1836
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8064
      _ExtentY        =   5736
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin VB.ComboBox cboBase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   3372
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   240
      Top             =   4080
   End
   Begin VB.ComboBox cboEPersona 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3840
      Width           =   3372
   End
   Begin VB.ComboBox cboUsuarios 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3480
      Width           =   3372
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   3372
   End
   Begin VB.ComboBox cboGestion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4560
      Width           =   4455
   End
   Begin VB.CheckBox chkFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Todas "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   10212
      _Version        =   1245187
      _ExtentX        =   18013
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   6840
         TabIndex        =   14
         Top             =   240
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
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
         Appearance      =   14
         Picture         =   "frmCO_ControlReportes.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnCubo 
         Height          =   492
         Left            =   8400
         TabIndex        =   15
         Top             =   240
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cubo"
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
         Appearance      =   14
         Picture         =   "frmCO_ControlReportes.frx":07BC
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   5292
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   6120
      TabIndex        =   18
      Top             =   2640
      Width           =   1332
      _Version        =   1245187
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
      Left            =   6120
      TabIndex        =   19
      Top             =   3000
      Width           =   1332
      _Version        =   1245187
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   4560
      TabIndex        =   22
      Top             =   1440
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8064
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Gestión de Cobros"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   2160
      TabIndex        =   17
      Top             =   360
      Width           =   4812
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   6
      Left            =   5040
      TabIndex        =   11
      Top             =   2280
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   5
      Left            =   5040
      TabIndex        =   10
      Top             =   3840
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   3
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   4272
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      Top             =   2640
      Width           =   1692
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   2
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmCO_ControlReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCubo_Click()
    lblStatus.Visible = True
    Call sbCubo
End Sub

Private Sub btnReporte_Click()

    lblStatus.Visible = False
    Call sbReportes
 
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbUnchecked Then
  dtpInicio.Enabled = True
Else
  dtpInicio.Enabled = False
End If
  
dtpCorte.Enabled = dtpInicio.Enabled
  
End Sub



Private Sub sbReporteGestiones()
Dim strSQL As String, vSubTitulo As String
Dim i As Byte

Me.MousePointer = vbHourglass


If cboEPersona.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{SOCIOS.ESTADOACTUAL}  = '" & SIFGlobal.fxCodText(cboEPersona.Text) & "'"
End If


If cboUsuarios.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{CBR_USUARIOS.USUARIO} = '" & cboUsuarios.Text & "'"
End If

If cboGestion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{CBR_GESTIONES.COD_GESTION} = '" & SIFGlobal.fxCodText(cboGestion.Text) & "'"
End If


vSubTitulo = "Gestiones : " & cboGestion.Text & "  Estado : " & cboEPersona.Text _
                 & "  Usuario : " & cboUsuarios.Text & "  Fechas: "


If chkFechas.Value = vbChecked Then
  vSubTitulo = vSubTitulo & " Todas"
Else
  vSubTitulo = vSubTitulo & " I." & Format(dtpInicio.Value, "dd/mm/yyyy") & " C." & Format(dtpCorte.Value, "dd/mm/yyyy")
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "CDATE({CBR_SEGUIMIENTO.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
End If

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"
     
    .Connect = glogon.ConectRPT
     
  Select Case lblReporte.Tag
   Case "01" 'Gestiones Realizadas
        If cboTipo.Text = "Resumen" Then
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlGestionesRealizadasRsm.rpt")
               .Formulas(1) = "Titulo='GESTIONES REALIZADAS'"
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlGestionesRealizadas.rpt")
               .Formulas(1) = "Titulo='GESTIONES REALIZADAS'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
     
   Case "02" 'Personas bajo Control
   Case "03" 'Personas sin Control
   Case "04" 'Gestiones x Usuarios
        If cboTipo.Text = "Resumen" Then
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlGestionesUsuariosRsm.rpt")
               .Formulas(1) = "Titulo='GESTIONES REALIZADAS x USUARIOS'"
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlGestionesUsuarios.rpt")
               .Formulas(1) = "Titulo='GESTIONES REALIZADAS x USUARIOS'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
   
  End Select

    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbReporteRecuperacion()
Dim strSQL As String, vSubTitulo As String
Dim i As Byte

Me.MousePointer = vbHourglass

Select Case Mid(cboEPersona.Text, 1, 2)
 Case "00" 'Todos
   strSQL = ""
 Case "01" 'Socios
   strSQL = "{vCBRControlRecuperacion.ESTADOACTUAL} = 'S'"
 Case "02" 'Opex
   strSQL = "({vCBRControlRecuperacion.ESTADOACTUAL} = 'A' OR {vCBRControlRecuperacion.ESTADOACTUAL} = 'P')"
 Case "03" 'No Socios
   strSQL = "{vCBRControlRecuperacion.ESTADOACTUAL} = 'N'"
 Case "04" 'Ren.Interna
   strSQL = "{SOCIvCBRControlRecuperacionOS.ESTADOACTUAL} = 'A'"
 Case "05" 'Ren.Patronal
   strSQL = "{SOCIOS.ESTADOACTUAL} = 'P'"
End Select


If cboUsuarios.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCBRControlRecuperacion.USUARIO} = '" & cboUsuarios.Text & "'"
End If

If cboGestion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCBRControlRecuperacion.COD_GESTION} = '" & fxCodigoCbo(cboGestion) & "'"
End If


vSubTitulo = "Gestiones : " & cboGestion.Text & "  Estado : " & cboEPersona.Text _
                 & "  Usuario : " & cboUsuarios.Text & "  Fechas: "


If chkFechas.Value = vbChecked Then
  vSubTitulo = vSubTitulo & " Todas"
Else
  vSubTitulo = vSubTitulo & " I." & Format(dtpInicio.Value, "dd/mm/yyyy") & " C." & Format(dtpCorte.Value, "dd/mm/yyyy")
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "CDATE({vCBRControlRecuperacion.FECHAGestion}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
End If

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"
     
    .Connect = glogon.ConectRPT
     
  Select Case lblReporte.Tag
   Case "09" 'Recuperación x Gestión
        If cboTipo.Text = "Resumen" Then
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlRecuperacionPorGestion.rpt")
               .Formulas(1) = "Titulo='RECUPERACION X GESTION'"
        Else
               .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlRecuperacionPorGestion.rpt")
               .Formulas(1) = "Titulo='RECUPERACION X GESTION'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
   
   Case "10" 'Recuperación x Usuario
   
   
   Case "11" 'Recuperación x Línea
   Case "12" 'Recuperación x Garantía
   Case "13" 'Recuperación Estadística
   
   

   
  End Select

    .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbReportes()

Select Case lblReporte.Tag
   Case "01" 'Gestiones Realizadas
        Call sbReporteGestiones
   Case "02" 'Personas bajo Control
   Case "03" 'Personas sin Control
   Case "04" 'Gestiones x Usuarios
        Call sbReporteGestiones
   Case "05" 'Comisiones x Gestión
   Case "06" 'Comisiones x Usuario
   Case "07" 'Cobro x Gestión
   Case "08" 'Cobro x Usuario
   Case "09" 'Recuperación x Gestión
        Call sbReporteRecuperacion
   Case "10" 'Recuperación x Usuario
        Call sbReporteRecuperacion
   Case "11" 'Recuperación x Línea
        Call sbReporteRecuperacion
   Case "12" 'Recuperación x Garantía
        Call sbReporteRecuperacion
   Case "13" 'Recuperación Estadística
        Call sbReporteRecuperacion
End Select


End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

lsw.ColumnHeaders.Add , , "", 4352

Call Formularios(Me)
Call RefrescaTags(Me)

cboBase.Clear
cboBase.AddItem "Fecha Gestion"
cboBase.AddItem "Fecha Abono"
cboBase.Text = "Fecha Gestion"

End Sub


Private Sub sbCubo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Procesando Información Espere!....Este proceso puede durar varios minutos."
lblStatus.Refresh

vMensaje = "Cobros_Recuperacion"

If chkFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpInicio.Value
  vFechaCorte = dtpCorte.Value
End If

strSQL = "exec spCbrControlRecuperacionAnalisisCubo '" & Format(vFechaInicio, "yyyy/mm/dd") & "','" & Format(dtpCorte, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


lblReporte.Caption = Item.Text
lblReporte.Tag = Item.Tag

End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, itmX As ListViewItem

TimerX.Interval = 0

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"


strSQL = "select cod_estado + ' - ' + rtrim(descripcion) as 'ItmX'" _
       & " from AFI_ESTADOS_PERSONA  where ACTIVO = 1"
Call sbLlenaCbo(cboEPersona, strSQL, True, False)

strSQL = "select cod_gestion + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  cbr_gestiones"
Call sbLlenaCbo(cboGestion, strSQL)

strSQL = "select usuario as Itmx from cbr_usuarios"
Call sbLlenaCbo(cboUsuarios, strSQL)

With lsw.ListItems
  .Clear
  Set itmX = .Add(, , "Gestiones Realizadas")
      itmX.Tag = "01"
  Set itmX = .Add(, , "Gestiones x Usuarios")
      itmX.Tag = "04"
'  Set itmX = .Add(, , "Personas bajo Control")
'      itmX.Tag = "02"
'  Set itmX = .Add(, , "Personas sin Control")
'      itmX.Tag = "03"
'  Set itmX = .Add(, , "Comisiones x Gestión")
'      itmX.Tag = "05"
'  Set itmX = .Add(, , "Comisiones x Usuario")
'      itmX.Tag = "06"
'  Set itmX = .Add(, , "Cobro x Gestión")
'      itmX.Tag = "07"
'  Set itmX = .Add(, , "Cobro x Usuario")
'      itmX.Tag = "08"
  Set itmX = .Add(, , "Recuperación x Gestión")
      itmX.Tag = "09"
  Set itmX = .Add(, , "Recuperación x Usuario")
      itmX.Tag = "10"
  Set itmX = .Add(, , "Recuperación x Línea")
      itmX.Tag = "11"
  Set itmX = .Add(, , "Recuperación x Garantía")
      itmX.Tag = "12"

      

End With


End Sub
