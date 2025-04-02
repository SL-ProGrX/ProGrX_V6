VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxPReportesGenerales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CxP: Reportes Generales"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   9630
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   8281
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkProvTodos 
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      ToolTipText     =   "Todos los Proveedores"
      Top             =   480
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   444
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4575
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Filtros:"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   3120
         TabIndex        =   6
         Top             =   3960
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmCxPReportesGenerales.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   1320
         TabIndex        =   11
         Top             =   2280
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
      Begin XtremeSuiteControls.CheckBox chkFecTodas 
         Height          =   372
         Left            =   2280
         TabIndex        =   13
         Top             =   1920
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Todas las Fechas?"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkUsuarios 
         Height          =   372
         Left            =   2280
         TabIndex        =   14
         Top             =   2640
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Todos los Usuarios?"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkAnticipos 
         Height          =   372
         Left            =   2280
         TabIndex        =   15
         Top             =   3000
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Excluir Adelantos?"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkImpuesto 
         Height          =   252
         Left            =   2280
         TabIndex        =   16
         Top             =   3360
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Excluir Impuesto?"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1440
         TabIndex        =   17
         Top             =   1560
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1560
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
      Begin VB.Label lblx01 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblx02 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtProveedor 
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
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
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   1440
      Width           =   5055
      _Version        =   1441793
      _ExtentX        =   8916
      _ExtentY        =   661
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   1440
      Width           =   4575
      _Version        =   1441793
      _ExtentX        =   8070
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Informes disponibles:"
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
   Begin VB.Label lblSincroniza 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sincronizando Estado de la CxP con los pagos reales en Tesoreria. Espere.....!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCxPReportesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSubTitulo As String

Private Sub chkFecTodas_Click()
If chkFecTodas.Value = vbChecked Then
 dtpInicio.Enabled = False
Else
 dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkProvTodos_Click()
If chkProvTodos.Value = vbChecked Then
 txtCodigo.Enabled = False
Else
 txtCodigo.Enabled = True
End If

txtProveedor.Enabled = txtCodigo.Enabled

End Sub

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuario.Text = ""
   txtUsuario.Locked = True
Else
   txtUsuario.Text = glogon.Usuario
   txtUsuario.Locked = False
End If

End Sub

Private Sub cmdReporte_Click()
Dim vSQL As String
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes Módulo de Cuentas x Pagar"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 

 Select Case lblReporte.Tag
    Case "x00" 'Listado de Proveedores
         vSQL = fxSQL(0)
         .Formulas(3) = "fxTitulo = 'LISTADO DE PROVEEDORES'"
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .ReportFileName = SIFGlobal.fxPathReportes("CxP_Proveedores.rpt")
         .SelectionFormula = vSQL
    

    
    Case "x01" 'Antiguedad de Saldos
         
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'ANTIGUEDAD DE SALDOS - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_AntiguedadSaldos_CorteResumen.rpt")
            If chkProvTodos.Value = xtpChecked Or Not IsNumeric(txtCodigo.Text) Then
                .StoredProcParam(0) = 0
            Else
                .StoredProcParam(0) = txtCodigo.Text
            End If
            .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
            .StoredProcParam(2) = "R"
         Else
            .Formulas(3) = "fxTitulo = 'ANTIGUEDAD DE SALDOS - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_AntiguedadSaldos_Corte.rpt")
            If chkProvTodos.Value = xtpChecked Or Not IsNumeric(txtCodigo.Text) Then
                .StoredProcParam(0) = 0
            Else
                .StoredProcParam(0) = txtCodigo.Text
            End If
            .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
            .StoredProcParam(2) = "D"
         End If
         
            .Formulas(4) = "fxSubTitulo = 'CORTE: " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    
    Case "x02" 'Saldos de Facturas al Corte
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'SALDOS FACTURAS AL CORTE - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_Facturas_CorteResumen.rpt")
            If chkProvTodos.Value = xtpChecked Or Not IsNumeric(txtCodigo.Text) Then
                .StoredProcParam(0) = 0
            Else
                .StoredProcParam(0) = txtCodigo.Text
            End If
            .StoredProcParam(1) = "-x-"
            .StoredProcParam(2) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
            .StoredProcParam(3) = "R"
         Else
            .Formulas(3) = "fxTitulo = 'SALDOS FACTURAS AL CORTE - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_Facturas_Corte.rpt")
            If chkProvTodos.Value = xtpChecked Or Not IsNumeric(txtCodigo.Text) Then
                .StoredProcParam(0) = 0
            Else
                .StoredProcParam(0) = txtCodigo.Text
            End If
            .StoredProcParam(1) = "-x-"
            .StoredProcParam(2) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
            .StoredProcParam(3) = "D"
         End If
            
         
            .Formulas(4) = "fxSubTitulo = 'CORTE: " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    
    Case "x03.1" 'Solicitudes en Bancos Pendientes de Pago
         vSQL = fxSQL(0)
         .Formulas(3) = "fxTitulo = 'Solicitudes en Bancos Pendientes de Pago'"
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .ReportFileName = SIFGlobal.fxPathReportes("CxP_Facturas_Pendientes_Bancos.rpt")
'         .SelectionFormula = vSQL
        
    
    Case "x03.2" 'Programacion de Pagos
         vSQL = fxSQL(1)
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'PROGRAMACION DE PAGOS - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_ProgramacionAntiguedadResumen.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'PROGRAMACION DE PAGOS - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_ProgramacionAntiguedad.rpt")
         End If
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
    
    Case "x04" 'Pagos Realizados
         vSQL = fxSQL(2)
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'PAGOS REALIZADOS - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_PagosRegResumen.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'PAGOS REALIZADOS - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_PagosRegDetalle.rpt")
         End If
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
    
    Case "x05" 'Cobros x Cargos Registros
    Case "x06" 'Anticipos Registrados
    Case "x07" 'Saldos de Cargos Flotantes
    Case "x08" 'Conceptos Facturas Servicios
    
    
    Case "x09" ' "Saldos CxP al Corte [x]"
           Me.MousePointer = vbHourglass
           lblSincroniza.Visible = True
                    strSQL = "exec spCxP_SincronizaTesoreria"
                    Call ConectionExecute(strSQL)
           lblSincroniza.Visible = False
           Me.MousePointer = vbDefault
            
            .Formulas(3) = "fxTitulo = 'SALDOS CxP AL CORTE'"
           If chkProvTodos.Value = vbChecked Then
            .Formulas(4) = "fxSubTitulo = 'Fecha Corte : " & Format(dtpCorte.Value, "dd/mm/yyyy") & " / Todos los Proveedores'"
            .StoredProcParam(0) = 0
            .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
           Else
            .Formulas(4) = "fxSubTitulo = 'Fecha Corte : " & Format(dtpCorte.Value, "dd/mm/yyyy") & " / Proveedor : " & txtProveedor.Text & "'"
            .StoredProcParam(0) = txtCodigo
            .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59"
           End If
            
           If Mid(cboTipo.Text, 1, 1) = "D" Then
                .ReportFileName = SIFGlobal.fxPathReportes("CxP_SaldosCorte.rpt")
           Else
                .ReportFileName = SIFGlobal.fxPathReportes("CxP_SaldosCorte_Rsm.rpt")
           End If
    
    Case "x10", "x11", "x12" 'Facturas a Crédito
         vSQL = fxSQL(4)
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'FACTURAS REGISTRADAS - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_Facturas_Resumen.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'FACTURAS REGISTRADAS - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_Facturas.rpt")
         End If
         
         
         Select Case lblReporte.Tag
            Case Is = "x10" 'Todas
                vSubTitulo = vSubTitulo & " (Fact.Todas)"
            Case Is = "x11" 'A Crédito
                If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
                vSQL = vSQL & "{vCxP_Facturas.FORMA_PAGO} = 'CR'"
                vSubTitulo = vSubTitulo & " (Fact.Contado)"
            Case Is = "x12" 'De Contado
                If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
                vSQL = vSQL & "{vCxP_Facturas.FORMA_PAGO} = 'CO'"
                vSubTitulo = vSubTitulo & " (Fact.Contado)"
         End Select
         
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
         
         
    Case "x13" ' Informe para Tributación
         vSQL = fxSQL(4)
         If Mid(cboTipo.Text, 1, 1) = "R" Then
            .Formulas(3) = "fxTitulo = 'INFORME TRIBUTARIO - RESUMEN'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_InformeTriburario_Resumen.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'INFORME TRIBUTARIO - DETALLE'"
            .ReportFileName = SIFGlobal.fxPathReportes("CxP_InformeTriburario.rpt")
         End If
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
    
    Case Else
 End Select

'.Action = 1

 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Function fxSQL(i As Integer) As String
Dim vSQL As String

vSQL = ""
vSubTitulo = ""

Select Case i
  Case 0 'Listado de Proveedores
     
        If chkProvTodos.Value = vbUnchecked Then
          vSQL = "{CXP_PROVEEDORES.COD_PROVEEDOR} = " & txtCodigo
          vSubTitulo = "PROVEEDOR: " & UCase(txtProveedor)
        Else
          vSubTitulo = "TODOS LOS PROVEEDORES"
        End If
             
        Select Case Mid(cbo.Text, 1, 1)
           Case "A"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'A'"
           Case "I"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'I'"
           Case "S"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'S'"
           
           Case "F"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "isnull({CXP_PROVEEDORES.FUSION}) = TRUE"
        End Select
        
        vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
        
        fxSQL = vSQL
        Exit Function
     
  Case 1 'Programacion
         If chkProvTodos.Value = vbUnchecked Then
           vSQL = "{CXP_PROVEEDORES.COD_PROVEEDOR} = " & txtCodigo
           vSubTitulo = "PROVEEDOR: " & UCase(txtProveedor)
         Else
           vSubTitulo = "TODOS LOS PROVEEDORES"
         End If
              
         Select Case Mid(cbo.Text, 1, 1)
            Case "A"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'A'"
            Case "I"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'I'"
            Case "S"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'S'"
            Case "F"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "isnull({CXP_PROVEEDORES.FUSION}) = TRUE"
         End Select
         
         vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
         
         
         If chkUsuarios.Value = vbUnchecked Then
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{vCxP_ProgramacionAntiguedad.Usuario_Factura} = '" & txtUsuario.Text & "'"
              vSubTitulo = vSubTitulo & " [USUARIOS : " & txtUsuario.Text & "]"
         Else
              vSubTitulo = vSubTitulo & " [USUARIOS : TODOS ]"
         End If
         
         
        If chkAnticipos.Value = vbChecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & " mid({vCxP_ProgramacionAntiguedad.COD_FACTURA},1,4) <> 'ANT.'"
            vSubTitulo = vSubTitulo & " [Excluye Anticipos]"
        End If
         
         fxSQL = vSQL
         Exit Function
    

  Case 2 'Pagos Realizados
     
        vSQL = "ISNULL({CXP_PAGOPROV.TESORERIA}) = FALSE"
        
        If chkProvTodos.Value = vbUnchecked Then
          If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
          vSQL = "{CXP_PROVEEDORES.COD_PROVEEDOR} = " & txtCodigo
          vSubTitulo = "PROVEEDOR: " & UCase(txtProveedor)
        Else
          vSubTitulo = "TODOS LOS PROVEEDORES"
        End If
             
        Select Case Mid(cbo.Text, 1, 1)
           Case "A"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'A'"
           Case "I"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'I'"
            Case "S"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'S'"
           Case "F"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "isnull({CXP_PROVEEDORES.FUSION}) = TRUE"
        End Select
        
        vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
        
        If chkFecTodas.Value = vbUnchecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & "{CXP_PAGOPROV.FECHA_TRASLADA}" & fxFechaReportes
            vSubTitulo = vSubTitulo & " [Pagos Entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
        Else
            vSubTitulo = vSubTitulo & " [TODAS LAS FECHAS]"
        End If
        
        
        
        If chkAnticipos.Value = vbChecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & " mid({CXP_PAGOPROV.COD_FACTURA},1,4) <> 'ANT.'"
            vSubTitulo = vSubTitulo & " [Excluye Anticipos]"
        End If
        
        fxSQL = vSQL
        Exit Function


  Case 3 'Antiguedad de Saldos
        If chkProvTodos.Value = vbUnchecked Then
          vSQL = "{CXP_PROVEEDORES.COD_PROVEEDOR} = " & txtCodigo
          vSubTitulo = "PROVEEDOR: " & UCase(txtProveedor)
        Else
          vSubTitulo = "TODOS LOS PROVEEDORES"
        End If
             
        Select Case Mid(cbo.Text, 1, 1)
           Case "A"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'A'"
           Case "I"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'I'"
            Case "S"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{CXP_PROVEEDORES.ESTADO} = 'S'"
           Case "F"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "isnull({CXP_PROVEEDORES.FUSION}) = TRUE"
        End Select
        
        vSubTitulo = vSubTitulo & " [ESTADO : " & UCase(cbo.Text) & "]"
        
        
        If chkUsuarios.Value = vbUnchecked Then
             If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
             vSQL = vSQL & "{vCxP_AntiguedadSaldos.Usuario_Factura} = '" & txtUsuario.Text & "'"
             vSubTitulo = vSubTitulo & " [USUARIOS : " & txtUsuario.Text & "]"
        Else
             vSubTitulo = vSubTitulo & " [USUARIOS : TODOS ]"
        End If
        
        fxSQL = vSQL
        Exit Function

    Case 4 'Facturas Registradas e Informes Tributarios

'        vSQL = "ISNULL({vCxP_Facturas.TESORERIA}) = FALSE"
        
        If chkProvTodos.Value = vbUnchecked Then
          If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
          vSQL = "{vCxP_Facturas.COD_PROVEEDOR} = " & txtCodigo
          vSubTitulo = "PROVEEDOR: " & UCase(txtProveedor)
        Else
          vSubTitulo = "TODOS LOS PROVEEDORES"
        End If
             
        Select Case Mid(cbo.Text, 1, 1)
           Case "A"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{vCxP_Facturas.ESTADO} = 'A'"
           Case "I"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "{vCxP_Facturas.ESTADO} = 'I'"
            Case "S"
               If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
               vSQL = vSQL & "{vCxP_Facturas.ESTADO} = 'S'"
           Case "F"
              If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
              vSQL = vSQL & "isnull({vCxP_Facturas.FUSION}) = TRUE"
        End Select

        vSubTitulo = vSubTitulo & " [PROV. EST: " & UCase(cbo.Text) & "]"
        
        If chkFecTodas.Value = vbUnchecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & "{vCxP_Facturas.FECHA}" & fxFechaReportes
            vSubTitulo = vSubTitulo & " [Entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
        Else
            vSubTitulo = vSubTitulo & " [TODAS LAS FECHAS]"
        End If
        
        
        If chkUsuarios.Value = vbUnchecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & "{vCxP_Facturas.FECHA}" & fxFechaReportes
        
            vSubTitulo = vSubTitulo & " [Usuario:" & txtUsuario.Text & "]"
        Else
            vSubTitulo = vSubTitulo & " [TODOS LOS USUARIOS]"
        End If
        
        
        If chkAnticipos.Value = vbChecked Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            vSQL = vSQL & " mid({vCxP_Facturas.COD_FACTURA},1,4) <> 'ANT.'"
            vSubTitulo = vSubTitulo & " [Excluye Anticipos]"
        End If
        
        
        fxSQL = vSQL
        Exit Function
End Select

End Function


Private Function fxFechaReportes(Optional vTipo As Integer = 0) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function


Private Sub Form_Load()
vModulo = 30

Set Me.imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cbo.Clear
cbo.AddItem "Todos"
cbo.AddItem "Activos"
cbo.AddItem "Inactivos"
cbo.AddItem "Suspendidos"
cbo.AddItem "Fusionados"
cbo.Text = "Todos"

cboTipo.Clear
cboTipo.AddItem "Detallado"
cboTipo.AddItem "Resumido"
cboTipo.Text = "Detallado"


lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Informe.:", lsw.Width - 150

lsw.HideColumnHeaders = True

lsw.ListItems.Clear

lsw.ListItems.Add , "x00", "Listado de Proveedores"
lsw.ListItems.Add , "x01", "Antiguedad de Saldos"
lsw.ListItems.Add , "x02", "Saldo de Facturas al Corte [x]"
lsw.ListItems.Add , "x03.1", "Solicitudes en Bancos Pendientes de Pago"
lsw.ListItems.Add , "x03.2", "Programación de Pagos"
lsw.ListItems.Add , "x04", "Pagos Realizados"
lsw.ListItems.Add , "x05", "Cobros Registrados x Cargos"
lsw.ListItems.Add , "x06", "Anticipos Registrados"
lsw.ListItems.Add , "x07", "Saldos de Cargos Flotantes"
lsw.ListItems.Add , "x08", "Conceptos de Facturas"
lsw.ListItems.Add , "x09", "Saldos CxP al Corte [x]"
lsw.ListItems.Add , "x10", "Facturas Registradas"
lsw.ListItems.Add , "x11", "Facturas a Crédito"
lsw.ListItems.Add , "x12", "Facturas a Contado"
lsw.ListItems.Add , "x13", "Informe para Tributación"

lblReporte.Tag = "x00"
lblReporte.Caption = "Listado de Proveedores"

chkFecTodas.Value = vbChecked
chkUsuarios.Value = vbChecked

Call chkUsuarios_Click
Call chkFecTodas_Click

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
cbo.Enabled = False

dtpInicio.Enabled = False
dtpCorte.Enabled = False
chkFecTodas.Enabled = False

lblReporte.Tag = Item.Key
lblReporte.Caption = Item.Text

Select Case Item.Key
  Case "x01", "x02", "x09"
    dtpCorte.Enabled = True
    cboTipo.Enabled = True
  Case Else
    cbo.Enabled = True
    cboTipo.Enabled = True
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
    chkFecTodas.Enabled = True
End Select

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id"
  gBusquedas.Col2Name = "Cédula F/J"
  gBusquedas.Col3Name = "Razón Social"
  
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select COD_PROVEEDOR, CEDJUR, DESCRIPCION from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If

End Sub

Private Sub txtCodigo_LostFocus()
txtProveedor = fxSIFCCodigos("D", txtCodigo, "proveedores")
End Sub


Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion, cedjur from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
  cmdReporte.SetFocus
End If

End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Nombre as 'Usuario',Descripcion as 'Nombre' from Usuarios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUsuario.Text = gBusquedas.Resultado
End If
End Sub
