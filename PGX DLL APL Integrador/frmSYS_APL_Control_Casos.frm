VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSYS_APL_Control_Casos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "APL: Control de Casos"
   ClientHeight    =   7980
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   14004
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   14004
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1812
      Left            =   3600
      TabIndex        =   20
      Top             =   5880
      Width           =   10092
      _Version        =   1310720
      _ExtentX        =   17801
      _ExtentY        =   3196
      _StockProps     =   79
      Caption         =   "Resumen por rango de fechas:"
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
      Begin XtremeSuiteControls.ListView lswResumen 
         Height          =   1452
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   8052
         _Version        =   1310720
         _ExtentX        =   14203
         _ExtentY        =   2561
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
         Sorted          =   -1  'True
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton optResumen 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "General"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optResumen 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Línea/PYME"
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
      End
      Begin XtremeSuiteControls.RadioButton optResumen 
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Institución"
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
      End
      Begin XtremeSuiteControls.RadioButton optResumen 
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Usuario"
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
      End
      Begin XtremeSuiteControls.PushButton btnExportarLsw 
         Height          =   372
         Left            =   9600
         TabIndex        =   29
         Top             =   0
         Width           =   372
         _Version        =   1310720
         _ExtentX        =   656
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmSYS_APL_Control_Casos.frx":0000
      End
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   7932
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3252
      _Version        =   1310720
      _ExtentX        =   5736
      _ExtentY        =   13991
      _StockProps     =   79
      Caption         =   "Consulta"
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
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   492
         Left            =   1680
         TabIndex        =   1
         Top             =   6960
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Consulta"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmSYS_APL_Control_Casos.frx":0805
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   240
         TabIndex        =   2
         Top             =   5040
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   240
         TabIndex        =   3
         Top             =   5760
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   550
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
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDominio 
         Height          =   312
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   550
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
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   312
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   240
         TabIndex        =   6
         Top             =   6480
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
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
         Height          =   312
         Left            =   1680
         TabIndex        =   7
         Top             =   6480
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtPlan 
         Height          =   312
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtInstitucion 
         Height          =   312
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDepartamento 
         Height          =   312
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSeccion 
         Height          =   312
         Left            =   240
         TabIndex        =   26
         Top             =   4320
         Width           =   2772
         _Version        =   1310720
         _ExtentX        =   4890
         _ExtentY        =   556
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   492
         Left            =   240
         TabIndex        =   28
         Top             =   6960
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Exportar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmSYS_APL_Control_Casos.frx":1223
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sección:"
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
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   4080
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
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
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Institución:"
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
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan:"
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
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1092
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   5520
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dominio:"
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
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea [Pyme]:"
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
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   372
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   4800
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas:"
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
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   6240
         Width           =   1092
      End
      Begin VB.Image imgBanner 
         Height          =   9396
         Left            =   0
         Picture         =   "frmSYS_APL_Control_Casos.frx":1A28
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3240
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   3480
      TabIndex        =   13
      Top             =   240
      Width           =   10212
      _Version        =   524288
      _ExtentX        =   18013
      _ExtentY        =   9335
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   40
      SpreadDesigner  =   "frmSYS_APL_Control_Casos.frx":2992
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmSYS_APL_Control_Casos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnConsulta_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vAPL_Analisis_Main" _
       & " WHERE COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"

Select Case cboEstado.Text
    Case "Recibidas"
        strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
        
        strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        
        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Registro_Usuario = '" & txtUsuario.Tag & "'"
        End If
        
    Case "Pendientes"
        strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
    
        strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Registro_Usuario = '" & txtUsuario.Tag & "'"
        End If
    
    Case "Autorizadas"
        strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
    
        strSQL = strSQL & " and Aprobada_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Aprobada_Usuario = '" & txtUsuario.Tag & "'"
        End If
    
    Case "Denegadas"
        strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
        
        strSQL = strSQL & " and Denegada_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Denegada_Usuario = '" & txtUsuario.Tag & "'"
        End If
    
    Case "Formalizadas"
        strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
    
        strSQL = strSQL & " and Formalizai_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        
        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Formalizai_Usuario = '" & txtUsuario.Tag & "'"
        End If
        
    Case "[TODAS]"
        strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

        If txtUsuario.Text <> "" Then
            strSQL = strSQL & " and Registro_Usuario = '" & txtUsuario.Tag & "'"
        End If

End Select


If txtLinea.Text <> "" Then
    strSQL = strSQL & " and COD_LINEA = '" & txtLinea.Tag & "'"
End If

If txtPlan.Text <> "" Then
    strSQL = strSQL & " and COD_PLAN = '" & txtPlan.Tag & "'"
End If

If txtInstitucion.Text <> "" Then
    strSQL = strSQL & " and COD_INSTITUCION = " & txtInstitucion.Tag
End If

If txtDepartamento.Text <> "" Then
    strSQL = strSQL & " and COD_DEPARTAMENTO = '" & txtDepartamento.Tag & "'"
End If

If txtSeccion.Text <> "" Then
    strSQL = strSQL & " and COD_SECCION = '" & txtSeccion.Tag & "'"
End If


With vGrid

.MaxRows = 0

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  For i = 2 To .MaxCols
    .Col = i
    Select Case i
        Case 2 'Dominio
            .Text = rs!Cod_Dominio & ""
        Case 3 'No.Solicitud
            .Text = rs!APL_OPERACION & ""
        Case 4 'Estado
            .Text = rs!Estado_Desc & ""
        Case 5 'Linea
            .Text = rs!Linea_Desc & ""
        Case 6 'Identificacion
            .Text = rs!Cedula & ""
        Case 7 'Nombre
            .Text = rs!Nombre & ""
        Case 8 'Monto
            .Text = Format(rs!Factura_Monto, "Standard")
        Case 9 'Plazo
            .Text = CStr(rs!Plazo & "")
        Case 10 'Tasa
            .Text = Format(rs!Tasa, "Standard")
        Case 11 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        Case 12 'Plan Desc
            .Text = rs!Plan_desc & ""
        Case 13 'Institución
            .Text = rs!Institucion_desc & ""
        Case 14 'Factura
            .Text = rs!FACTURA_NUMERO & ""
        Case 15 'ProGrX Operacion
            .Text = CStr(rs!Operacion & "")
        Case 16 'Registro Fecha
            .Text = rs!registro_Fecha & ""
        Case 17 'Registro Usuario
            .Text = rs!registro_usuario & ""
        Case 18 'Atiende Fecha
            .Text = rs!Atiende_Fecha & ""
        Case 19 'Atiende Usuario
            .Text = rs!Atiende_Usuario & ""
        Case 20 'Tiempo
            .Text = "0"
        Case 21 'Ex.Documento
            .Text = rs!Formaliza_Documento & ""
        Case 22 'Ex.Fecha
            .Text = rs!Formalizai_Fecha & ""
        Case 23 'Ex.Usuario
            .Text = rs!Formalizai_Usuario & ""
        Case 24 'Cobro Estado
            .Text = rs!Cobro_Estado & ""
        Case 25 'Cobro Remesa
            .Text = CStr(rs!Cobro_Remesa & "")
        Case 26 'Cobro Fecha
            .Text = rs!Cobro_Fecha & ""
        Case 27 'Cobro Usuario
            .Text = rs!Cobro_Usuario & ""
        Case 28 'Cancela Fecha
            .Text = rs!Cancela_Fecha & ""
        Case 29 'Cancela Documento
            .Text = rs!Cancela_Documento & ""
        Case 30 'Cancela Usuario
            .Text = rs!Cancela_Usuario & ""
        Case 31 'Departamento
            .Text = rs!Departamento_desc & ""
        Case 32 'Seccion
            .Text = rs!Seccion_Desc & ""
    
        Case 33 'Tel.Cel
            .Text = rs!CLIENTE_CELULAR & ""
        Case 34 'Tel
            .Text = rs!CLIENTE_TELEFONO & ""
        Case 35 'EMail
            .Text = rs!CLIENTE_EMAIL & ""
        Case 36 'Profesión
            .Text = rs!PROFESION & ""
        Case 37 'Estado Civil
            .Text = rs!ESTADO_CIVIL & ""
        Case 38 'Provincia
            .Text = rs!Provincia_Desc & ""
        Case 39 'Canton
            .Text = rs!Canton_Desc & ""
        Case 40 'Distrito
            .Text = rs!Distrito_Desc & ""
    
   
    
    End Select
  
  Next i
  rs.MoveNext
Loop
rs.Close

End With

Call sbResumen

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
 

End Sub

Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "Dominio"
    vHeaders.Headers(3) = "No.Solicitud"
    vHeaders.Headers(4) = "Estado"
    vHeaders.Headers(5) = "Linea"
    vHeaders.Headers(6) = "Identificacion"
    vHeaders.Headers(7) = "Nombre"
    vHeaders.Headers(8) = "Monto"
    vHeaders.Headers(9) = "Plazo"
    vHeaders.Headers(10) = "Tasa"
    vHeaders.Headers(11) = "Cuota"
    vHeaders.Headers(12) = "Plan Desc"
    vHeaders.Headers(13) = "Institución"
    vHeaders.Headers(14) = "Factura"
    vHeaders.Headers(15) = "ProGrX Operacion"
    vHeaders.Headers(16) = "Registro Fecha"
    vHeaders.Headers(17) = "Registro Usuario"
    vHeaders.Headers(18) = "Atiende Fecha"
    vHeaders.Headers(19) = "Atiende Usuario"
    vHeaders.Headers(20) = "Tiempo"
    vHeaders.Headers(21) = "Ex.Documento"
    vHeaders.Headers(22) = "Ex.Fecha"
    vHeaders.Headers(23) = "Ex.Usuario"
    vHeaders.Headers(24) = "Cobro Estado"
    vHeaders.Headers(25) = "Cobro Remesa"
    vHeaders.Headers(26) = "Cobro Fecha"
    vHeaders.Headers(27) = "Cobro Usuario"
    vHeaders.Headers(28) = "Cancela Fecha"
    vHeaders.Headers(29) = "Cancela Documento"
    vHeaders.Headers(30) = "Cancela Usuario"
    vHeaders.Headers(31) = "Departamento"
    vHeaders.Headers(32) = "Seccion"
    
    vHeaders.Headers(33) = "Tel. Movil"
    vHeaders.Headers(34) = "Tel. Hab."
    vHeaders.Headers(35) = "Email"
    vHeaders.Headers(36) = "Profesión"
    vHeaders.Headers(37) = "Estado Civil"
    vHeaders.Headers(38) = "Provincia"
    vHeaders.Headers(39) = "Cantón"
    vHeaders.Headers(40) = "Distrito"
    
  
    

 Call sbSIFGridExportar(vGrid, vHeaders, "Apl_Consulta_Detallada")
End Sub

Private Sub btnExportarLsw_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lswResumen)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboDominio_Click()
  txtLinea.Text = ""
  txtPlan.Text = ""
  txtInstitucion.Text = ""
  txtDepartamento.Text = ""
  txtSeccion.Text = ""
  txtUsuario.Text = ""
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 38

cboEstado.Clear
cboEstado.AddItem "Recibidas"
cboEstado.AddItem "Pendientes"
cboEstado.AddItem "Autorizadas"
cboEstado.AddItem "Denegadas"
cboEstado.AddItem "Formalizadas"
cboEstado.AddItem "[TODAS]"

cboEstado.Text = "Recibidas"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -30, dtpCorte.Value)


vGrid.MaxRows = 0

With lswResumen.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", 3000
    .Add , , "Casos", 1200, vbRightJustify
    .Add , , "T.: Monto", 2800, vbRightJustify
    
    .Add , , "Formalizadas", 1200, vbRightJustify
    .Add , , "T.: Formaliza", 2800, vbRightJustify
    
    .Add , , "Recibido", 1200, vbRightJustify
    .Add , , "T.:Recibido", 2800, vbRightJustify
    
    .Add , , "Pendientes", 1200, vbRightJustify
    .Add , , "T.:Pendiente", 2800, vbRightJustify
    
    .Add , , "Aprobadas", 1200, vbRightJustify
    .Add , , "T.: Aprobado", 2800, vbRightJustify
    
    .Add , , "Denegadas", 1200, vbRightJustify
    .Add , , "T.: Denegado", 2800, vbRightJustify

End With


vPaso = True

strSQL = "exec spAPL_Dominios_Vinculados '" & gAPL.APL_Dominio & "'"
Call sbCbo_Llena_New(cboDominio, strSQL, False, True)

vPaso = False

End Sub

Private Sub Form_Resize()

On Error Resume Next


vGrid.Height = Me.Height - (vGrid.Top + gbResumen.Height + 850)
vGrid.Width = Me.Width - (vGrid.Left + 250)


imgBanner.Height = Me.Height
gbFiltros.Height = Me.Height

gbResumen.Left = vGrid.Left
gbResumen.Top = vGrid.Top + vGrid.Height + 80

gbResumen.Width = vGrid.Width - 60

lswResumen.Width = gbResumen.Width - (lswResumen.Left + 60)
btnExportarLsw.Left = lswResumen.Left + lswResumen.Width - (btnExportarLsw.Width)
End Sub



Private Sub sbResumen()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ", SUM( 1 ) AS 'TOTAL_CASOS'" _
       & " , SUM( CONVERT(DEC(18,2),LTRIM(FACTURA_MONTO)) ) 'TOTAL_MONTO'" _
       & "     , SUM(CASE WHEN ESTADO = 'R' THEN 1 ELSE 0 END ) AS 'RECIBIDO_CASOS'" _
       & "     , SUM(CASE WHEN ESTADO = 'R' THEN CONVERT(DEC(18,2), FACTURA_MONTO) ELSE 0 END) 'RECIBIDO_MONTO'" _
       & "     , SUM(CASE WHEN ESTADO = 'P' THEN 1 ELSE 0 END ) AS 'PENDIENTE_CASOS'" _
       & "     , SUM(CASE WHEN ESTADO = 'P' THEN CONVERT(DEC(18,2), FACTURA_MONTO) ELSE 0 END) 'PENDIENTE_MONTO'" _
       & "     , SUM(CASE WHEN ESTADO = 'A' THEN 1 ELSE 0 END ) AS 'APROBADAS_CASOS'" _
       & "     , SUM(CASE WHEN ESTADO = 'A' THEN CONVERT(DEC(18,2), FACTURA_MONTO) ELSE 0 END) 'APROBADAS_MONTO'" _
       & "     , SUM(CASE WHEN ESTADO = 'D' THEN 1 ELSE 0 END ) AS 'DENEGADAS_CASOS'" _
       & "     , SUM(CASE WHEN ESTADO = 'D' THEN CONVERT(DEC(18,2), FACTURA_MONTO) ELSE 0 END) 'DENEGADAS_MONTO'" _
       & "     , SUM(CASE WHEN ESTADO = 'F' THEN 1 ELSE 0 END ) AS 'FORMALIZA_CASOS'" _
       & "     , SUM(CASE WHEN ESTADO = 'F' THEN CONVERT(DEC(18,2), FACTURA_MONTO) ELSE 0 END) 'FORMALIZA_MONTO'" _
       & " From vAPL_Analisis_Main" _
       & " WHERE COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'" _
       & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"


Select Case True
    Case optResumen.Item(0).Value 'General
            strSQL = "select COD_DOMINIO, '' as 'CODIGO', 'GENERAL' as 'DESCRIPCION'" _
                   & strSQL _
                   & " GROUP BY COD_DOMINIO"
    
    Case optResumen.Item(1).Value 'Linea
            strSQL = "select COD_DOMINIO, COD_LINEA as 'CODIGO', LINEA_DESC as 'DESCRIPCION'" _
                   & strSQL _
                   & " GROUP BY COD_DOMINIO, COD_LINEA, LINEA_DESC"
    
    
    Case optResumen.Item(2).Value 'Institucion
            strSQL = "select COD_DOMINIO, COD_INSTITUCION as 'CODIGO', INSTITUCION_DESC as 'DESCRIPCION'" _
                   & strSQL _
                   & " GROUP BY COD_DOMINIO, COD_INSTITUCION, INSTITUCION_DESC"
    
    Case optResumen.Item(3).Value 'Usuario
            strSQL = "select COD_DOMINIO, REGISTRO_USUARIO as 'CODIGO', REGISTRO_USUARIO as 'DESCRIPCION'" _
                   & strSQL _
                   & " GROUP BY COD_DOMINIO, REGISTRO_USUARIO"
    
    
End Select


With lswResumen.ListItems

    .Clear
    
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Codigo)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!Total_Casos, "###,###,##0")
          itmX.SubItems(3) = Format(rs!Total_Monto, "Standard")
          itmX.SubItems(4) = Format(rs!Formaliza_Casos, "###,###,##0")
          itmX.SubItems(5) = Format(rs!Formaliza_Monto, "Standard")
          itmX.SubItems(6) = Format(rs!Recibido_Casos, "###,###,##0")
          itmX.SubItems(7) = Format(rs!Recibido_Monto, "Standard")
          itmX.SubItems(8) = Format(rs!Pendiente_Casos, "###,###,##0")
          itmX.SubItems(9) = Format(rs!Pendiente_Monto, "Standard")
          itmX.SubItems(10) = Format(rs!Aprobadas_Casos, "###,###,##0")
          itmX.SubItems(11) = Format(rs!Aprobadas_Monto, "Standard")
          itmX.SubItems(12) = Format(rs!Denegadas_Casos, "###,###,##0")
          itmX.SubItems(13) = Format(rs!Denegadas_Monto, "Standard")
          
      rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical


End Sub


Private Sub lswResumen_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswResumen.SortKey = ColumnHeader.Index - 1
  If lswResumen.SortOrder = 0 Then lswResumen.SortOrder = 1 Else lswResumen.SortOrder = 0
  lswResumen.Sorted = True
End Sub

Private Sub optResumen_Click(Index As Integer)
Call sbResumen
End Sub

Private Sub txtDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "[Id]"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Consulta = "select COD_DEPARTAMENTO, DESCRIPCION from vAPL_Departamentos"
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    If txtInstitucion.Text <> "" Then
       gBusquedas.Filtro = gBusquedas.Filtro & " AND COD_INSTITUCION = " & txtInstitucion.Tag
    End If
    
    frmBusquedas.Show vbModal
    
    txtDepartamento.Tag = gBusquedas.Resultado
    txtDepartamento.Text = gBusquedas.Resultado2
    
    
    txtSeccion.Text = ""
    txtSeccion.Tag = ""
    
End If

End Sub

Private Sub txtInstitucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "[Id]"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Consulta = "select COD_INSTITUCION, DESCRIPCION from APL_INSTITUCIONES"
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    frmBusquedas.Show vbModal
    
    txtInstitucion.Tag = gBusquedas.Resultado
    txtInstitucion.Text = gBusquedas.Resultado2
    
    txtDepartamento.Text = ""
    txtSeccion.Text = ""
    
End If

End Sub

Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Linea"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "COD_LINEA"
    gBusquedas.Orden = "COD_LINEA"
    gBusquedas.Consulta = "select COD_LINEA, DESCRIPCION  from APL_LINEAS "
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    frmBusquedas.Show vbModal
    
    txtLinea.Tag = gBusquedas.Resultado
    txtLinea.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Plan"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "COD_PLAN"
    gBusquedas.Orden = "COD_PLAN"
    gBusquedas.Consulta = "select COD_PLAN, DESCRIPCION from APL_LINEAS_PLANES"
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    frmBusquedas.Show vbModal
    
    txtPlan.Tag = gBusquedas.Resultado
    txtPlan.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "[Id]"
    gBusquedas.Col2Name = "Descripción"
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Consulta = "select COD_SECCION, DESCRIPCION from vAPL_Secciones"
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    If txtInstitucion.Text <> "" Then
       gBusquedas.Filtro = gBusquedas.Filtro & " AND COD_INSTITUCION = " & txtInstitucion.Tag
    End If
    
    If txtDepartamento.Text <> "" Then
       gBusquedas.Filtro = gBusquedas.Filtro & " AND COD_DEPARTAMENTO = '" & txtDepartamento.Tag & "'"
    End If
    
    frmBusquedas.Show vbModal
    
    txtSeccion.Tag = gBusquedas.Resultado
    txtSeccion.Text = gBusquedas.Resultado2
End If
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Usuario Id"
    gBusquedas.Col2Name = ""
    gBusquedas.Col3Name = ""
    gBusquedas.Columna = "USUARIO"
    gBusquedas.Orden = "USUARIO"
    gBusquedas.Consulta = "select USUARIO from APL_DOMINIOS_USUARIOS "
    gBusquedas.Filtro = " AND COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
    
    frmBusquedas.Show vbModal
    
    txtUsuario.Tag = gBusquedas.Resultado
    txtUsuario.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pDominio As String, pOperacion As Long
Dim frm As Form

With vGrid
    .Row = Row
    .Col = 2
    pDominio = .Text
    .Col = 3
    pOperacion = .Text
        
    Call sbFormsCall("frmSYS_APL_Caso_Preview", , , , , Me, True)
    
    Call sbFormActivo("frmSYS_APL_Caso_Preview", frm)
    Call frm.sbConsulta_Externa(pDominio, pOperacion)

End With

End Sub
