VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Begin VB.Form frmTES_BitacoraEspecial 
   Caption         =   "[Tesorería] Bitácora Especial"
   ClientHeight    =   8124
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   15120
   Icon            =   "frmTES_BitacoraEspecial.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8124
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   7872
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5652
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   10692
      _Version        =   524288
      _ExtentX        =   18860
      _ExtentY        =   9970
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
      MaxCols         =   12
      SpreadDesigner  =   "frmTES_BitacoraEspecial.frx":6852
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ListView lswBancos 
      Height          =   3372
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   3012
      _Version        =   1245186
      _ExtentX        =   5313
      _ExtentY        =   5948
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ListView lswDocumentos 
      Height          =   1812
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3012
      _Version        =   1245186
      _ExtentX        =   5313
      _ExtentY        =   3196
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ListView lswMovimientos 
      Height          =   1572
      Left            =   3360
      TabIndex        =   10
      Top             =   480
      Width           =   3012
      _Version        =   1245186
      _ExtentX        =   5313
      _ExtentY        =   2773
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   1560
      TabIndex        =   11
      Top             =   6360
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFechas 
      Height          =   312
      Left            =   1560
      TabIndex        =   12
      Top             =   6720
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1560
      TabIndex        =   13
      Top             =   7080
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
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
      Left            =   1560
      TabIndex        =   14
      Top             =   7440
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkDocumentos 
      Height          =   204
      Left            =   2880
      TabIndex        =   19
      Top             =   4080
      Width           =   204
      _Version        =   1245186
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "Todos"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkBancos 
      Height          =   204
      Left            =   2880
      TabIndex        =   20
      Top             =   120
      Width           =   204
      _Version        =   1245186
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "Todos"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMovimientos 
      Height          =   252
      Left            =   5400
      TabIndex        =   21
      Top             =   120
      Width           =   972
      _Version        =   1245186
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpMovFecInicio 
      Height          =   312
      Left            =   7920
      TabIndex        =   22
      Top             =   480
      Width           =   1332
      _Version        =   1245186
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpMovFecCorte 
      Height          =   312
      Left            =   9240
      TabIndex        =   23
      Top             =   480
      Width           =   1332
      _Version        =   1245186
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   7920
      TabIndex        =   24
      Top             =   840
      Width           =   2652
      _Version        =   1245186
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fraRevision 
      Height          =   492
      Left            =   6720
      TabIndex        =   25
      Top             =   1524
      Width           =   5292
      _Version        =   1245186
      _ExtentX        =   9334
      _ExtentY        =   868
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   252
         Left            =   2760
         TabIndex        =   26
         Top             =   120
         Width           =   2292
         _Version        =   1245186
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Buscar Usuario/Fecha Revisión"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
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
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   312
         Left            =   1200
         TabIndex        =   27
         Top             =   120
         Width           =   1332
         _Version        =   1245186
         _ExtentX        =   2350
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
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   11
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   7692
      End
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   612
      Left            =   10680
      TabIndex        =   29
      Top             =   480
      Width           =   1212
      _Version        =   1245186
      _ExtentX        =   2138
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmTES_BitacoraEspecial.frx":714F
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   612
      Left            =   11880
      TabIndex        =   30
      Top             =   480
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Informe"
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
      Picture         =   "frmTES_BitacoraEspecial.frx":7B6D
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   612
      Left            =   13440
      TabIndex        =   31
      Top             =   480
      Width           =   1572
      _Version        =   1245186
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmTES_BitacoraEspecial.frx":8329
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Referencia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   6720
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   7080
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   7440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   1212
   End
   Begin VB.Label Label1 
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
      Index           =   10
      Left            =   6840
      TabIndex        =   7
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
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
      Height          =   315
      Index           =   9
      Left            =   6840
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas y Usuario del Movimiento ...:"
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
      Index           =   7
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   4572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos ...:"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuentas Bancarias...:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Documento ...:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2412
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmTES_BitacoraEspecial.frx":8B2E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3200
   End
End
Attribute VB_Name = "frmTES_BitacoraEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "# Solicitud"
    vHeaders.Headers(3) = "# Documento"
    vHeaders.Headers(4) = "Tipo.Doc."
    vHeaders.Headers(5) = "Monto"
    vHeaders.Headers(6) = "Estado"
    vHeaders.Headers(7) = "Fec.Mov."
    vHeaders.Headers(8) = "Movimiento"
    vHeaders.Headers(9) = "Detalle"
    vHeaders.Headers(10) = "Usuario"
    vHeaders.Headers(11) = "Revisado por"
    vHeaders.Headers(12) = "Revisado Fecha"

Call sbSIFGridExportar(vGrid, vHeaders, "Tesoreria_BitacoraEspecial")



End Sub

Private Sub btnInforme_Click()
        vGrid.PrintHeader = "Tesorería: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpMovFecInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpMovFecCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Sub

Private Sub cboFechas_Click()

If cboFechas.Text = "[Todas]" Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub


Private Sub cboRevision_Click()
If cboRevision.ListCount = 0 Then Exit Sub
Call sbBuscar
End Sub

Private Sub chkBancos_Click()
Dim i As Integer

For i = 1 To lswBancos.ListItems.Count
  lswBancos.ListItems.Item(i).Checked = chkBancos.Value
Next i

End Sub

Private Sub chkDocumentos_Click()
Dim i As Integer

For i = 1 To lswDocumentos.ListItems.Count
  lswDocumentos.ListItems.Item(i).Checked = chkDocumentos.Value
Next i

End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vCadena As String, curMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.nsolicitud,isnull(C.ndocumento,0) as 'NDocumento',C.Tipo,C.monto,case when C.estado = 'I' or C.estado = 'E' or C.estado = 'T' then 'Emitido'" _
       & " when C.estado = 'A' then 'Anulado' when C.estado = 'P' then 'Pendiente' end as Estado" _
       & ",H.FECHA,M.DESCRIPCION,H.DETALLE, H.USUARIO,H.revisado_usuario,H.revisado_Fecha,H.ID,  case when H.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " inner join TES_HISTORIAL H ON  C.NSOLICITUD = H.NSOLICITUD" _
       & " inner join TES_TIPOS_MOVIMIENTOS M ON H.COD_MOVIMIENTO = M.COD_MOVIMIENTO " _


'Lista de Tes_Bancos
vCadena = " and C.id_banco in(0"
For i = 1 To lswBancos.ListItems.Count
  If lswBancos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "," & lswBancos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & ")"


'Lista de Tipos de Documentos
vCadena = " and C.Tipo in('"
For i = 1 To lswDocumentos.ListItems.Count
  If lswDocumentos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswDocumentos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"


'Lista de Tipos de Movimientos
vCadena = " and M.cod_movimiento in('"
For i = 1 To lswMovimientos.ListItems.Count
  If lswMovimientos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswMovimientos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"

'Fechas del Movimiento
If chkRevision.Value = vbChecked Then
     strSQL = strSQL & " and H.Revisado_fecha between '" & Format(dtpMovFecInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpMovFecCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
Else
     strSQL = strSQL & " and H.fecha between '" & Format(dtpMovFecInicio.Value, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpMovFecCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
End If

'Usuario que Realiza el Movimiento
If Len(Trim(txtUsuario)) > 0 Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and H.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and H.Usuario = '" & txtUsuario & "'"
     End If
End If

Select Case cboFechas.Text
  Case "Emisión"
    strSQL = strSQL & " and C.fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "Anulación"
    strSQL = strSQL & " and C.fecha_anula between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "Solicitud"
    strSQL = strSQL & " and C.fecha_solicitud between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

Select Case cboEstado.Text
  Case "Emitido"
     strSQL = strSQL & " and C.estado in('I','T','E')"
  Case "Anulado"
     strSQL = strSQL & " and C.estado = 'A'"
  Case "Solicitado"
     strSQL = strSQL & " and C.estado = 'P'"
End Select

Select Case Mid(cboRevision.Text, 1, 1)
   Case "P" 'Pendientes
        strSQL = strSQL & " and H.Revisado_Fecha is null"
   Case "R" 'Revisados
        strSQL = strSQL & " and H.Revisado_Fecha is not null"
   Case "T" 'Todos
End Select

If chkRevision.Value = vbChecked Then
    strSQL = strSQL & " order by H.Revisado_fecha"
Else
    strSQL = strSQL & " order by H.fecha"
End If


vPaso = True
vGrid.MaxRows = 0
curMonto = 0
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  curMonto = curMonto + rs!Monto
  
  vGrid.col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!Id
  
  
  vGrid.col = 2
  vGrid.Text = CStr(rs!NSolicitud)
  vGrid.col = 3
  vGrid.Text = CStr(rs!nDocumento) & ""
  vGrid.col = 4
  vGrid.Text = rs!Tipo
  vGrid.col = 5
  vGrid.Text = CStr(rs!Monto)
  vGrid.col = 6
  vGrid.Text = rs!Estado
  vGrid.col = 7
  vGrid.Text = rs!fecha
  vGrid.col = 8
  vGrid.Text = rs!Descripcion
  vGrid.col = 9
  vGrid.Text = rs!Detalle
  vGrid.col = 10
  vGrid.Text = rs!Usuario
  vGrid.col = 11
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.col = 12
  vGrid.Text = rs!Revisado_Fecha & ""
  
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporte()
Dim strSQL As String, vRango As String

On Error GoTo vError

Me.MousePointer = vbHourglass

'If Mid(cbo, 1, 2) = "01" Then
'  vRango = "Código : " & txtBuscarPor
'  strSQL = "MID({CHEQUES.CODIGO},1," & Len(Trim(txtBuscarPor)) & ") = '" & txtBuscarPor & "'"
'Else
'  vRango = "Beneficiario : " & txtBuscarPor
'  strSQL = "MID({CHEQUES.BENEFICIARIO},1," & Len(Trim(txtBuscarPor)) & ") = '" & txtBuscarPor & "'"
'End If
'
'
'If cboFechas.Text <> "Todas" Then
'    If Mid(cboFechas, 1, 2) = "01" Then
'      vRango = vRango & " Emision entre " & dtpInicio.Value & " y " & dtpCorte.Value
'      strSQL = strSQL & " AND {CHEQUES.FECHA_EMISION} in date(" & Year(dtpInicio.Value) & "," & Month(dtpInicio.Value) _
'             & "," & Day(dtpInicio.Value) & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) _
'             & "," & Day(dtpCorte.Value) & ")"
'    Else
'      vRango = vRango & " Anulación entre " & dtpInicio.Value & " y " & dtpCorte.Value
'      strSQL = strSQL & " AND {CHEQUES.FECHA_ANULA} in date(" & Year(dtpInicio.Value) & "," & Month(dtpInicio.Value) _
'             & "," & Day(dtpInicio.Value) & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) _
'             & "," & Day(dtpCorte.Value) & ")"
'    End If
'End If
'
'
'With frmContenedor.Crt
'    .Reset
'    .WindowShowRefreshBtn = True
'    .WindowShowPrintSetupBtn = True
'    .WindowState = crptMaximized
'    .WindowShowSearchBtn = True
'    .WindowTitle = "Reportes Módulo de Banking"
'
'    .Connect = glogon.ConectRPT
'
'    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
'    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'    .Formulas(2) = "rango='" & vRango & "'"
'
'    .ReportFileName = SIFGlobal.fxPathReportes("Banking_Desembolsos.rpt")
'    .SelectionFormula = strSQL
'
'
'    .PrintReport
'End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub chkMovimientos_Click()
Dim i As Integer

For i = 1 To lswMovimientos.ListItems.Count
  lswMovimientos.ListItems.Item(i).Checked = chkMovimientos.Value
Next i

End Sub

Private Sub chkRevision_Click()
If chkRevision.Value = vbChecked Then
   txtUsuario.BackColor = cboRevision.BackColor
Else
   txtUsuario.BackColor = vbWhite
End If
End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 9
vGrid.AppearanceStyle = fxGridStyle

cboFechas.AddItem "Emisión"
cboFechas.AddItem "Anulación"
cboFechas.AddItem "Solicitud"
cboFechas.AddItem "[Todas]"

cboFechas.Text = "[Todas]"

cboEstado.Clear
cboEstado.AddItem "Solicitado"
cboEstado.AddItem "Emitido"
cboEstado.AddItem "Anulado"
cboEstado.AddItem "[Todos]"
cboEstado.Text = "[Todos]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)


dtpMovFecInicio.Value = dtpInicio.Value
dtpMovFecCorte.Value = dtpCorte.Value


lswBancos.ColumnHeaders.Add , , "", 3150
lswBancos.ListItems.Clear
strSQL = "select id_Banco as IdX, rtrim(Descripcion) as ItmX from Tes_Bancos where estado = 'A'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswBancos.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!idX
     itmX.Checked = chkBancos.Value
 rs.MoveNext
Loop
rs.Close


lswDocumentos.ColumnHeaders.Add , , "", 3150
lswDocumentos.ListItems.Clear
strSQL = "select TIPO,DESCRIPCION from TES_TIPOS_DOC"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswDocumentos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!Tipo
     itmX.Checked = chkDocumentos.Value
 rs.MoveNext
Loop
rs.Close


lswMovimientos.ColumnHeaders.Add , , "", 3150
lswMovimientos.ListItems.Clear
strSQL = "select COD_MOVIMIENTO,DESCRIPCION from TES_TIPOS_MOVIMIENTOS"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswMovimientos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!COD_MOVIMIENTO
     itmX.Checked = chkMovimientos.Value
 rs.MoveNext
Loop
rs.Close

vPaso = True
cboRevision.AddItem "TODOS"
cboRevision.AddItem "Pendientes"
cboRevision.AddItem "Revisados"
cboRevision.Text = "TODOS"
vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - 2980

imgBanner.Height = Me.Height

End Sub





Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Or col > 1 Or Not fraRevision.Enabled Then Exit Sub
 
vGrid.Row = Row
vGrid.col = 1
If vGrid.Value = vbChecked Then
   strSQL = "update TES_HISTORIAL set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = dbo.MyGetdate()" _
          & " where [id]= " & vGrid.CellTag
   vGrid.col = 2
   strSQL = strSQL & " and NSolicitud = " & vGrid.Text
   
   Call ConectionExecute(strSQL)

   vGrid.col = 11
   vGrid.Text = glogon.Usuario
   vGrid.col = 12
   vGrid.Text = Date
   
End If

End Sub

Private Sub vGrid_DblClick(ByVal col As Long, ByVal Row As Long)
Dim frm As Form

If Row <= 0 And col = 0 Then Exit Sub
If vGrid.MaxRows <= 0 Then Exit Sub

vGrid.Row = Row
vGrid.col = 2

If vGrid.Text = "" Then Exit Sub

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(vGrid.Text)
     Exit For
   End If
 Next frm

End Sub


