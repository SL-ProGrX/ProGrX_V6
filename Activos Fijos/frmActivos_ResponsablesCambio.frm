VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_ResponsablesCambio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traslado/Cambio de Responsables"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnProcesar 
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Top             =   0
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Procesar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmActivos_ResponsablesCambio.frx":0000
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   11295
      _Version        =   1572864
      _ExtentX        =   19923
      _ExtentY        =   8281
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Responsables"
      Item(0).ControlCount=   14
      Item(0).Control(0)=   "txtPersona"
      Item(0).Control(1)=   "txtDepartamento"
      Item(0).Control(2)=   "txtSeccion"
      Item(0).Control(3)=   "txtNuevoPersona"
      Item(0).Control(4)=   "txtNuevoDepartamento"
      Item(0).Control(5)=   "txtNuevoSeccion"
      Item(0).Control(6)=   "scSubTitulos(1)"
      Item(0).Control(7)=   "scSubTitulos(0)"
      Item(0).Control(8)=   "Label5(3)"
      Item(0).Control(9)=   "Label5(2)"
      Item(0).Control(10)=   "Label5(1)"
      Item(0).Control(11)=   "Label5(0)"
      Item(0).Control(12)=   "Label5(10)"
      Item(0).Control(13)=   "Label5(11)"
      Item(1).Caption =   "Activos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4095
         Left            =   -70000
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
         _ExtentY        =   7223
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmActivos_ResponsablesCambio.frx":0719
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPersona 
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDepartamento 
         Height          =   315
         Left            =   3000
         TabIndex        =   12
         Top             =   1680
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSeccion 
         Height          =   315
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2040
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoPersona 
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   3240
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoDepartamento 
         Height          =   315
         Left            =   3000
         TabIndex        =   15
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3600
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoSeccion 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   3960
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11028
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   11
         Left            =   1080
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Departamento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Persona"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   22
         Top             =   3240
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Persona"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   21
         Top             =   2040
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sección"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   20
         Top             =   3960
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sección"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         Top             =   3600
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Departamento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19918
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Responsable Actual"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   2640
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19918
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Trasladar a .:"
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
         VisualTheme     =   3
         Alignment       =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7170
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Solicitado Por"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Solicitado Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Resuelto Por"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Resuelto Fecha"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Procesado Por"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Procesado Fecha"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboMotivo 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   675
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
      _ExtentY        =   1191
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   1320
      TabIndex        =   10
      Top             =   480
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.PushButton btnDescartar 
      Height          =   375
      Left            =   9840
      TabIndex        =   26
      Top             =   0
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Descartar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmActivos_ResponsablesCambio.frx":0ECC
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   315
      Left            =   9000
      TabIndex        =   27
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1080
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin XtremeSuiteControls.FlatEdit txtFechaSolicitud 
      Height          =   315
      Left            =   9000
      TabIndex        =   31
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1440
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin XtremeSuiteControls.FlatEdit txtFechaResolucion 
      Height          =   315
      Left            =   9000
      TabIndex        =   32
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1800
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   9960
      TabIndex        =   34
      Top             =   600
      Width           =   1335
      _Version        =   1572864
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   29
      Top             =   600
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha de Aplicación"
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
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   7
      Left            =   7800
      TabIndex        =   33
      Top             =   1080
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estado"
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
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Boleta"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resolución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7800
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Notas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Motivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitud"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmActivos_ResponsablesCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Dim vEdita As Boolean, vCodigo As String, vTipo As String
Dim vMascara As String, vScroll As Boolean

Private Sub btnDescartar_Click()
On Error GoTo vError

If txtCodigo.Text = "" Then
    MsgBox "No se ha indicado ninguna boleta de traslado de responsable?", vbExclamation
    Exit Sub
End If
If txtEstado.Tag <> "S" Then
    MsgBox "La boleta de traslado de responsable no se encuentra solicitada!", vbExclamation
    Exit Sub
End If


Dim i As Integer

i = MsgBox("Esta seguro que desea DESCARTAR esta boleta de Cambio de Responsable?", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Responsable_Cambio_Descarta '" & txtCodigo.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Boleta de Cambio de Responsables: Descartada Satisfactoriamente!", vbInformation
   Call sbConsulta(txtCodigo.Text)
   Call sbActivos_Boleta_Cambio_Responsable(txtCodigo.Text)
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnProcesar_Click()
On Error GoTo vError

If txtCodigo.Text = "" Then
    MsgBox "No se ha indicado ninguna boleta de traslado de responsable?", vbExclamation
    Exit Sub
End If
If txtEstado.Tag <> "S" Then
    MsgBox "La boleta de traslado de responsable no se encuentra solicitada!", vbExclamation
    Exit Sub
End If

Dim i As Integer

i = MsgBox("Esta seguro que desea procesar esta boleta de Cambio de Responsable?", vbYesNo)
If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spActivos_Responsable_Cambio_Procesa '" & txtCodigo.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Boleta de Cambio de Responsables: Procesada Satisfactoriamente!", vbInformation
   Call sbConsulta(txtCodigo.Text)
   Call sbActivos_Boleta_Cambio_Responsable(txtCodigo.Text)
  
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_traslado from Activos_Traslados"
    
    If txtCodigo = "" And FlatScrollBar.Value = 1 Then txtCodigo.Text = "0"
    If txtCodigo = "" And FlatScrollBar.Value = 0 Then txtCodigo.Text = "999999999"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " Where cod_traslado > '" & Format(txtCodigo, vMascara) & "' order by cod_traslado asc"
    Else
       strSQL = strSQL & " Where cod_traslado < '" & Format(txtCodigo, vMascara) & "' order by cod_traslado desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_traslado)
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 36

 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 strSQL = "select rtrim(cod_Motivo) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " FROM ACTIVOS_TRASLADOS_MOTIVOS WHERE ACTIVO = 1 order by cod_Motivo"
 Call sbCbo_Llena_New(cboMotivo, strSQL, False, False)
 
 vMascara = "0000000000"
 vEdita = True
 
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 
 btnDescartar.Tag = btnProcesar.Tag
  
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""
txtCodigo = ""

tcMain.Item(0).Selected = True

dtpFecha.Value = fxFechaServidor

txtFechaSolicitud.Text = ""
txtFechaResolucion.Text = ""


txtNotas = ""

txtEstado = ""
txtEstado.Tag = "S"

txtPersona.Tag = ""
txtDepartamento.Tag = ""
txtSeccion.Tag = ""

txtPersona.Text = ""
txtDepartamento.Text = ""
txtSeccion.Text = ""

txtNuevoPersona.Tag = ""
txtNuevoDepartamento.Tag = ""
txtNuevoSeccion.Tag = ""

txtNuevoPersona.Text = ""
txtNuevoDepartamento.Text = ""
txtNuevoSeccion.Text = ""

vGrid.MaxRows = 0
vGrid.MaxCols = 6

txtCodigo.Enabled = True

With StatusBarX.Panels
  .Item(1).Text = ""
  .Item(2).Text = ""
  .Item(3).Text = ""
  .Item(4).Text = ""
  .Item(5).Text = ""
End With


End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'Select Case Item.Index
'    Case 0
'
'    Case 1
'       If txtPersona.Tag = "" Then
'            MsgBox "No ha indicado un responsable actual, verifique!", vbExclamation
'            vGrid.MaxRows = 0
'            Exit Sub
'       End If
'
'       Call sbConsulta_Placas(txtCodigo.Text)
'
'End Select
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      
      cboMotivo.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
        Call txtCodigo_KeyDown(vbKeyF4, 1)
        
    Case "REPORTES"
       If txtCodigo.Text <> "" Then
            Call sbActivos_Boleta_Cambio_Responsable(txtCodigo.Text)
       End If
       
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(ByVal pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpiaPantalla

strSQL = "exec spActivos_Responsable_Cambio_Consulta '" & pCodigo & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  cboMotivo.SetFocus
  
  vEdita = True
  vCodigo = rs!cod_traslado
  txtCodigo.Text = rs!cod_traslado
  
  Call sbCboAsignaDato(cboMotivo, rs!Motivo, True, rs!Cod_Motivo)
  
  txtNotas = rs!Notas & ""
    
  txtEstado.Tag = rs!Estado
  txtEstado.Text = rs!Estado_Desc
  
  txtFechaSolicitud.Text = rs!Registro_Usuario & ""
  txtFechaResolucion.Text = rs!Procesado_Usuario & ""
  
  dtpFecha.Value = rs!Fecha_Aplicacion
  
  txtPersona.Tag = rs!Identificacion
  txtDepartamento.Tag = rs!Cod_Departamento
  txtSeccion.Tag = rs!Cod_Seccion
  
  txtPersona.Text = rs!Persona
  txtDepartamento.Text = rs!Departamento
  txtSeccion.Text = rs!Seccion
  
  txtNuevoPersona.Tag = rs!Identificacion_Destino
  txtNuevoDepartamento.Tag = rs!Cod_Departamento_Destino
  txtNuevoSeccion.Tag = rs!Cod_Seccion_Destino
  
  txtNuevoPersona.Text = rs!Persona_Destino
  txtNuevoDepartamento.Text = rs!Departamento_Destino
  txtNuevoSeccion.Text = rs!Seccion_Destino
  
    
  With StatusBarX.Panels
    .Item(1) = rs!Registro_Usuario
    .Item(2) = rs!Registro_fecha
    .Item(3) = rs!Cerrado_fecha & ""
    .Item(4) = rs!Cerrado_Usuario & ""
    .Item(5) = rs!Procesado_Usuario & ""
    .Item(6) = rs!Procesado_fecha & ""
  End With
    
  'Carga Activos
  strSQL = "exec spActivos_Responsable_Cambio_Consulta_Placas '" & vCodigo & "', '" & txtPersona.Tag & "', '" & glogon.Usuario & "'"
  Call OpenRecordSet(rs, strSQL)
  With vGrid
    .MaxRows = 0
     
     Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .Value = rs!Asignado
        .Col = 2
        .Text = rs!num_placa
        .Col = 3
        .Text = rs!Descripcion
        .Col = 4
        .Text = Format(rs!DEPRECIACION_AC, "Standard")
        .Col = 5
        .Text = Format(rs!DEPRECIACION_MES, "Standard")
        .Col = 6
        .Text = Format(rs!VALOR_LIBROS, "Standard")
        rs.MoveNext
     Loop
     
  End With
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConsulta_Placas(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass
    
'Carga Activos
strSQL = "exec spActivos_Responsable_Cambio_Consulta_Placas '" & pCodigo & "', '" & txtPersona.Tag & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
With vGrid
  .MaxRows = 0
   
   Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .Col = 1
      .Value = rs!Asignado
      .Col = 2
      .Text = rs!num_placa
      .Col = 3
      .Text = rs!Descripcion
      .Col = 4
      .Text = Format(rs!DEPRECIACION_AC, "Standard")
      .Col = 5
      .Text = Format(rs!DEPRECIACION_MES, "Standard")
      .Col = 6
      .Text = Format(rs!VALOR_LIBROS, "Standard")
      rs.MoveNext
   Loop
   rs.Close
   
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Function fxValida() As Boolean
Dim vMensaje As String, i As Long, iPlacas As Long

vMensaje = ""
fxValida = True

On Error GoTo vError

'vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1, 4)
'
'If Not fxInvPeriodos(dtpSolicitud.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."


If Len(txtNotas.Text) < 10 Then vMensaje = vMensaje & vbCrLf & "- No ha indicado una nota válida"
If cboMotivo.ListCount < 0 Then vMensaje = vMensaje & vbCrLf & "- No exite o no ha indicado un motivo"

If txtPersona.Tag = "" Then vMensaje = vMensaje & vbCrLf & "- No se ha indicado un responsable actual"
If txtNuevoPersona.Tag = "" Then vMensaje = vMensaje & vbCrLf & "- No se ha indicado un responsable destino"

If vGrid.MaxRows = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha cargado ningún activo a trasladar"

iPlacas = 0
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 1
  If vGrid.Value = 1 Then
     iPlacas = iPlacas + 1
  End If
Next i

If iPlacas = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha Seleccionado ningún activo a trasladar"


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim i As Long, iLineaUno As Boolean

On Error GoTo vError


If txtEstado.Tag <> "S" Then
  MsgBox "Esta Transaccion no esta solicitada, No se puede Modificar...", vbExclamation
  Exit Sub
End If

                            
strSQL = "exec spActivos_Responsable_Cambio_Boleta_Add '" & txtCodigo.Text & "', '" & cboMotivo.ItemData(cboMotivo.ListIndex) _
       & "', '" & txtNotas.Text & "', '" & txtPersona.Tag & "', '" & txtNuevoPersona.Tag & "', '" & glogon.Usuario _
       & "', '" & Format(dtpFecha.Value, "yyyy-mm-dd") & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
    vCodigo = rs!Boleta
    txtCodigo.Text = rs!Boleta
    If vEdita Then
        Call Bitacora("Modifica", "Boleta de Cambio Responsable:  " & vCodigo)
    Else
        Call Bitacora("Registra", "Boleta de Cambio Responsable:  " & vCodigo)
    End If
    
    strSQL = ""
    iLineaUno = True
    For i = 1 To vGrid.MaxRows
      vGrid.Row = i
      
      vGrid.Col = 1
      
      If vGrid.Value = 1 Then
       
        vGrid.Col = 2
        If iLineaUno Then
          iLineaUno = False
          strSQL = strSQL & Space(10) & "exec spActivos_Responsable_Cambio_Boleta_Placas '" & vCodigo & "', '" & vGrid.Text & "', '" & glogon.Usuario & "', 1"
        Else
          strSQL = strSQL & Space(10) & "exec spActivos_Responsable_Cambio_Boleta_Placas '" & vCodigo & "', '" & vGrid.Text & "', '" & glogon.Usuario & "', 0"
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
      
      End If
    Next i
    
    'Ultimo Lote
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
    End If
    
    MsgBox "Boleta Registrada Satisfactoriamente!", vbInformation
    Call sbConsulta(vCodigo)

Else
    MsgBox rs!Mensaje, vbExclamation
End If



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete pv_InvTranSac where tipo = '" & vTipo & "' and Boleta = '" & vCodigo & "'"
'  Call ConectionExecute(strSQL)
'
'  Call Bitacora("Elimina", "Transac.Inv.Tipo (" & vTipo & ") Cod." & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'
'InvTransacRep.Tipo = vTipo
'InvTransacRep.Boleta = vCodigo
'InvTransacRep.Reporte = ""
'
'Select Case UCase(ButtonMenu.Key)
'  Case "REPBOLETA"
'     frmInvTransacReporteOrden.Show vbModal
'
'  Case "REPLISTADOGENERAL"
'     Call MuestraForms(frmInvTransacReportes)
'
'End Select


End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    tcMain.Item(0).Selected = True
    cboMotivo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Boleta Id"
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col1Name = "Responsable"
    
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "cod_Traslado"
    gBusquedas.Orden = "cod_Traslado"
    gBusquedas.Consulta = "select Cod_Traslado, identificacion,Persona, Estado_Desc, Registro_Fecha from vActivos_Traslados_Boletas"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
        Call sbConsulta(gBusquedas.Resultado)
    End If
    
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub


Private Sub sbResponsableActual(pIdentificacion As String)

On Error GoTo vError

txtDepartamento.Text = ""
txtSeccion.Text = ""

txtDepartamento.Tag = ""
txtSeccion.Tag = ""

strSQL = "select * from vActivos_Personas where identificacion = '" & pIdentificacion & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
    txtDepartamento.Text = rs!Departamento
    txtSeccion.Text = rs!Seccion
    
    txtDepartamento.Tag = rs!Cod_Departamento
    txtSeccion.Tag = rs!Cod_Seccion
End If
rs.Close

Call sbConsulta_Placas(txtCodigo.Text)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbResponsableNuevo(pIdentificacion As String)
On Error GoTo vError

txtNuevoDepartamento.Text = ""
txtNuevoSeccion.Text = ""

txtNuevoDepartamento.Tag = ""
txtNuevoSeccion.Tag = ""

strSQL = "select * from vActivos_Personas where identificacion = '" & pIdentificacion & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
    txtNuevoDepartamento.Text = rs!Departamento
    txtNuevoSeccion.Text = rs!Seccion
    
    txtNuevoDepartamento.Tag = rs!Cod_Departamento
    txtNuevoSeccion.Tag = rs!Cod_Seccion
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtNuevoPersona_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "select identificacion,Nombre from Activos_Personas"
    gBusquedas.Filtro = " and Identificacion <> '" & txtPersona.Tag & "'"
    frmBusquedas.Show vbModal
    txtNuevoPersona.Tag = gBusquedas.Resultado
    txtNuevoPersona.Text = gBusquedas.Resultado2
    Call sbResponsableNuevo(txtNuevoPersona.Tag)
    
End If

End Sub

Private Sub txtPersona_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "select identificacion,Nombre from Activos_Personas"
    gBusquedas.Filtro = " and Identificacion <> '" & txtNuevoPersona.Tag & "'"
    frmBusquedas.Show vbModal
    txtPersona.Tag = gBusquedas.Resultado
    txtPersona.Text = gBusquedas.Resultado2
    Call sbResponsableActual(txtPersona.Tag)
    
End If
End Sub

