VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmActivos_Reasignacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Responsable (Traslado de Activo)"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6612
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   11292
      _Version        =   1441793
      _ExtentX        =   19918
      _ExtentY        =   11663
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
      Item(0).Caption =   "Traslados"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "txtBoleta"
      Item(0).Control(1)=   "Label1(0)"
      Item(0).Control(2)=   "Label1(1)"
      Item(0).Control(3)=   "dtpFecha"
      Item(0).Control(4)=   "txtPersona"
      Item(0).Control(5)=   "txtDepartamento"
      Item(0).Control(6)=   "Label5(11)"
      Item(0).Control(7)=   "Label5(10)"
      Item(0).Control(8)=   "txtSeccion"
      Item(0).Control(9)=   "txtNuevoPersona"
      Item(0).Control(10)=   "Label5(0)"
      Item(0).Control(11)=   "Label5(1)"
      Item(0).Control(12)=   "txtNuevoDepartamento"
      Item(0).Control(13)=   "txtNuevoSeccion"
      Item(0).Control(14)=   "Label5(2)"
      Item(0).Control(15)=   "Label5(3)"
      Item(0).Control(16)=   "scSubTitulos(0)"
      Item(0).Control(17)=   "scSubTitulos(1)"
      Item(0).Control(18)=   "txtNotas"
      Item(0).Control(19)=   "Label5(4)"
      Item(0).Control(20)=   "cboMotivo"
      Item(0).Control(21)=   "Label5(5)"
      Item(0).Control(22)=   "cmdNuevo"
      Item(0).Control(23)=   "cmdGuardar"
      Item(1).Caption =   "Boletas"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "lbl"
      Item(1).Control(1)=   "chkTodos"
      Item(1).Control(2)=   "lswHistorial"
      Item(1).Control(3)=   "Label1(2)"
      Item(1).Control(4)=   "dtpInicio"
      Item(1).Control(5)=   "dtpCorte"
      Item(1).Control(6)=   "chkFechasTodas"
      Item(1).Control(7)=   "Label1(3)"
      Item(1).Control(8)=   "txtBoletaInicio"
      Item(1).Control(9)=   "txtBoletaCorte"
      Item(1).Control(10)=   "chkTodosActivos"
      Item(1).Control(11)=   "chkBoletasTodas"
      Item(1).Control(12)=   "btnBoletas(0)"
      Item(1).Control(13)=   "btnBoletas(1)"
      Begin XtremeSuiteControls.ListView lswHistorial 
         Height          =   4572
         Left            =   -70000
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   11292
         _Version        =   1441793
         _ExtentX        =   19918
         _ExtentY        =   8064
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkFechasTodas 
         Height          =   252
         Left            =   -64480
         TabIndex        =   34
         Top             =   5400
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Left            =   -69760
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBoleta 
         Height          =   432
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   762
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   312
         Left            =   7920
         TabIndex        =   7
         Top             =   480
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
      Begin XtremeSuiteControls.FlatEdit txtPersona 
         Height          =   312
         Left            =   3000
         TabIndex        =   8
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDepartamento 
         Height          =   312
         Left            =   3000
         TabIndex        =   9
         Top             =   2160
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSeccion 
         Height          =   312
         Left            =   3000
         TabIndex        =   12
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2520
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoPersona 
         Height          =   312
         Left            =   3000
         TabIndex        =   13
         Top             =   3720
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoDepartamento 
         Height          =   312
         Left            =   3000
         TabIndex        =   16
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   4080
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNuevoSeccion 
         Height          =   312
         Left            =   3000
         TabIndex        =   17
         Top             =   4440
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboMotivo 
         Height          =   312
         Left            =   3000
         TabIndex        =   24
         Top             =   5040
         Width           =   6252
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton cmdNuevo 
         Height          =   312
         Left            =   9960
         TabIndex        =   26
         ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
         Top             =   0
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmActivos_Reasignacion.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   612
         Left            =   9480
         TabIndex        =   27
         ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
         Top             =   5760
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cambio de Responsable"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmActivos_Reasignacion.frx":0632
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1032
         Left            =   3000
         TabIndex        =   22
         Top             =   5400
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
         _ExtentY        =   1820
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -67240
         TabIndex        =   32
         Top             =   5400
         Visible         =   0   'False
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
         Left            =   -65920
         TabIndex        =   33
         Top             =   5400
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtBoletaInicio 
         Height          =   312
         Left            =   -67240
         TabIndex        =   36
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBoletaCorte 
         Height          =   312
         Left            =   -65920
         TabIndex        =   37
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkBoletasTodas 
         Height          =   252
         Left            =   -64480
         TabIndex        =   38
         Top             =   5760
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTodosActivos 
         Height          =   252
         Left            =   -64480
         TabIndex        =   39
         Top             =   6120
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos los Activos ?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnBoletas 
         Height          =   612
         Index           =   0
         Left            =   -61720
         TabIndex        =   40
         ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmActivos_Reasignacion.frx":0D4B
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnBoletas 
         Height          =   612
         Index           =   1
         Left            =   -60400
         TabIndex        =   41
         ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmActivos_Reasignacion.frx":1769
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   3
         Left            =   -69400
         TabIndex        =   35
         Top             =   5760
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Boletas Entre:"
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
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   -69400
         TabIndex        =   31
         Top             =   5400
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha de aplicación:"
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
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   372
         Left            =   -70000
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   11292
         _Version        =   1441793
         _ExtentX        =   19918
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Indique las Boletas de Traslados de Activos que visualizar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   5
         Left            =   1080
         TabIndex        =   25
         Top             =   5040
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Motivo"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   4
         Left            =   1080
         TabIndex        =   23
         Top             =   5400
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
         BackColor       =   -2147483633
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
         Height          =   372
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   3120
         Width           =   11292
         _Version        =   1441793
         _ExtentX        =   19918
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Trasladar a .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
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
         Height          =   372
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   11292
         _Version        =   1441793
         _ExtentX        =   19918
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Responsable Actual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         Top             =   4080
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Departamento"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   4440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sección"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   2520
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sección"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Top             =   3720
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Persona"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   10
         Left            =   1080
         TabIndex        =   11
         Top             =   1800
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Persona"
         BackColor       =   -2147483633
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
         Height          =   252
         Index           =   11
         Left            =   1080
         TabIndex        =   10
         Top             =   2160
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Departamento"
         BackColor       =   -2147483633
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha de aplicación"
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
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Boleta No."
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   762
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblDescripcion 
      Height          =   432
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
      _ExtentY        =   762
      _StockProps     =   79
      Caption         =   "xx"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   372
      Index           =   21
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "No. Placa"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11532
   End
End
Attribute VB_Name = "frmActivos_Reasignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMascara As String


Private Sub sbBoleta(vBoleta As String)
'Imprime la Boleta de Traslado

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de Activos Fijos"
 
 .Connect = glogon.ConectRPT

 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxSubTitulo = 'TRASLADO DE ACTIVOS Y CAMBIO DE RESPONSABLES'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Activos_BoletaTraslado.rpt")
 .SelectionFormula = "{ACTIVOS_TRASLADOS.COD_TRASLADO} = '" & vBoleta & "'"
  
 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub chkBoletasTodas_Click()
If chkBoletasTodas.Value = vbChecked Then
   txtBoletaInicio.Enabled = False
Else
   txtBoletaInicio.Enabled = True
End If

txtBoletaCorte.Enabled = txtBoletaInicio.Enabled

End Sub

Private Sub btnBoletas_Click(Index As Integer)
Dim i As Integer

Select Case Index
  Case 0 'Buscar
    Call sbBoletasConsulta

  Case 1 'Reporte

        For i = 1 To lswHistorial.ListItems.Count
         If lswHistorial.ListItems.Item(i).Checked Then
            Call sbBoleta(lswHistorial.ListItems.Item(i).Text)
         End If
        Next i

End Select
End Sub

Private Sub chkFechasTodas_Click()
If chkFechasTodas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lswHistorial.ListItems.Count
 lswHistorial.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String, vBoleta As String
Dim vMensaje As String, i As Integer

On Error GoTo vError


vMensaje = ""

'Verificacion
If Len(Trim(txtCodigo.Tag)) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó ningun activo..."
If Len(Trim(txtPersona.Tag)) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó ningun responsable actual..."
If Len(Trim(txtNuevoPersona.Tag)) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó ningun responsable nuevo..."
'Verificar nota y verifica fecha de aplicacion


If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass


strSQL = "exec  spActivos_ResponsableCambio '" & txtBoleta.Text & "','" & txtCodigo.Text & "','" & cboMotivo.ItemData(cboMotivo.ListIndex) _
       & "','" & txtNuevoPersona.Tag & "','" & glogon.Usuario & "','" & txtNotas.Text & "','P','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL, 0)
  vBoleta = rs!Boleta
rs.Close


Call Bitacora("Registra", "Cambio Responsable, Activo: " & txtCodigo.Text & ", Persona Origen: " & txtPersona.Tag _
            & " a Persona Destino: " & txtNuevoPersona.Tag)
            
'Boleta
Call sbBoleta(vBoleta)

Me.MousePointer = vbDefault
MsgBox "Traslado Realizado Satisfactoriamente...", vbInformation

Call cmdNuevo_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdNuevo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

tcMain.Item(0).Selected = True

strSQL = "select isnull(max(cod_traslado),0) as Boleta from Activos_traslados"
Call OpenRecordSet(rs, strSQL, 0)
  If IsNumeric(rs!Boleta) Then
         txtBoleta = Format(CLng(rs!Boleta) + 1, vMascara)
  End If
rs.Close

txtCodigo = ""
txtCodigo.Tag = ""
lblDescripcion.Caption = ""

txtPersona.Text = ""
txtDepartamento.Text = ""
txtSeccion.Text = ""

txtPersona.Tag = ""
txtDepartamento.Tag = ""
txtSeccion.Tag = ""


'Nuevos
txtNuevoPersona.Text = ""
txtNuevoDepartamento.Text = ""
txtNuevoSeccion.Text = ""

txtNuevoPersona.Tag = ""
txtNuevoDepartamento.Tag = ""
txtNuevoSeccion.Tag = ""

txtNotas.Text = ""


End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 36

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 strSQL = "select rtrim(cod_Motivo) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " FROM ACTIVOS_TRASLADOS_MOTIVOS WHERE ACTIVO = 1 order by cod_Motivo"
 Call sbCbo_Llena_New(cboMotivo, strSQL, False, True)
 
vMascara = "0000000000"

dtpFecha.Value = gActivos.Periodo
dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)

With lswHistorial.ColumnHeaders
    .Clear
    .Add , , "Boleta No.", 1400
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Usuario", 1800, vbCenter
    .Add , , "Persona Origen", 3000
    .Add , , "Persona Destino", 3000
    .Add , , "Motivo", 2400
End With

Call Formularios(Me)
Call RefrescaTags(Me)

Call cmdNuevo_Click

End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True

strSQL = "select A.num_placa,A.nombre,T.descripcion as 'TipoActivo', A.cod_departamento,A.cod_seccion,A.identificacion" _
       & ",D.descripcion as 'Departamento',S.descripcion as 'Seccion',P.Nombre as 'Persona'" _
       & " from Activos_Principal A inner join Activos_tipo_Activo T" _
       & " on A.tipo_activo = T.tipo_activo" _
       & " inner join Activos_departamentos D on A.cod_departamento = D.cod_departamento" _
       & " inner join Activos_secciones S on A.cod_seccion = S.cod_seccion" _
       & " and A.cod_departamento = S.cod_departamento" _
       & " inner join Activos_Personas P on A.identificacion = P.identificacion" _
       & " where A.num_placa = '" & txtCodigo & "'"
       
Call OpenRecordSet(rs, strSQL, 0)

'Limpia Datps
Call cmdNuevo_Click

If Not rs.EOF And Not rs.BOF Then
  txtCodigo.Text = rs!num_placa
  txtCodigo.Tag = rs!num_placa
  lblDescripcion.Caption = rs!Nombre
  
  txtPersona.Text = rs!Persona
  txtPersona.Tag = rs!Identificacion
  
  txtDepartamento.Tag = rs!Cod_Departamento
  txtDepartamento.Text = rs!departamento
  
  txtSeccion.Tag = rs!Cod_Seccion
  txtSeccion.Text = rs!seccion
  
  txtNuevoPersona.SetFocus
  
End If
rs.Close

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBoletasConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vWhere As Boolean


On Error GoTo vError

Me.MousePointer = vbHourglass

lswHistorial.ListItems.Clear
chkTodos.Value = vbUnchecked
vWhere = False

strSQL = "select * " _
       & " from vActivos_TrasladosHistorico "
       
If chkTodosActivos.Value = vbUnchecked Then
   strSQL = strSQL & " where num_placa = '" & txtCodigo.Tag & "'"
   vWhere = True
End If

If chkFechasTodas.Value = vbUnchecked Then
   strSQL = strSQL & IIf(vWhere, " And ", " Where ") _
          & " Fecha_Aplicacion between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
          & " and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
   vWhere = True
End If

If chkBoletasTodas.Value = vbUnchecked Then
   strSQL = strSQL & IIf(vWhere, " And ", " Where ") _
          & " cod_Traslado between '" & Format(txtBoletaInicio.Text, vMascara) _
          & "' and '" & Format(txtBoletaCorte.Text, vMascara) & "'"
   vWhere = True
End If

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswHistorial.ListItems.Add(, , rs!cod_traslado)
     itmX.SubItems(1) = rs!registro_fecha
     itmX.SubItems(2) = rs!registro_usuario
     itmX.SubItems(3) = rs!Persona
     itmX.SubItems(4) = rs!Persona_Destino
     itmX.SubItems(5) = rs!Motivo
    
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Or txtCodigo.Tag = "" Then
  Exit Sub
End If

Call sbBoletasConsulta
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo.Text <> "" Then Call sbConsulta
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  
  gBusquedas.Col1Name = "Id Placa"
  gBusquedas.Col2Name = "Id Alterna"
  gBusquedas.Col3Name = "Nombre"
  
  gBusquedas.Consulta = "select num_placa, Placa_Alterna, Nombre from Activos_Principal"
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta
End If

End Sub

Private Sub txtDepartamento_Change()
txtSeccion.Tag = ""
txtSeccion = ""
End Sub

Private Sub txtDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSeccion.SetFocus
  
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from Activos_departamentos"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtDepartamento.Tag) Then
       txtDepartamento.Tag = gBusquedas.Resultado
       txtDepartamento = gBusquedas.Resultado2
       txtSeccion.SetFocus
    End If
End If

End Sub

Private Sub sbResponsableNuevo(pIdentificacion As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

txtNuevoDepartamento.Text = ""
txtNuevoSeccion.Text = ""

txtNuevoDepartamento.Tag = ""
txtNuevoSeccion.Tag = ""

strSQL = "select * from vActivos_Personas where identificacion = '" & txtPersona.Tag & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
    txtNuevoDepartamento.Text = rs!departamento
    txtNuevoSeccion.Text = rs!seccion
    
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
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtNuevoPersona.Tag = gBusquedas.Resultado
    txtNuevoPersona.Text = gBusquedas.Resultado2
    Call sbResponsableNuevo(txtNuevoPersona.Tag)
    
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNuevoPersona.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select num_placa,nombre from Activos_Principal"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta
End If

End Sub
