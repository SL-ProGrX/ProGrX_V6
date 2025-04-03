VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_ControlComPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Comisiones de Cobro"
   ClientHeight    =   8136
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11976
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8136
   ScaleWidth      =   11976
   Begin XtremeSuiteControls.GroupBox fraFiltros 
      Height          =   2292
      Left            =   2520
      TabIndex        =   61
      Top             =   3600
      Width           =   7932
      _Version        =   1245187
      _ExtentX        =   13991
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Filtros adicionales"
      ForeColor       =   4210752
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1800
         TabIndex        =   62
         Top             =   600
         Width           =   6012
         _Version        =   1245187
         _ExtentX        =   10605
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
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   312
         Left            =   1800
         TabIndex        =   63
         Top             =   1080
         Width           =   6012
         _Version        =   1245187
         _ExtentX        =   10605
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
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   0
         Left            =   6000
         TabIndex        =   64
         Top             =   1680
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
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
         Picture         =   "frmCO_ControlComPago.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   420
         Index           =   1
         Left            =   7320
         TabIndex        =   65
         Top             =   1680
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   741
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
         Picture         =   "frmCO_ControlComPago.frx":0A1E
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   11
         Left            =   360
         TabIndex        =   67
         Top             =   600
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cuenta Bancaria:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   12
         Left            =   360
         TabIndex        =   66
         Top             =   1080
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Usuario:"
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
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   11976
      _ExtentX        =   21124
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6492
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   11652
      _Version        =   1245187
      _ExtentX        =   20553
      _ExtentY        =   11451
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
      ItemCount       =   4
      Item(0).Caption =   "Remesa"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label8(0)"
      Item(0).Control(1)=   "Label8(1)"
      Item(0).Control(2)=   "Label8(2)"
      Item(0).Control(3)=   "Label8(3)"
      Item(0).Control(4)=   "Label8(4)"
      Item(0).Control(5)=   "Label8(5)"
      Item(0).Control(6)=   "Label8(6)"
      Item(0).Control(7)=   "txtRemesa"
      Item(0).Control(8)=   "txtFecha"
      Item(0).Control(9)=   "txtEstado"
      Item(0).Control(10)=   "txtUsuario"
      Item(0).Control(11)=   "txtNotas"
      Item(0).Control(12)=   "dtpInicio"
      Item(0).Control(13)=   "dtpCorte"
      Item(0).Control(14)=   "btnBarra(0)"
      Item(0).Control(15)=   "btnBarra(1)"
      Item(0).Control(16)=   "btnBarra(2)"
      Item(0).Control(17)=   "lswRemesas"
      Item(0).Control(18)=   "fraReporte"
      Item(1).Caption =   "Cargar"
      Item(1).ControlCount=   12
      Item(1).Control(0)=   "Label8(9)"
      Item(1).Control(1)=   "Label8(10)"
      Item(1).Control(2)=   "cboCarga"
      Item(1).Control(3)=   "cboOficina"
      Item(1).Control(4)=   "chkFiltros"
      Item(1).Control(5)=   "chkCarga"
      Item(1).Control(6)=   "lswCarga"
      Item(1).Control(7)=   "txtCargaTotal"
      Item(1).Control(8)=   "btnBarra(3)"
      Item(1).Control(9)=   "btnBarra(4)"
      Item(1).Control(10)=   "btnBarra(5)"
      Item(1).Control(11)=   "Label8(18)"
      Item(2).Caption =   "Trasladar"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "Label8(14)"
      Item(2).Control(1)=   "Label2(16)"
      Item(2).Control(2)=   "cboTraslado"
      Item(2).Control(3)=   "lswTraslado"
      Item(2).Control(4)=   "txtPagoTotal"
      Item(2).Control(5)=   "btnBarra(6)"
      Item(2).Control(6)=   "btnBarra(7)"
      Item(2).Control(7)=   "Label8(19)"
      Item(3).Caption =   "Informes"
      Item(3).ControlCount=   10
      Item(3).Control(0)=   "opt(0)"
      Item(3).Control(1)=   "txtRepRemesas"
      Item(3).Control(2)=   "Label16(2)"
      Item(3).Control(3)=   "lblRemesa"
      Item(3).Control(4)=   "opt(1)"
      Item(3).Control(5)=   "Label16(4)"
      Item(3).Control(6)=   "chkRemesaInd"
      Item(3).Control(7)=   "lswRep"
      Item(3).Control(8)=   "btnBarra(8)"
      Item(3).Control(9)=   "chkDetalle"
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   3132
         Left            =   1560
         TabIndex        =   3
         Top             =   3240
         Width           =   10092
         _Version        =   1245187
         _ExtentX        =   17801
         _ExtentY        =   5524
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCarga 
         Height          =   3972
         Left            =   -70000
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1245187
         _ExtentX        =   20553
         _ExtentY        =   7006
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswTraslado 
         Height          =   4092
         Left            =   -69880
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   11412
         _Version        =   1245187
         _ExtentX        =   20129
         _ExtentY        =   7218
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRep 
         Height          =   3612
         Left            =   -70000
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   11652
         _Version        =   1245187
         _ExtentX        =   20553
         _ExtentY        =   6371
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   0
         Left            =   -69520
         TabIndex        =   7
         Top             =   5160
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1245187
         _ExtentX        =   8488
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(Pendientes) Remesa"
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
      Begin XtremeSuiteControls.GroupBox fraReporte 
         Height          =   2052
         Left            =   4320
         TabIndex        =   8
         Top             =   960
         Width           =   7452
         _Version        =   1245187
         _ExtentX        =   13144
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Informes"
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
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   0
            Left            =   1920
            TabIndex        =   9
            Top             =   1200
            Width           =   1572
            _Version        =   1245187
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Pendientes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkRepFechas 
            Height          =   252
            Left            =   4440
            TabIndex        =   10
            Top             =   360
            Width           =   1212
            _Version        =   1245187
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
            Height          =   312
            Left            =   3120
            TabIndex        =   11
            Top             =   360
            Width           =   1212
            _Version        =   1245187
            _ExtentX        =   2138
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
            Height          =   312
            Left            =   1920
            TabIndex        =   12
            Top             =   360
            Width           =   1212
            _Version        =   1245187
            _ExtentX        =   2138
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.ComboBox cboRepOficina 
            Height          =   312
            Left            =   1920
            TabIndex        =   13
            Top             =   720
            Width           =   4932
            _Version        =   1245187
            _ExtentX        =   8700
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
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   0
            Left            =   5760
            TabIndex        =   14
            Top             =   1200
            Width           =   612
            _Version        =   1245187
            _ExtentX        =   1080
            _ExtentY        =   741
            _StockProps     =   79
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
            Picture         =   "frmCO_ControlComPago.frx":13AB
         End
         Begin XtremeSuiteControls.PushButton btnReporte 
            Height          =   420
            Index           =   1
            Left            =   6360
            TabIndex        =   15
            Top             =   1200
            Width           =   492
            _Version        =   1245187
            _ExtentX        =   868
            _ExtentY        =   741
            _StockProps     =   79
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
            Picture         =   "frmCO_ControlComPago.frx":1B67
         End
         Begin XtremeSuiteControls.RadioButton optReporte 
            Height          =   252
            Index           =   1
            Left            =   3600
            TabIndex        =   16
            Top             =   1200
            Width           =   1692
            _Version        =   1245187
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Trasladadas"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   1212
            _Version        =   1245187
            _ExtentX        =   2138
            _ExtentY        =   656
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   372
            Index           =   8
            Left            =   360
            TabIndex        =   17
            Top             =   720
            Width           =   1452
            _Version        =   1245187
            _ExtentX        =   2561
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Oficina/Agencia:"
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
         Begin VB.Image imgRepRefresca 
            Height          =   192
            Left            =   6600
            Picture         =   "frmCO_ControlComPago.frx":230A
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   360
            Width           =   192
         End
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   0
         Left            =   4320
         TabIndex        =   19
         Top             =   480
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Nueva"
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
         Picture         =   "frmCO_ControlComPago.frx":2423
      End
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   432
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   762
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   5160
         TabIndex        =   22
         Top             =   1320
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   5160
         TabIndex        =   23
         Top             =   1680
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1560
         TabIndex        =   24
         Top             =   2040
         Width           =   10092
         _Version        =   1245187
         _ExtentX        =   17801
         _ExtentY        =   1397
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1560
         TabIndex        =   25
         Top             =   1320
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   556
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
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   1
         Left            =   5640
         TabIndex        =   27
         Top             =   480
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         Picture         =   "frmCO_ControlComPago.frx":2BDC
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   2
         Left            =   6120
         TabIndex        =   28
         Top             =   480
         Width           =   492
         _Version        =   1245187
         _ExtentX        =   868
         _ExtentY        =   741
         _StockProps     =   79
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
         Picture         =   "frmCO_ControlComPago.frx":33A9
      End
      Begin XtremeSuiteControls.ComboBox cboCarga 
         Height          =   312
         Left            =   -67600
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1245187
         _ExtentX        =   13568
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -67600
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1245187
         _ExtentX        =   13568
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
      Begin XtremeSuiteControls.CheckBox chkFiltros 
         Height          =   372
         Left            =   -67600
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Filtros adicionales?"
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
      End
      Begin XtremeSuiteControls.CheckBox chkCarga 
         Height          =   252
         Left            =   -69880
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTraslado 
         Height          =   312
         Left            =   -67840
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1245187
         _ExtentX        =   13568
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   3
         Left            =   -63880
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
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
         Picture         =   "frmCO_ControlComPago.frx":3B65
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   4
         Left            =   -62560
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Cargar"
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
         Picture         =   "frmCO_ControlComPago.frx":4583
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   5
         Left            =   -61240
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmCO_ControlComPago.frx":4D45
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   6
         Left            =   -63040
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   741
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
         Picture         =   "frmCO_ControlComPago.frx":5731
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   7
         Left            =   -61720
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Traslado"
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
         Picture         =   "frmCO_ControlComPago.frx":614F
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   420
         Index           =   8
         Left            =   -60760
         TabIndex        =   39
         Top             =   5640
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   741
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
         Picture         =   "frmCO_ControlComPago.frx":6954
      End
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   312
         Left            =   -59200
         TabIndex        =   40
         Top             =   4560
         Visible         =   0   'False
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "15"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkRemesaInd 
         Height          =   372
         Left            =   -60640
         TabIndex        =   41
         Top             =   5040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Indicar Remesa"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   1
         Left            =   -69520
         TabIndex        =   42
         Top             =   5520
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1245187
         _ExtentX        =   8488
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(Traslado) Remesa"
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
      Begin XtremeSuiteControls.FlatEdit txtCargaTotal 
         Height          =   312
         Left            =   -60760
         TabIndex        =   43
         Top             =   6120
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoTotal 
         Height          =   312
         Left            =   -60880
         TabIndex        =   44
         Top             =   6000
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkDetalle 
         Height          =   372
         Left            =   -68920
         TabIndex        =   68
         Top             =   5880
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1245187
         _ExtentX        =   7006
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Incluir el Detalle en el Informe?"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   0
         Left            =   600
         TabIndex        =   60
         Top             =   480
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   1
         Left            =   600
         TabIndex        =   59
         Top             =   1320
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   2
         Left            =   4320
         TabIndex        =   58
         Top             =   1320
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Estado:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   3
         Left            =   600
         TabIndex        =   57
         Top             =   1680
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Registro:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   4
         Left            =   4320
         TabIndex        =   56
         Top             =   1680
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Usuario:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   5
         Left            =   600
         TabIndex        =   55
         Top             =   2040
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Notas:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   6
         Left            =   1560
         TabIndex        =   54
         Top             =   2880
         Width           =   2892
         _Version        =   1245187
         _ExtentX        =   5101
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Últimas Remesas Registradas:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   9
         Left            =   -69400
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   372
         Index           =   10
         Left            =   -69400
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Oficina/Agencia:"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   14
         Left            =   -69160
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Remesa Id:"
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lista de Operaciones Pendientes de Traslado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   16
         Left            =   -69880
         TabIndex        =   50
         Top             =   1560
         Visible         =   0   'False
         Width           =   11412
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   -70000
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   11652
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -69880
         TabIndex        =   48
         Top             =   4560
         Visible         =   0   'False
         Width           =   5292
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesas - visualizar últimas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   -64600
         TabIndex        =   47
         Top             =   4560
         Visible         =   0   'False
         Width           =   5412
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   18
         Left            =   -62560
         TabIndex        =   46
         Top             =   6120
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   252
         Index           =   19
         Left            =   -62680
         TabIndex        =   45
         Top             =   6000
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total:"
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
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago de Comisiones de Cobros por Recuperación"
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
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   9612
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmCO_ControlComPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListViewItem, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim mConcepto As String


Private Sub sbFiltros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select fecha_inicio,fecha_corte from CBR_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!fecha_corte
rs.Close


'Cargado de Bancos

'strSQL = "Select B.cod_Banco as 'Idx',TB.descripcion as 'Itmx'" _
'         & " from afi_bene_pago B" _
'         & " inner join tes_bancos TB on B.cod_banco = TB.id_banco " _
'         & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio" _
'         & " and B.consec = O.consec and  registra_fecha between '" & Format(vFechaInicio, "yyyymmdd") & "  00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
'         & " where B.ESTADO = 'S' and B.tesoreria is null group by B.cod_Banco,TB.descripcion"
'Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
'
'
''Cargado de Usuarios
'strSQL = "Select O.Registra_User as 'IdX',O.Registra_User as 'Itmx'" _
'         & " from afi_bene_pago B" _
'         & " inner join tes_bancos TB on B.cod_banco = TB.id_banco " _
'         & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio" _
'         & " and B.consec = O.consec and  registra_fecha between '" & Format(vFechaInicio, "yyyymmdd") & "' and '" & Format(vFechaCorte, "yyyymmdd") & "'" _
'         & " where B.ESTADO = 'S' and B.tesoreria is null"
'Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboTraslado_Click()
    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
End Sub

Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
   fraFiltros.Visible = True
   Call sbFiltros
Else
   fraFiltros.Visible = False
End If
End Sub


Private Sub cboBanco_Click()
If vPaso Then Exit Sub
  lswCarga.ListItems.Clear
End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim i As Integer

On Error GoTo vError

Select Case Index
  Case 0 'NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(cod_remesa),0) + 1 as Ultimo from CBR_REMESAS"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert CBR_REMESAS(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa.Text = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa Comisiones de Cobros: " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update CBR_REMESAS set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa Comisiones de Cobros: " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case 1 'BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "delete CBR_REMESAS where Cod_Remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            Call Bitacora("Elimina", "Remesa Comisiones de Cobros: " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case 2 'REPORTES"
     fraReporte.Visible = Not fraReporte.Visible


  '---------Carga
  Case 3 'Carga: Buscar
    If cboCarga.ListCount = 0 Then Exit Sub
    Call sbCargaBuscar
  
  Case 4 'Carga: Cargar
    If lswCarga.ListItems.Count = 0 Then Exit Sub
    Call sbCarga
  
  Case 5 'Carga: Cerrar Remesa
    Call sbCerrar

  '---------Traslado
  Case 6 'Traslado: Buscar
    If cboTraslado.ListCount = 0 Then Exit Sub
    Call sbTrasladoBuscar
  
  Case 7 'Traslado: Traslado
    If cboTraslado.ListCount = 0 Then Exit Sub
    Call sbTraslado
  
  '---------Reportes
  Case 8
    Call sbInforme_Remesa

End Select


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboCarga_Click()
Dim vFechaInicio As Date, vFechaCorte As Date

lswCarga.ListItems.Clear

If vPaso Then Exit Sub
If cboCarga.ListCount <= 0 Then Exit Sub

vPaso = True
cboBanco.Clear


strSQL = "select fecha_inicio,fecha_corte from CBR_REMESAS where COD_REMESA  = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!fecha_corte
rs.Close

'Seleccionar Bancos
 
strSQL = "exec spCbrComision_ConsultaBancos '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' , '" & Format(vFechaCorte, "yyyy/mm/dd") & "  23:59:59'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboBanco.AddItem (Format(rs!id_Banco, "0000") & "..." & Trim(rs!BancoDesc))
  cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!id_Banco)
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboBanco.Text = (Format(rs!id_Banco, "0000") & "..." & Trim(rs!BancoDesc))
End If
rs.Close

cboBanco.AddItem "TODOS"
cboBanco.ItemData(cboBanco.ListCount - 1) = "TODOS"
cboBanco.Text = "TODOS"


vPaso = False
Call cboBanco_Click

End Sub


Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency

vPaso = True

For i = 1 To lswCarga.ListItems.Count
  lswCarga.ListItems.Item(i).Checked = chkCarga.Value
  
   If chkCarga.Value = vbChecked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(6))
   End If
  
Next i

vPaso = False

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub sbInforme_Remesa()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Cobros"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Planes")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     
     If chkDetalle.Value = vbChecked Then
         .ReportFileName = SIFGlobal.fxPathReportes("Cobro_RemesasDetalle_Full.rpt")
         vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO: DETALLADO"
     Else
         .ReportFileName = SIFGlobal.fxPathReportes("Cobro_RemesasDetalle.rpt")
         vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO: RESUMEN"
     End If
     
  
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     If chkDetalle.Value = vbChecked Then
         .ReportFileName = SIFGlobal.fxPathReportes("Cobro_RemesasDetalleAgrupado_Full.rpt")
         vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO: DETALLADO AGRUPADO"
     Else
         .ReportFileName = SIFGlobal.fxPathReportes("Cobro_RemesasDetalleAgrupado.rpt")
         vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO: RESUMEN AGRUPADO"
     End If
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE COBROS: PAGO DE COMISIONES'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{CBR_REMESAS.COD_REMESA} = " & lblRemesa.Tag
 .PrintReport


End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 4
End Sub

Private Sub Form_Load()

Dim strSQL As String

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 tcMain.Item(0).Selected = True
 
 With lswRemesas.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Estado", 1400
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With
 
 With lswRep.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With
  
 
 With lswCarga.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1500
    .Add , , "Identificación", 2100, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Mov.Inicio", 1500, vbCenter
    .Add , , "Mov.Corte", 1500, vbCenter
    .Add , , "Recuperado", 2100, vbRightJustify
    .Add , , "Comisión", 2100, vbRightJustify
    .Add , , "Banco", 2500
    .Add , , "Emite", 1000, vbCenter
    .Add , , "Cuenta", 2500

 End With
 
 
 With lswTraslado.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1500
    .Add , , "Nombre", 3500
    .Add , , "Banco", 2500
    .Add , , "Comisión", 2100, vbRightJustify
    .Add , , "Id Banco", 450
    .Add , , "Emite", 1000, vbCenter
    .Add , , "Cuenta", 2500
    .Add , , "Identificación", 2100, vbCenter
    .Add , , "Cta Banco", 2500
 End With
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 
strSQL = "select rtrim(cod_oficina) as 'Idx', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboRepOficina, strSQL, True, False)
 
End Sub


Private Sub sbConsulta(vRemesa As Long)

Call sbLimpia
  
strSQL = "select * from CBR_REMESAS where COD_REMESA  = " & vRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!cod_remesa
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "P"
      txtEstado = "Remesa en Proceso"
  End Select
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!fecha_corte
  
  txtNotas.Text = rs!notas
  
'  With glogon
'    .strSQL = "select isnull(sum(aportes_liq + rendi_liq),0) as Total from fnd_liquidacion" _
'            & " where consec in (select consec from fnd_remesa_asg where remesa = " & vRemesa & ")"
'    .Recordset.Open .strSQL, .Conection, adOpenStatic
'    txtTotal.Text = Format(.Recordset!Total, "Standard")
'    .Recordset.Close
'  End With
  
End If
rs.Close


End Sub


Private Sub lswCarga_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCarga.SortKey = ColumnHeader.Index - 1
  If lswCarga.SortOrder = 0 Then lswCarga.SortOrder = 1 Else lswCarga.SortOrder = 0
  lswCarga.Sorted = True
End Sub

Private Sub lswCarga_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curTotal As Currency

If vPaso Then Exit Sub

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(6))
Else
   curTotal = curTotal - CCur(Item.SubItems(6))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub


Private Sub lswRemesas_Click()
If lswRemesas.ListItems.Count <= 0 Then Exit Sub
Call sbConsulta(lswRemesas.SelectedItem)
End Sub



Private Sub sbReporte()
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Cobros"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
' .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Remesas.rpt")
' .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLimpia()

Me.MousePointer = vbHourglass

fraFiltros.Visible = False

Select Case tcMain.Selected.Index
  Case 0 'Remesas
     txtEstado = ""
     txtFecha = ""
     txtUsuario = ""
     txtRemesa = ""
     
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
    dtpRepInicio.Value = dtpInicio.Value
    dtpRepCorte.Value = dtpInicio.Value
    
    fraReporte.Visible = False
    
    txtNotas.Text = ""
     
     strSQL = "select TOP 50 * from CBR_REMESAS order by fecha desc"
     
     
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                
                Select Case rs!Estado
                  Case "A"
                     itmX.SubItems(3) = "Remesa Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Remesa Cerrada"
                  Case "T"
                     itmX.SubItems(3) = "Remesa Trasladada"
                End Select
                
                itmX.SubItems(4) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!fecha_corte, "dd/mm/yyyy")
                itmX.SubItems(6) = rs!notas
                
                
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear

    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from CBR_REMESAS where estado = 'A' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
      
      cboCarga.ItemData(cboCarga.ListCount - 1) = CStr(rs!cod_remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
    Call chkFiltros_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from CBR_REMESAS where estado = 'C' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.ListCount - 1) = CStr(rs!cod_remesa)
      
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!fecha_corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from CBR_REMESAS order by fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                itmX.SubItems(3) = rs!Fecha_Inicio
                itmX.SubItems(4) = rs!fecha_corte
                itmX.SubItems(5) = rs!notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close

    
 End Select


Me.MousePointer = vbDefault

End Sub



Private Sub sbCargaBuscar()
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0



strSQL = "select fecha_inicio,fecha_corte from CBR_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!fecha_corte
rs.Close

'Seleccionar x Bancos
strSQL = "exec spCbrComision_ConsultaPendientes '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00','" & Format(vFechaCorte, "yyyy/mm/dd") _
           & " 23:59:59'"
If cboBanco.Text <> "TODOS" Then
   strSQL = strSQL & "," & cboBanco.ItemData(cboBanco.ListIndex)
End If
           
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswCarga.ListItems.Add(, , rs!Usuario)
     itmX.SubItems(1) = rs!Cedula
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = rs!Recupera_Inicio
     itmX.SubItems(4) = rs!Recupera_Corte
     itmX.SubItems(5) = Format(rs!Recupera_Monto, "Standard")
     itmX.SubItems(6) = Format(rs!Comision, "Standard")
     itmX.SubItems(7) = rs!BancoDesc
     itmX.SubItems(8) = rs!TIPO_DOCUMENTO
     itmX.SubItems(9) = rs!Cuenta_Ahorros
     itmX.Checked = chkCarga.Value
     
     If itmX.Checked Then
        curTotal = curTotal + CCur(itmX.SubItems(6))
     End If
     
 rs.MoveNext
 
 PrgBar.Value = PrgBar.Value + 1
 
Loop
rs.Close

PrgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CBR_REMESAS" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','P')"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close



'Actualiza el Estado de la Remesa como cerrada
strSQL = "exec spCbrComision_RemesaCierra " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Genera", "Remesa de Fondos Remesa : " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CBR_REMESAS" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','P') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

'Calcula los casos a procesar
vCasos = 1
For i = 1 To lswCarga.ListItems.Count
 If lswCarga.ListItems.Item(i).Checked Then
    vCasos = vCasos + 1
 End If
Next i

PrgBar.Max = vCasos
PrgBar.Value = 1
PrgBar.Visible = True


With lswCarga.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
 
     strSQL = "exec spCbrComision_RemesaCarga " & cboCarga.ItemData(cboCarga.ListIndex) _
             & ",'" & .Item(i).Text & "'"
     Call ConectionExecute(strSQL)
   
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Cobros a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))
End If

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbCargaBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub




Private Sub sbTrasladoBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from CBR_REMESAS where cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!fecha_corte
rs.Close

strSQL = "select P.*, U.CEDULA, U.NOMBRE, B.DESCRIPCION as 'BancoDesc',B.CtaConta" _
       & " from CBR_REMESAS_PAGO P inner join CBR_USUARIOS U on P.USUARIO = U.USUARIO" _
       & " inner join Tes_Bancos B on P.COD_BANCO = B.ID_BANCO" _
       & " Where P.COD_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex) & " And P.TESORERIA_FECHA Is Null" _
       & " order by U.Nombre"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
Set itmX = .ListItems.Add(, , rs!Usuario)
       itmX.SubItems(1) = rs!Nombre
       itmX.SubItems(2) = rs!BancoDesc
       itmX.SubItems(3) = Format(rs!Comision, "Standard")
       
       itmX.SubItems(4) = rs!cod_banco
       itmX.SubItems(5) = rs!TIPO_EMISION
       itmX.SubItems(6) = rs!Cuenta_Ahorros & ""
       itmX.SubItems(7) = rs!Cedula
       itmX.SubItems(8) = rs!ctaConta
       
       itmX.Checked = vbChecked
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(3))
       End If
       
       rs.MoveNext
       PrgBar.Value = PrgBar.Value + 1
 Loop

End With

rs.Close

PrgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswTraslado.ListItems.Clear

End Sub


Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String) As Long                                 'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & mConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function

Private Sub sbTraslado()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngSolicitud As Long, vFecha As Date
Dim vTipo As String, vBanco As Integer
Dim i As Integer, vCasos As Integer

Dim vUnidad As String
Dim vCtaBanco As String, vCtaGasto As String
Dim vBeneficiario As String, vMonto As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

vCasos = 0
vFecha = fxFechaServidor
vUnidad = fxCBRParametro("24")
mConcepto = fxCBRParametro("26")
vCtaGasto = fxCBRParametro("02")

With lswTraslado.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
    vBeneficiario = .Item(i).SubItems(1)
    vMonto = .Item(i).SubItems(3)
    vBanco = .Item(i).SubItems(4)
    vTipo = .Item(i).SubItems(5)
    vCtaBanco = .Item(i).SubItems(8)
    
    lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, vMonto, .Item(i).SubItems(7) _
                    , vBeneficiario, 0, "Módulo de Cobros", 0 _
                   , "Pago de Comisiones: " & cboTraslado.ItemData(cboTraslado.ListIndex) _
                   , .Item(i).SubItems(6), vFecha, vUnidad)
    
    
    'Mata el Pasivo de la Nota de Debito de la Formalizacion contra Bancos
    Call sbCreaDetalle(lngSolicitud, vCtaBanco, vMonto, "H", 1, vUnidad)
    Call sbCreaDetalle(lngSolicitud, vCtaGasto, vMonto, "D", 2, vUnidad)

    'Actualiza Campo Tesoreria
    strSQL = "update CBR_REMESAS_PAGO set tesoreria_nSolicitud = " & lngSolicitud _
           & ", tesoreria_fecha = dbo.MyGetdate(), tesoreria_Usuario = '" & glogon.Usuario _
           & "' where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
           & "  and Usuario = '" & .Item(i).Text & "'"
    Call ConectionExecute(strSQL)

 'Actualiza Bitacora
    Call Bitacora("Registra", "Traspaso de Remesa de Cobros a Tesoreria:" & .Item(i).Text)
 
   
   ' PrgBar.Value = PrgBar.Value + 1
    vCasos = vCasos + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Traslado de Remesa de Cobros a Tesoreria: " & cboTraslado.ItemData(cboTraslado.ListIndex))
    'Actualiza y Carga Remesa
    strSQL = "update CBR_REMESAS SET Estado = 'T'" _
           & "  Where cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
    Call ConectionExecute(strSQL)
End If
End With



Call sbLimpia


Me.MousePointer = vbDefault

PrgBar.Visible = False

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswRemesas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRemesas.SortKey = ColumnHeader.Index - 1
  If lswRemesas.SortOrder = 0 Then lswRemesas.SortOrder = 1 Else lswRemesas.SortOrder = 0
  lswRemesas.Sorted = True

End Sub






Private Sub lswRep_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = Item.Text & " ¦ " & Item.SubItems(1) _
            & " ¦ " & Item.SubItems(2)
lblRemesa.Tag = Item.Text
End Sub

Private Sub lswTraslado_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswTraslado.SortKey = ColumnHeader.Index - 1
  If lswTraslado.SortOrder = 0 Then lswTraslado.SortOrder = 1 Else lswTraslado.SortOrder = 0
  lswTraslado.Sorted = True
End Sub

Private Sub lswTraslado_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curTotal As Currency

If vPaso Then Exit Sub

If Trim(txtPagoTotal.Text) = "" Then txtPagoTotal.Text = 0

curTotal = CCur(txtPagoTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(3))
Else
   curTotal = curTotal - CCur(Item.SubItems(3))
End If

txtPagoTotal.Text = Format(curTotal, "Standard")

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call sbLimpia
End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If
End Sub


Private Sub txtRetiro_Change()
' txtConRemesa.Text = ""
End Sub


Private Sub sbConsultaRetiro()

On Error GoTo vError


'strSQL = "select A.* from CBR_REMESAS A inner join fnd_remesa_asg X on A.remesa = X.remesa where consec = " & txtRetiro.Text
'Call OpenRecordSet(rs, strSQL)
'If rs.BOF Or rs.EOF Then
' txtConRemesa.Text = "** No se encontró retiro/liq. en las remesas registradas **"
'Else
' txtConRemesa.Text = "Remesa   " & vbTab & " ...:" & rs!cod_remesa & vbCrLf
' txtConRemesa.Text = txtConRemesa & "Fecha   " & vbTab & " ...:" & rs!fecha & vbCrLf
' txtConRemesa.Text = txtConRemesa & "Usuario  " & vbTab & " ...:" & rs!Usuario
'End If
'rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
' txtConRemesa.Text = ""

End Sub

Private Sub txtRetiro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaRetiro
End Sub


