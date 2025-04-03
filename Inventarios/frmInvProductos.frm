VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmInvProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11160
   Begin XtremeSuiteControls.FlatEdit txtModFecha 
      Height          =   285
      Left            =   8040
      TabIndex        =   102
      ToolTipText     =   "Modifica Fecha"
      Top             =   7680
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   10935
      _Version        =   1572864
      _ExtentX        =   19288
      _ExtentY        =   11668
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
      ItemCount       =   8
      Item(0).Caption =   "General"
      Item(0).ControlCount=   55
      Item(0).Control(0)=   "cboLineaSub"
      Item(0).Control(1)=   "cboTipo"
      Item(0).Control(2)=   "txtCodigoBarras"
      Item(0).Control(3)=   "txtLineaCod"
      Item(0).Control(4)=   "txtUnidadCod"
      Item(0).Control(5)=   "txtMarcaCod"
      Item(0).Control(6)=   "txtObservacion"
      Item(0).Control(7)=   "txtExistenciaMin"
      Item(0).Control(8)=   "txtExistenciaMax"
      Item(0).Control(9)=   "txtComMonto"
      Item(0).Control(10)=   "txtComUnidad"
      Item(0).Control(11)=   "txtCodFabricante"
      Item(0).Control(12)=   "txtModelo"
      Item(0).Control(13)=   "chkActivo"
      Item(0).Control(14)=   "txtLineaDesc"
      Item(0).Control(15)=   "txtUnidadDesc"
      Item(0).Control(16)=   "txtMarcaDesc"
      Item(0).Control(17)=   "imgBarra"
      Item(0).Control(18)=   "Label2(12)"
      Item(0).Control(19)=   "Label2(11)"
      Item(0).Control(20)=   "Label4(2)"
      Item(0).Control(21)=   "Label4(1)"
      Item(0).Control(22)=   "Label2(5)"
      Item(0).Control(23)=   "Label2(4)"
      Item(0).Control(24)=   "Label2(3)"
      Item(0).Control(25)=   "Label2(2)"
      Item(0).Control(26)=   "Label4(0)"
      Item(0).Control(27)=   "Label3(0)"
      Item(0).Control(28)=   "Label2(1)"
      Item(0).Control(29)=   "Label2(0)"
      Item(0).Control(30)=   "chkExistencia"
      Item(0).Control(31)=   "txtImpConsumo"
      Item(0).Control(32)=   "txtImpVentas"
      Item(0).Control(33)=   "txtCostoRegular"
      Item(0).Control(34)=   "txtUtilidadGeneral"
      Item(0).Control(35)=   "txtPrecioRegular"
      Item(0).Control(36)=   "chkLotes"
      Item(0).Control(37)=   "Label2(10)"
      Item(0).Control(38)=   "Label2(9)"
      Item(0).Control(39)=   "Label5(0)"
      Item(0).Control(40)=   "Label5(1)"
      Item(0).Control(41)=   "Label2(6)"
      Item(0).Control(42)=   "txtCabys"
      Item(0).Control(43)=   "Label2(7)"
      Item(0).Control(44)=   "Label3(1)"
      Item(0).Control(45)=   "btnCabys"
      Item(0).Control(46)=   "chkStock"
      Item(0).Control(47)=   "chkVentaEnLinea"
      Item(0).Control(48)=   "txtPrecioRegularIVA"
      Item(0).Control(49)=   "Label2(8)"
      Item(0).Control(50)=   "Label7(0)"
      Item(0).Control(51)=   "txtV_ItemsCantidad"
      Item(0).Control(52)=   "txtV_ItemsFrecuencia"
      Item(0).Control(53)=   "Label8"
      Item(0).Control(54)=   "Label7(1)"
      Item(1).Caption =   "Precios"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "txtPrecioCod"
      Item(1).Control(1)=   "txtPrecioDesc"
      Item(1).Control(2)=   "txtPrecio"
      Item(1).Control(3)=   "txtUtilidad"
      Item(1).Control(4)=   "lswPrecios"
      Item(2).Caption =   "Desc/Bonif."
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "vGridBon"
      Item(2).Control(1)=   "vGridDes"
      Item(2).Control(2)=   "Label5(3)"
      Item(2).Control(3)=   "Label5(2)"
      Item(3).Caption =   "Proveedores"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswProv"
      Item(4).Caption =   "Existencias"
      Item(4).ControlCount=   5
      Item(4).Control(0)=   "Label9(0)"
      Item(4).Control(1)=   "dtpInvCorte"
      Item(4).Control(2)=   "btnExistencia(0)"
      Item(4).Control(3)=   "btnExistencia(1)"
      Item(4).Control(4)=   "lswExistencia"
      Item(5).Caption =   "Barras"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "gbBarrasPrinter(0)"
      Item(5).Control(1)=   "gbBarrasPrinter(1)"
      Item(6).Caption =   "Movimientos"
      Item(6).ControlCount=   2
      Item(6).Control(0)=   "lswMov"
      Item(6).Control(1)=   "ShortcutCaption1(0)"
      Item(7).Caption =   "Similares"
      Item(7).ControlCount=   9
      Item(7).Control(0)=   "btnSimilar(0)"
      Item(7).Control(1)=   "btnSimilar(1)"
      Item(7).Control(2)=   "txtSimDesc"
      Item(7).Control(3)=   "txtSimCodigo"
      Item(7).Control(4)=   "lswSim"
      Item(7).Control(5)=   "txtSimCabys"
      Item(7).Control(6)=   "ShortcutCaption1(1)"
      Item(7).Control(7)=   "ShortcutCaption1(2)"
      Item(7).Control(8)=   "ShortcutCaption1(3)"
      Begin XtremeSuiteControls.ListView lswSim 
         Height          =   5295
         Left            =   -70000
         TabIndex        =   85
         Top             =   1200
         Visible         =   0   'False
         Width           =   10815
         _Version        =   1572864
         _ExtentX        =   19076
         _ExtentY        =   9340
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswMov 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1572864
         _ExtentX        =   18865
         _ExtentY        =   10186
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswExistencia 
         Height          =   6135
         Left            =   -67480
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   8055
         _Version        =   1572864
         _ExtentX        =   14208
         _ExtentY        =   10821
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswProv 
         Height          =   6015
         Left            =   -69640
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   10610
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswPrecios 
         Height          =   5655
         Left            =   -69040
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   9975
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSimCabys 
         Height          =   312
         Left            =   -62320
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSimDesc 
         Height          =   312
         Left            =   -67960
         TabIndex        =   87
         Top             =   840
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1572864
         _ExtentX        =   9970
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbBarrasPrinter 
         Height          =   2892
         Index           =   0
         Left            =   -64240
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   4572
         _Version        =   1572864
         _ExtentX        =   8064
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Impresión de Códigos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnCodigoBarras 
            Height          =   312
            Index           =   2
            Left            =   2400
            TabIndex        =   35
            Top             =   960
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Imprimir"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtBarCopias 
            Height          =   312
            Left            =   1800
            TabIndex        =   37
            Top             =   480
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Text            =   "1"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   36
            Top             =   480
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. de Copias:"
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
      Begin XtremeSuiteControls.PushButton btnExistencia 
         Height          =   312
         Index           =   0
         Left            =   -68680
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Consulta"
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
      End
      Begin FPSpreadADO.fpSpread vGridBon 
         Height          =   2412
         Left            =   -67360
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   5892
         _Version        =   524288
         _ExtentX        =   10393
         _ExtentY        =   4254
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvProductos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridDes 
         Height          =   3015
         Left            =   -67360
         TabIndex        =   6
         Top             =   3480
         Visible         =   0   'False
         Width           =   5895
         _Version        =   524288
         _ExtentX        =   10398
         _ExtentY        =   5318
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvProductos.frx":060B
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPrecioCod 
         Height          =   312
         Left            =   -69040
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "..."
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPrecioDesc 
         Height          =   312
         Left            =   -67600
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
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
         Text            =   "..."
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPrecio 
         Height          =   312
         Left            =   -62320
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUtilidad 
         Height          =   312
         Left            =   -60880
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInvCorte 
         Height          =   312
         Left            =   -68920
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
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
      Begin XtremeSuiteControls.PushButton btnExistencia 
         Height          =   312
         Index           =   1
         Left            =   -68680
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Informe"
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
      End
      Begin XtremeSuiteControls.GroupBox gbBarrasPrinter 
         Height          =   5535
         Index           =   1
         Left            =   -69760
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   9763
         _StockProps     =   79
         Caption         =   "Definición del Código de Barras"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin VB.PictureBox picEan 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawWidth       =   2
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   240
            ScaleHeight     =   82.177
            ScaleMode       =   0  'User
            ScaleWidth      =   114.632
            TabIndex        =   30
            Top             =   3120
            Width           =   3372
            Begin VB.VScrollBar Scroll1 
               Height          =   1755
               LargeChange     =   5
               Left            =   3060
               TabIndex        =   31
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtBarGenera 
            Height          =   312
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "Enter 12 digits"
            Top             =   480
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Text            =   "859200600"
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCodigoBarras 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   33
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Generar"
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
         End
         Begin XtremeSuiteControls.PushButton btnCodigoBarras 
            Height          =   315
            Index           =   1
            Left            =   3480
            TabIndex        =   34
            Top             =   480
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Registrar"
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
         End
         Begin VB.Label lbCo1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCol/2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   2040
            TabIndex        =   29
            Top             =   2520
            Width           =   852
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Check Digit:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   2760
            Width           =   1392
         End
         Begin VB.Label lbCo1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCol/3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   2040
            TabIndex        =   27
            Top             =   2760
            Width           =   852
         End
         Begin VB.Label lbCo1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCol/1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   2040
            TabIndex        =   26
            Top             =   2280
            Width           =   852
         End
         Begin VB.Label lbCo1 
            BackStyle       =   0  'Transparent
            Caption         =   "lblCol/0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   2040
            TabIndex        =   25
            Top             =   2040
            Width           =   852
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Código Producto:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   2520
            Width           =   1692
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor / Línea:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   23
            Top             =   2280
            Width           =   1692
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Prefijo (ICN):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   22
            Top             =   2040
            Width           =   1392
         End
      End
      Begin XtremeSuiteControls.PushButton btnSimilar 
         Height          =   312
         Index           =   0
         Left            =   -60040
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   372
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
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
         Appearance      =   16
         Picture         =   "frmInvProductos.frx":0C19
      End
      Begin XtremeSuiteControls.PushButton btnSimilar 
         Height          =   312
         Index           =   1
         Left            =   -59680
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   372
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
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
         Appearance      =   16
         Picture         =   "frmInvProductos.frx":1339
      End
      Begin XtremeSuiteControls.ComboBox cboLineaSub 
         Height          =   312
         Left            =   3240
         TabIndex        =   41
         Top             =   1200
         Width           =   6852
         _Version        =   1572864
         _ExtentX        =   12091
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   1320
         TabIndex        =   42
         Top             =   480
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigoBarras 
         Height          =   312
         Left            =   4440
         TabIndex        =   43
         Top             =   480
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLineaCod 
         Height          =   312
         Left            =   1320
         TabIndex        =   44
         Top             =   840
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadCod 
         Height          =   312
         Left            =   1680
         TabIndex        =   45
         Top             =   1920
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMarcaCod 
         Height          =   312
         Left            =   1680
         TabIndex        =   46
         Top             =   2280
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   792
         Left            =   1680
         TabIndex        =   47
         Top             =   2640
         Width           =   8412
         _Version        =   1572864
         _ExtentX        =   14838
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.FlatEdit txtExistenciaMin 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   48
         Top             =   3720
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtExistenciaMax 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   49
         Top             =   4080
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComMonto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   50
         Top             =   4440
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComUnidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   51
         Top             =   4800
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodFabricante 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   52
         Top             =   5160
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtModelo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   2640
         TabIndex        =   53
         Top             =   5520
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   252
         Left            =   8760
         TabIndex        =   54
         Top             =   480
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activo? "
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
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
         Height          =   312
         Left            =   3240
         TabIndex        =   55
         Top             =   840
         Width           =   6852
         _Version        =   1572864
         _ExtentX        =   12086
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
         Height          =   312
         Left            =   3360
         TabIndex        =   56
         Top             =   1920
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMarcaDesc 
         Height          =   312
         Left            =   3360
         TabIndex        =   57
         Top             =   2280
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkExistencia 
         Height          =   255
         Left            =   8160
         TabIndex        =   70
         Top             =   6000
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Controla Existencia?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtImpConsumo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   71
         Top             =   3720
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImpVentas 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   72
         Top             =   4080
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCostoRegular 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   73
         Top             =   4440
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUtilidadGeneral 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   74
         Top             =   4800
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPrecioRegular 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   75
         Top             =   5160
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkLotes 
         Height          =   255
         Left            =   8160
         TabIndex        =   76
         Top             =   6360
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Utiliza Lotes?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtCabys 
         Height          =   312
         Left            =   8040
         TabIndex        =   82
         Top             =   1560
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSimCodigo 
         Height          =   312
         Left            =   -70000
         TabIndex        =   86
         Top             =   840
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCabys 
         Height          =   312
         Left            =   10080
         TabIndex        =   94
         ToolTipText     =   "Heredar CABYS de la Sub/Linea"
         Top             =   1560
         Width           =   372
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
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
         Appearance      =   16
         Picture         =   "frmInvProductos.frx":18DD
      End
      Begin XtremeSuiteControls.CheckBox chkStock 
         Height          =   255
         Left            =   5760
         TabIndex        =   95
         Top             =   6000
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Articulo de Stock?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkVentaEnLinea 
         Height          =   255
         Left            =   5760
         TabIndex        =   96
         Top             =   6360
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activa Venta en Línea?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtPrecioRegularIVA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   97
         Top             =   5520
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtV_ItemsCantidad 
         Height          =   315
         Left            =   2640
         TabIndex        =   103
         Top             =   6240
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtV_ItemsFrecuencia 
         Height          =   315
         Left            =   3600
         TabIndex        =   104
         Top             =   6240
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   107
         Top             =   5880
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Frecuencia en días"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   495
         Left            =   360
         TabIndex        =   106
         Top             =   6000
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Ítems permitidos para la venta por persona"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   105
         Top             =   6000
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cantidad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Regular (Con IVA)"
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
         Index           =   8
         Left            =   5760
         TabIndex        =   98
         Top             =   5520
         Width           =   2535
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   252
         Index           =   3
         Left            =   -62320
         TabIndex        =   92
         Top             =   600
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   14
         Caption         =   "Cabys"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   252
         Index           =   2
         Left            =   -67960
         TabIndex        =   91
         Top             =   600
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1572864
         _ExtentX        =   9970
         _ExtentY        =   444
         _StockProps     =   14
         Caption         =   "Descripción"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   252
         Index           =   1
         Left            =   -70000
         TabIndex        =   90
         Top             =   600
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   14
         Caption         =   "Código"
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
         Alignment       =   1
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub/Linea"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   89
         Top             =   1200
         Width           =   1332
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   10692
         _Version        =   1572864
         _ExtentX        =   18860
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Ultimos Movimientos Recibidos por este Producto"
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
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabys"
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
         Index           =   7
         Left            =   7080
         TabIndex        =   83
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo Regular"
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
         Index           =   6
         Left            =   5760
         TabIndex        =   81
         Top             =   4440
         Width           =   3252
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "% I.V.A."
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
         Index           =   1
         Left            =   5760
         TabIndex        =   80
         Top             =   4080
         Width           =   2412
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "% I. Servicio/Consumo "
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
         Index           =   0
         Left            =   5760
         TabIndex        =   79
         Top             =   3720
         Width           =   2532
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Regular (Sin IVA)"
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
         Index           =   9
         Left            =   5760
         TabIndex        =   78
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Margen Utilidad"
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
         Index           =   10
         Left            =   5760
         TabIndex        =   77
         Top             =   4800
         Width           =   2532
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Barras"
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
         Index           =   0
         Left            =   3600
         TabIndex        =   69
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo P/S"
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
         Index           =   1
         Left            =   240
         TabIndex        =   68
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea"
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
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad "
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
         Index           =   0
         Left            =   360
         TabIndex        =   66
         Top             =   1920
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código del Fabricante"
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
         Index           =   2
         Left            =   360
         TabIndex        =   65
         Top             =   5160
         Width           =   2532
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia Mínima"
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
         Index           =   3
         Left            =   360
         TabIndex        =   64
         Top             =   3720
         Width           =   2292
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión sobre Monto"
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
         Index           =   4
         Left            =   360
         TabIndex        =   63
         Top             =   4440
         Width           =   2292
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión por Unidades"
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
         Index           =   5
         Left            =   360
         TabIndex        =   62
         Top             =   4800
         Width           =   2652
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Index           =   1
         Left            =   360
         TabIndex        =   61
         Top             =   2640
         Width           =   1212
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
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
         Index           =   2
         Left            =   360
         TabIndex        =   60
         Top             =   2280
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo "
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
         Index           =   11
         Left            =   360
         TabIndex        =   59
         Top             =   5520
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia Maxima"
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
         Index           =   12
         Left            =   360
         TabIndex        =   58
         Top             =   4080
         Width           =   2292
      End
      Begin VB.Image imgBarra 
         Height          =   252
         Left            =   6540
         Picture         =   "frmInvProductos.frx":1FFD
         Stretch         =   -1  'True
         ToolTipText     =   "Crear Código de Barras EAN"
         Top             =   480
         Width           =   252
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   252
         Index           =   0
         Left            =   -69760
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   732
         _Version        =   1572864
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Corte"
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bonificaciones por Unidades Vendidas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   852
         Index           =   2
         Left            =   -69400
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje de Descuentos por Monto de Ventas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1092
         Index           =   3
         Left            =   -69280
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Boleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repSep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "InventarioGeneral"
                  Text            =   "Inventario General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10320
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11874
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRegUsuario 
      Height          =   285
      Left            =   480
      TabIndex        =   99
      ToolTipText     =   "Registro Usuario"
      Top             =   7680
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin XtremeSuiteControls.FlatEdit txtRegFecha 
      Height          =   285
      Left            =   3000
      TabIndex        =   100
      ToolTipText     =   "Fecha Registro"
      Top             =   7680
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin XtremeSuiteControls.FlatEdit txtModUsuario 
      Height          =   285
      Left            =   5520
      TabIndex        =   101
      ToolTipText     =   "Modifica Usuario"
      Top             =   7680
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin XtremeSuiteControls.Label Label6 
      Height          =   250
      Left            =   240
      TabIndex        =   93
      Top             =   480
      Width           =   972
      _Version        =   1572864
      _ExtentX        =   1714
      _ExtentY        =   441
      _StockProps     =   79
      Caption         =   "Producto"
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
Attribute VB_Name = "frmInvProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean
Dim vPaso As Boolean


Private Sub btnCabys_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

  strSQL = "select Cabys from pv_prod_clasifica_Sub " _
         & "where COD_PRODCLAS = " & txtLineaCod.Text _
         & " and COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'"
  Call OpenRecordSet(rs, strSQL)

  If (rs!Cabys & "") <> "" Then
      txtCabys.Text = rs!Cabys & ""
  End If
  rs.Close

Exit Sub

vError:

End Sub

Private Sub btnCodigoBarras_Click(Index As Integer)

Select Case Index
    Case 0 'Genera
        Call sbBarras_Genera
    Case 1 'Registra
        Call sbBarras_Registra
    Case 2 'Imprime
        Call sbBarras_Imprime
End Select

End Sub

Private Sub btnExistencia_Click(Index As Integer)

Select Case Index
Case 0 'Consulta
    Call sbInvCorte_Consulta
Case 1 'Informe
    Call sbListViewExporFileTab(lswExistencia)
End Select
End Sub

Private Sub btnSimilar_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSimilar As Long, vSimilarActual As Long, vSimNuevo As Long

On Error GoTo vError

If txtSimCodigo.Text = "" Then Exit Sub


Select Case Index
  Case 0 'Agregar


            strSQL = "select isnull(similar,0) as Similar from pv_productos where cod_producto = '" _
                   & vCodigo & "'"
            Call OpenRecordSet(rs, strSQL)
             vSimilar = rs!similar
            rs.Close
            
            strSQL = "select isnull(similar,0) as Similar from pv_productos where cod_producto = '" _
                   & txtSimCodigo & "'"
            Call OpenRecordSet(rs, strSQL)
             vSimNuevo = rs!similar
            rs.Close
            
            
            If vSimilar = 0 And vSimNuevo > 0 Then
               strSQL = "update pv_productos set similar = " & vSimNuevo & " where cod_producto = '" & vCodigo & "'"
               Call ConectionExecute(strSQL)
            End If
            
            
            If vSimilar > 0 And vSimNuevo = 0 Then
               strSQL = "update pv_productos set similar = " & vSimilar & " where cod_producto = '" & txtSimCodigo & "'"
               Call ConectionExecute(strSQL)
            End If
            
            If vSimilar > 0 And vSimNuevo > 0 Then
               strSQL = "update pv_productos set similar = " & vSimilar & " where cod_producto = '" & txtSimCodigo & "'"
               Call ConectionExecute(strSQL)
            End If
            
            If vSimilar = 0 And vSimNuevo = 0 Then
                strSQL = "select max(isnull(similar,0)) + 1 as Similar from pv_productos"
                Call OpenRecordSet(rs, strSQL)
                 vSimNuevo = rs!similar
                rs.Close
               
                strSQL = "update pv_productos set similar = " & vSimNuevo & " where cod_producto = '" & vCodigo & "'"
                Call ConectionExecute(strSQL)
            
                strSQL = "update pv_productos set similar = " & vSimNuevo & " where cod_producto = '" & txtSimCodigo & "'"
                Call ConectionExecute(strSQL)
            
            End If

  Case 1 'Elimina
                strSQL = "update pv_productos set similar = null where cod_producto = '" & txtSimCodigo & "'"
                Call ConectionExecute(strSQL)
End Select



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboLineaSub_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadCod.SetFocus

End Sub

Private Sub cboTipo_Click()
If Mid(cboTipo.Text, 1, 1) = "P" Then
  chkExistencia.Value = vbChecked
Else
  chkExistencia.Value = vbUnchecked
End If

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkActivo.SetFocus
End Sub


Private Sub chkExistencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComMonto.SetFocus
End Sub

Private Sub chkActivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtLineaCod.SetFocus
End Sub


Private Sub sbDibujaBarras()
Dim i As Integer, m As Integer, d As Integer, B As Integer, A As Integer
Dim lngX As Long
With picEan
    .Cls
    .BackColor = vbWhite
    .FontSize = 8
    .DrawWidth = 2
    lngX = 11  ' 11       'X position (11 =must be 11 modules [1 module = usually 0.33 millimeters, in my case picEan.ScaleWidth <bar width> / 113] 11 on left side, 7 on right side
    For i = 1 To 14 '13 digits :-)
        d = CInt(Right(Left(txtBarGenera, i), 1))  'Current n°
        If i = 1 Or i = 14 Then             'Draw the guard bars at the begining and end
            picEan.Line (lngX, 5)-(lngX, 52)
            picEan.Line (lngX + 2, 5)-(lngX + 2, 52)
            lngX = lngX + 3
            If i = 1 Then                   'Print first digit
                .CurrentX = 1
                .CurrentY = 44
                picEan.Print d
                B = d                       'Store inf. what's the first n° for the module algorithm
            End If
        Else
            If i < 8 Then                   'On the left side, there are modules 1 or 2 (A, B) depending on the 1st digit = [Mdl(0 - 9, 0 or 1)]...
                A = CInt(Right(Left(MdlLeft(B), i - 1), 1))
            Else: A = 2                     '...on the right side always module 2 (C) = [Mdl(0 - 9, 2)]
            End If
            If i = 8 Then                   'Draw the centre pattern
                picEan.Line (lngX + 1, 5)-(lngX + 1, 52)
                picEan.Line (lngX + 3, 5)-(lngX + 3, 52)
                lngX = lngX + 5
            End If
            For m = 1 To 7                  '7 modules for each n° (System of 7 black or white sprites)
                If CInt(Right(Left(Mdl(A, d), m), 1)) = 1 Then picEan.Line (lngX, 5)-(lngX, 44) 'Draw modules(sprites) for each n°
                lngX = lngX + 1
            Next m
            .CurrentX = lngX - 8
            .CurrentY = 44
            picEan.Print d                  'Print n°s
        End If
    Next i
End With
End Sub




Private Sub sbBarras_Registra()
Dim strSQL As String, i As Integer

txtBarGenera.Text = "2000" & Mid(Format(Trim(txtLineaCod), "000"), 1, 3) & Mid(Format(Trim(txtCodigo), "00000"), 1, 5)

i = MsgBox("Desea reemplazar el código de Barras Actual por este : " & txtBarGenera, vbYesNo)

If i = vbYes Then
  strSQL = "update pv_productos set cod_barras = '" & txtBarGenera _
         & "' where cod_producto = '" & txtCodigo & "'"
  Call ConectionExecute(strSQL)
  MsgBox "Codigo Reemplazado...", vbInformation
End If

End Sub

Private Sub sbBarras_Genera()
Dim i As Integer, m As Integer, n As Integer, S As Integer

On Error GoTo vError

If Len(txtBarGenera) = 12 Then         'If there's enough digitts for algorithm
    For i = 1 To 11 Step 2      'Sum every number at even position
        m = m + CInt(Right(Left(txtBarGenera, i), 1))
    Next i
    For i = 2 To 12 Step 2      'Sum every n° at odd position
        n = n + CInt(Right(Left(txtBarGenera, i), 1))
    Next i
    S = 10 - ((n * 3 + m) Mod 10) 'Count the Check digit (s = a number to the nierest multiplicand by 10 from n * 3 + m
    If S = 10 Then S = 0
    
    lbCo1(0) = Left(txtBarGenera, 1)     'International country number
    
    lbCo1(1) = Mid(txtBarGenera, 2, 3) & " / " & Mid(txtBarGenera, 5, 3) ' Left(Right(txtBarGenera, Len(txtBarGenera) - 1), 6)    'Manufacturer n° (in my case there are 6 digits,- must be from 4 to 6)
    lbCo1(2) = Right(txtBarGenera, 5)    'Item reference digit (In my case there are 3 n°s, must be from 3 to 5
    lbCo1(3) = S                  'Check digit
    txtBarGenera = txtBarGenera & lbCo1(3)      'Full EAN code
    txtBarGenera.SelStart = Len(txtBarGenera) - 1 'make the cursor apear where it's been
 '   txtBarGenera.SelLength = 1
    If Scroll1.Visible Then Scroll1.Visible = False
    sbDibujaBarras
Else: MsgBox "Ingrese 12 numeros en la casilla de barras!", vbExclamation, App.Title
End If

vError:
    If Err.Number = 13 Then MsgBox "Ingrese Solo Numeros en la casilla de barras!", vbExclamation, App.Title    'In case someone puts other characters then numbers into textbox

End Sub


Private Sub sbBarras_Imprime()
Dim i As Integer, x As Integer, y As Integer

On Error GoTo vError

y = txtBarCopias


If picEan.BackColor = vbWhite Then
 Me.MousePointer = vbHourglass
  For x = 1 To y
   Call sbgBarrasPrintCod(picEan)
  Next x
 Me.MousePointer = vbDefault
Else
  MsgBox "No se Imprimio Codigo de Barras!", vbExclamation, App.Title
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbInvCorte_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curSum As Currency


lswExistencia.ListItems.Clear
curSum = 0

strSQL = "select cod_bodega,descripcion from pv_bodegas "
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswExistencia.ListItems.Add(, , rs!cod_bodega)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = Format(fxInvProcesoProd(txtCodigo, rs!cod_bodega, dtpInvCorte.Value), "###,###,###,##0")
      curSum = curSum + CCur(itmX.SubItems(2))
  rs.MoveNext
Loop
rs.Close
  

Set itmX = lswExistencia.ListItems.Add(, , "")
    itmX.SubItems(2) = "______________"

Set itmX = lswExistencia.ListItems.Add(, , "Total")
    itmX.SubItems(2) = Format(curSum, "###,###,###,##0")

If curSum > CCur(txtExistenciaMin) Then
   itmX.ForeColor = vbBlue
 Else
   itmX.ForeColor = vbRed
End If



End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_producto from pv_productos" _
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_producto > '" & txtCodigo & "' order by cod_producto asc"
    Else
       strSQL = strSQL & " where cod_producto < '" & txtCodigo & "' order by cod_producto desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!Cod_Producto)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 32
 
 vGridBon.AppearanceStyle = fxGridStyle
 vGridDes.AppearanceStyle = fxGridStyle


 cboTipo.Clear
 cboTipo.AddItem "Producto"
 cboTipo.AddItem "Servicio"
 cboTipo.AddItem "Activo Fijo"
 cboTipo.Text = "Producto"

 With lswSim.ColumnHeaders
    .Clear
    .Add , , "Código", 2500
    .Add , , "Descripción", 4800
    .Add , , "Cabys", 2500, vbCenter
End With

 With lswPrecios.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", 3800
    .Add , , "Precio", 2100, vbRightJustify
    .Add , , "Utilidad", 2100, vbRightJustify
 End With

 With lswProv.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", 4200
    .Add , , "Identificación", 2200, vbCenter
    .Add , , "Ult.Factura", 2200, vbCenter
 End With

 With lswExistencia.ColumnHeaders
    .Clear
    .Add , , "Bodega", 1200
    .Add , , "Descripción", 4200
    .Add , , "Existencia", 2200, vbCenter
 End With


 With lswMov.ColumnHeaders
    .Clear
    .Add , , "Origen", 1200
    .Add , , "Comprobante", 2100
    .Add , , "Fecha Aplica", 1600
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Cantidad", 1200, vbRightJustify
    .Add , , "Usuario", 1200
    .Add , , "Fecha System", 2500
    .Add , , "Bodega Origen", 2500
    .Add , , "Bodega Destino", 2500
 End With

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
  
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

'Activa Tabs
For i = 1 To tcMain.ItemCount - 1
  tcMain.Item(i).Enabled = True
Next i


vCodigo = ""
txtCodigo = ""

txtNombre = ""

txtCodigoBarras = ""
cboTipo.Text = "Producto"
chkActivo.Value = vbChecked

txtLineaCod = ""
txtLineaDesc = ""
txtUnidadCod = ""
txtUnidadDesc = ""
txtMarcaCod = ""
txtMarcaDesc = ""
txtObservacion = ""

txtCodFabricante = ""
txtModelo = ""
txtExistenciaMin.Text = "0"
txtExistenciaMax.Text = "0"

chkExistencia.Value = vbChecked
chkLotes.Value = xtpUnchecked

chkStock.Value = xtpChecked
chkVentaEnLinea.Value = xtpUnchecked


txtComMonto = "0"
txtComUnidad = "0"
txtCostoRegular = "0"
txtPrecioRegular = "0"
txtPrecioRegularIVA.Text = "0"
txtImpConsumo = "0"
txtImpVentas = "0"

txtUtilidadGeneral.Text = "0"

txtV_ItemsCantidad.Text = "100"
txtV_ItemsFrecuencia.Text = "1"

cboLineaSub.Clear

dtpInvCorte.Value = fxFechaServidor

txtRegFecha.Text = ""
txtRegUsuario.Text = ""

txtModFecha.Text = ""
txtModUsuario.Text = ""


tcMain.Item(0).Selected = True
For i = 1 To tcMain.ItemCount - 1
  tcMain.Item(i).Enabled = False
Next i


End Sub


Private Sub imgBarra_Click()
Dim strSQL As String, i As Integer

txtCodigoBarras = "2000" & Mid(Format(Trim(txtLineaCod), "000"), 1, 3) & Mid(Format(Trim(txtCodigo), "00000"), 1, 5)

End Sub




Private Sub lswPrecios_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

  txtPrecioCod.Text = Item.Text
  txtPrecioDesc.Text = Item.SubItems(1)
  txtPrecio.Text = Item.SubItems(2)
  txtUtilidad.Text = Item.SubItems(3)
  txtPrecio.SetFocus

End Sub


Private Sub lswProv_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
   strSQL = "insert pv_producto_prov(cod_producto,cod_proveedor) values('" _
          & vCodigo & "'," & Item.Text & ")"
Else
   strSQL = "delete pv_producto_prov where cod_producto = '" & vCodigo _
          & "' and cod_proveedor = " & Item.Text
End If
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbMovimientos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX  As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

With lswMov
 .ListItems.Clear
 'Ultima Entrada
 strSQL = "select Top 1 T.Tipo,T.Boleta,T.fecha,T.procesa_user,T.procesa_fecha" _
        & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
        & ",rtrim(D.cod_bodega_destino) + ' - ' + rtrim(X.descripcion) as BodegaD" _
        & ",D.precio,D.cantidad" _
        & " from pv_invTransac T inner join pv_invTraDet D" _
        & " on T.tipo = D.tipo and T.boleta = D.boleta" _
        & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
        & " left join pv_bodegas X on D.cod_bodega_destino = X.cod_bodega" _
        & " where D.cod_producto = '" & vCodigo & "' and T.Tipo = 'E' and T.estado = 'P'" _
        & "order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Entrada")
        itmX.SubItems(1) = rs!Boleta
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!Procesa_user & ""
        itmX.SubItems(6) = Format(rs!Procesa_Fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close

 'Ultima Salida
 strSQL = "select Top 1 T.Tipo,T.Boleta,T.fecha,T.procesa_user,T.procesa_fecha" _
        & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
        & ",rtrim(D.cod_bodega_destino) + ' - ' + rtrim(X.descripcion) as BodegaD" _
        & ",D.precio,D.cantidad" _
        & " from pv_invTransac T inner join pv_invTraDet D" _
        & " on T.tipo = D.tipo and T.boleta = D.boleta" _
        & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
        & " left join pv_bodegas X on D.cod_bodega_destino = X.cod_bodega" _
        & " where D.cod_producto = '" & vCodigo & "' and T.Tipo = 'S' and T.estado = 'P'" _
        & "order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Salida")
        itmX.SubItems(1) = rs!Boleta
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!Procesa_user & ""
        itmX.SubItems(6) = Format(rs!Procesa_Fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close

 'Ultimo Traslado
 strSQL = "select Top 1 T.Tipo,T.Boleta,T.fecha,T.procesa_user,T.procesa_fecha" _
        & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
        & ",rtrim(D.cod_bodega_destino) + ' - ' + rtrim(X.descripcion) as BodegaD" _
        & ",D.precio,D.cantidad" _
        & " from pv_invTransac T inner join pv_invTraDet D" _
        & " on T.tipo = D.tipo and T.boleta = D.boleta" _
        & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
        & " left join pv_bodegas X on D.cod_bodega_destino = X.cod_bodega" _
        & " where D.cod_producto = '" & vCodigo & "' and T.Tipo = 'T' and T.estado = 'P'" _
        & "order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Traslado")
        itmX.SubItems(1) = rs!Boleta
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!Procesa_user & ""
        itmX.SubItems(6) = Format(rs!Procesa_Fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
  '--- > Movimientos Auxiliares Externos
  
  'Ultima Compra
  
  strSQL = "select Top 1 T.cod_compra,T.fecha,T.genera_user,T.genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from cpr_Compras T inner join cpr_Compras_detalle D" _
         & " on T.cod_factura = D.cod_factura and T.cod_proveedor = D.cod_proveedor" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "' and T.estado = 'P'" _
         & " order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Compra")
        itmX.SubItems(1) = rs!cod_compra
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!genera_user & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
  'Ultima DevCompra
  
  strSQL = "select Top 1 T.cod_compra_Dev as Codigo,T.fecha,T.genera_user,T.genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from cpr_compras_dev T inner join cpr_compra_devDet D" _
         & " on T.cod_compra_Dev = D.cod_compra_Dev" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "' and T.estado = 'P'" _
         & " order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Compra Dev.")
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!genera_user & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
    
  'Ultima AnuCompra
  strSQL = "select Top 1 T.cod_compra,T.Anula_Fec_Afecta as fecha,T.Anula_User as genera_user,T.Anula_Fecha as genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from cpr_Compras T inner join cpr_Compras_detalle D" _
         & " on T.cod_factura = D.cod_factura and T.cod_proveedor = D.cod_proveedor" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "' and T.estado = 'A'" _
         & " order by T.Anula_Fec_Afecta desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Compra Anu.")
        itmX.SubItems(1) = rs!cod_compra
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!genera_user & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
  
  
  'Ultima Venta
  strSQL = "select Top 1 T.cod_factura,T.fecha,T.usuario,T.fecha as genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from pv_facturacion T inner join pv_Factura_detalle D" _
         & " on T.cod_factura = D.cod_factura and T.tipo = D.tipo" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "'" _
         & " order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Venta")
        itmX.SubItems(1) = rs!cod_Factura
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!Usuario & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
  
  'Ultima DevVenta
  strSQL = "select Top 1 T.cod_devolucion as Codigo,T.fecha,T.user_genera,T.fecha as genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from pv_devoluciones T inner join pv_devolucion_Detalle D" _
         & " on T.cod_devolucion = D.cod_devolucion" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "'" _
         & " order by T.fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Venta Dev.")
        itmX.SubItems(1) = rs!COD_DEVOLUCION
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!user_genera & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close
  
  
  'Ultima Anu.Venta
  strSQL = "select Top 1 T.cod_factura,T.fecha,T.anu_cajaUser as usuario,T.Anu_fecha as genera_fecha" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaO" _
         & ",rtrim(D.cod_bodega) + ' - ' + rtrim(O.descripcion) as BodegaD" _
         & ",D.precio,D.cantidad" _
         & " from pv_facturacion T inner join pv_Factura_detalle D" _
         & " on T.cod_factura = D.cod_factura and T.tipo = D.tipo" _
         & " inner join pv_Bodegas O on D.cod_Bodega = O.cod_bodega" _
         & " where D.cod_producto = '" & vCodigo & "' and T.estado = 'A'" _
         & " order by T.Anu_fecha desc"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
    Set itmX = .ListItems.Add(, , "Venta Anula")
        itmX.SubItems(1) = rs!cod_Factura
        itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(rs!Precio, "Standard")
        itmX.SubItems(4) = rs!Cantidad
        itmX.SubItems(5) = rs!Usuario & ""
        itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
        itmX.SubItems(7) = rs!BodegaO & ""
        itmX.SubItems(8) = rs!BodegaD & ""
        itmX.ForeColor = vbBlue
   End If
   rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub lswSim_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

txtSimCodigo.Text = Item.Text
txtSimDesc.Text = Item.SubItems(1)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX  As ListViewItem, curSum As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

Select Case Item.Index
  Case 1 'Precios
    txtPrecio.Text = ""
    txtPrecioCod.Text = ""
    txtPrecioDesc.Text = ""
    lswPrecios.ListItems.Clear
    strSQL = "select P.*,isnull(X.monto,0) as Monto,isnull(X.porc_utilidad,0) as Utilidad" _
           & " from pv_tipos_precios P left join pv_producto_precios X on P.cod_precio = X.cod_precio" _
           & " and X.cod_producto = '" & vCodigo & "'" _
           & " order by P.defecto desc,X.monto"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      Set itmX = lswPrecios.ListItems.Add(, , rs!cod_precio)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!Monto, "Standard")
          itmX.SubItems(3) = Format(rs!Utilidad, "Standard")
      rs.MoveNext
    Loop
    rs.Close
    
    
  Case 2 'Descuentos y Bonificaciones
     strSQL = "select consec,desde,hasta,bonificacion from pv_producto_bonif" _
            & " where cod_producto = '" & vCodigo & "' order by desde"
     Call sbCargaGrid(vGridBon, 4, strSQL)
     
     strSQL = "select consec,desde,hasta,porcentaje from pv_producto_desc" _
            & " where cod_producto = '" & vCodigo & "' order by desde"
     Call sbCargaGrid(vGridDes, 4, strSQL)
  
  Case 3 'Proveedores
    lswProv.ListItems.Clear
    
    strSQL = "select P.COD_PROVEEDOR, P.DESCRIPCION, P.CEDJUR, max(F.fecha) as 'FECHA_FACTURA',X.cod_proveedor as 'CodX' " _
           & "   from CXP_PROVEEDORES P" _
           & "       left join pv_producto_prov X on P.COD_PROVEEDOR = X.COD_PROVEEDOR" _
           & "            and X.cod_producto = '" & vCodigo & "'" _
           & "       left join vCxP_Facturas F on P.COD_PROVEEDOR = F.COD_PROVEEDOR" _
           & " GROUP BY P.COD_PROVEEDOR, P.DESCRIPCION, P.CEDJUR,X.cod_proveedor" _
           & " order by X.cod_proveedor desc, P.Descripcion"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      Set itmX = lswProv.ListItems.Add(, , rs!cod_Proveedor)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = rs!CEDJUR
          itmX.SubItems(3) = rs!Fecha_Factura & ""
          
          itmX.Checked = IIf(IsNull(rs!CodX), False, True)
          itmX.ForeColor = IIf(IsNull(rs!CodX), vbBlack, vbBlue)
      rs.MoveNext
    Loop
    rs.Close
     
   Case 4 'Existencias
      Call sbInvCorte_Consulta
   
   Case 5 'Codigos de Barras
      Call sbgBarrasInicializa
      txtBarGenera.Text = txtCodigoBarras.Text
      Call sbBarras_Genera

   Case 6 'Movimientos
     Call sbMovimientos

   Case 7 'Similares
    lswSim.ListItems.Clear
    strSQL = "select cod_producto,descripcion,cabys" _
           & " from pv_productos" _
           & " where cod_producto <> '" _
           & vCodigo & "' and similar in(select isnull(similar,0) from pv_productos" _
           & " where cod_producto = '" & vCodigo & "')"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswSim.ListItems.Add(, , rs!Cod_Producto)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = rs!Cabys & ""
      rs.MoveNext
    Loop
    rs.Close
    
    txtSimCodigo.Text = ""
    txtSimDesc.Text = ""
    txtSimCabys.Text = ""
End Select


vError:
    vPaso = False
    Me.MousePointer = vbDefault

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
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
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_producto,descripcion from pv_productos"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       If txtCodigo <> "" Then
         Call txtCodigo_KeyDown(vbKeyReturn, 1)
       End If
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,C.descripcion as ProdClas,U.descripcion as UnidadDesc,M.Descripcion as MarcaDesc" _
       & ",isnull(Cs.Descripcion,'') as 'LineaSub', isnull(P.COD_LINEA_SUB,'') as 'LineaSubCod'" _
       & " from pv_productos P inner join pv_unidades U on P.cod_unidad = U.cod_unidad" _
       & " inner join pv_prod_clasifica C on P.cod_prodclas = C.cod_prodclas" _
       & " inner join pv_marcas M on P.cod_marca = M.cod_marca" _
       & "  left join PV_PROD_CLASIFICA_SUB Cs on P.cod_prodclas = Cs.cod_prodclas and P.COD_LINEA_SUB = Cs.COD_LINEA_SUB" _
       & " where P.cod_producto = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!Cod_Producto
  txtCodigo.Text = rs!Cod_Producto
  
  txtNombre.Text = rs!Descripcion & ""
    
  txtCodigoBarras.Text = rs!cod_barras & ""
  txtCabys.Text = rs!Cabys & ""
   
  chkLotes.Value = rs!lotes
   
  Select Case rs!tipo_producto
    Case "P"
        cboTipo.Text = "Producto"
    Case "S"
        cboTipo.Text = "Servicio"
    Case "A"
        cboTipo.Text = "Activo Fijo"
  End Select
  
  chkActivo.Value = IIf((rs!Estado = "A"), vbChecked, vbUnchecked)
  
  chkStock.Value = rs!I_Stock
  chkVentaEnLinea.Value = rs!I_VentaEnLinea
  
  txtLineaCod.Text = rs!cod_prodclas
  txtLineaDesc.Text = rs!ProdClas
  
  txtUnidadCod.Text = Trim(rs!Cod_Unidad)
  txtUnidadDesc.Text = rs!UnidadDesc
  
  txtMarcaCod.Text = Trim(rs!cod_Marca)
  txtMarcaDesc.Text = rs!MarcaDesc
  
  txtObservacion.Text = rs!Observacion & ""
   
  txtCodFabricante.Text = Trim(rs!cod_fabricante & "")
  txtModelo.Text = Trim(rs!Modelo)
  
  txtExistenciaMin.Text = Format(rs!inventario_minimo, "Standard")
  txtExistenciaMax.Text = Format(rs!inventario_maximo, "Standard")
  
  chkExistencia.Value = IIf((rs!inventario_calcula = "S"), vbChecked, vbUnchecked)
  txtComMonto.Text = Format(rs!comision_monto, "Standard")
  txtComUnidad.Text = Format(rs!comision_unidad, "###,###,##0")
  txtCostoRegular.Text = Format(rs!costo_regular, "Standard")
  txtPrecioRegular.Text = Format(rs!precio_regular, "Standard")
  txtUtilidadGeneral.Text = Format(rs!porc_utilidad, "Standard")
  txtImpConsumo.Text = Format(rs!impuesto_consumo, "Standard")
  txtImpVentas.Text = Format(rs!impuesto_ventas, "Standard")


  txtV_ItemsCantidad.Text = rs!VENTA_QTY_MAX
  txtV_ItemsFrecuencia.Text = rs!VENTA_FREQ_DIAS

  'Cargar despues aqui la ultima entrada y ultima salida
  Call sbCboAsignaDato(cboLineaSub, rs!LineaSub, True, rs!LineaSubCod)
  
  'Activa Tabs
    tcMain.Item(0).Selected = True
    For i = 1 To tcMain.ItemCount - 1
      tcMain.Item(i).Enabled = True
    Next i
  
  
  'Control
    txtRegFecha.Text = rs!REGISTRO_FECHA & ""
    txtRegUsuario.Text = rs!user_crea & ""
    

    txtModFecha.Text = rs!ULTIMA_MODIFICACION & ""
    txtModUsuario.Text = rs!user_modifica & ""

  
  'Calcula el Precio Final
  txtPrecioRegularIVA.Text = Format((CCur(txtPrecioRegular.Text) + (CCur(txtPrecioRegular.Text) * CCur(txtImpVentas.Text) / 100)), "Standard")
  
      
  
Else
  
  'Activa Tabs
    tcMain.Item(0).Selected = True
    For i = 1 To tcMain.ItemCount - 1
      tcMain.Item(i).Enabled = False
    Next i
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
'Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDecimal
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Producto no es válido ..."
If txtCodigo.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Código del Producto no es válido ..."

If Not IsNumeric(txtExistenciaMin.Text) Then vMensaje = vMensaje & vbCrLf & " - La Existencia Máxima no es válida!"
If Not IsNumeric(txtExistenciaMax.Text) Then vMensaje = vMensaje & vbCrLf & " - La Existencia Máxima no es válida!"

If Not IsNumeric(txtCostoRegular.Text) Then vMensaje = vMensaje & vbCrLf & " - Costo Regular no es válido!"
If Not IsNumeric(txtV_ItemsCantidad.Text) Then vMensaje = vMensaje & vbCrLf & " - Venta Permitida por Persona: Cantidad de Items no es válida!"
If Not IsNumeric(txtV_ItemsFrecuencia.Text) Then vMensaje = vMensaje & vbCrLf & " - Venta Permitida por Persona: Frecuencia no es válida!"

If Len(txtUnidadDesc.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - La Unidad de medida del Producto no es válida!"
If Len(txtMarcaDesc.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - La Marca del Producto no es válida!"

If Len(txtLineaDesc.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - La Línea/familia del Producto no es válida!"


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vEdita Then
  strSQL = "update pv_productos set descripcion = '" & Trim(txtNombre.Text) & "'" _
         & ", observacion = '" & txtObservacion.Text & "', estado = '" & IIf((chkActivo.Value = vbChecked), "A", "I") _
         & "', tipo_producto = '" & Mid(cboTipo.Text, 1, 1) & "', Cabys = '" & txtCabys.Text _
         & "', cod_barras = '" & txtCodigoBarras & "', cod_unidad = '" & txtUnidadCod & "', cod_prodclas = " & txtLineaCod _
         & ", cod_marca = '" & txtMarcaCod & "', modelo = '" & txtModelo & "'" _
         & ", cod_fabricante = '" & txtCodFabricante & "', inventario_minimo = " & CCur(txtExistenciaMin) _
         & ", inventario_maximo = " & CCur(txtExistenciaMax) _
         & ", inventario_calcula = '" & IIf((chkExistencia.Value = vbChecked), "S", "N") _
         & "', costo_regular = " & CCur(txtCostoRegular) & ", precio_regular = " & CCur(txtPrecioRegular) _
         & ", porc_utilidad = " & CCur(txtUtilidadGeneral) _
         & ", impuesto_consumo = " & CCur(txtImpConsumo) & ", impuesto_ventas = " & CCur(txtImpVentas) _
         & ", comision_monto = " & CCur(txtComMonto) & ", comision_unidad = " & CCur(txtComUnidad) _
         & ", fracciones = 1, dir_fotografia = '', lotes = " & chkLotes.Value _
         & ", ultima_modificacion = dbo.MyGetdate(),user_modifica = '" & glogon.Usuario _
         & "', COD_LINEA_SUB = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'" _
         & ", I_Stock = " & chkStock.Value & ", I_VentaEnLinea = " & chkVentaEnLinea.Value _
         & ", VENTA_FREQ_DIAS = " & txtV_ItemsFrecuencia.Text & ", VENTA_QTY_MAX = " & txtV_ItemsCantidad.Text _
         & " where cod_producto = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Producto : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into pv_productos(cod_producto, descripcion, observacion, estado" _
          & ", cod_barras, cabys, cod_unidad, cod_marca, cod_prodclas, tipo_producto, cod_fabricante" _
          & ", inventario_minimo, inventario_maximo, inventario_calcula, existencia, costo_regular, precio_regular" _
          & ", impuesto_consumo, impuesto_ventas, comision_monto, comision_unidad" _
          & ", fracciones, dir_fotografia, descuento_tipo, descuento_valor, user_crea" _
          & ", porc_utilidad, modelo, lotes, COD_LINEA_SUB, I_Stock, I_VentaEnLinea, VENTA_FREQ_DIAS, VENTA_QTY_MAX, Registro_Fecha)" _
          & " values('" & vCodigo & "','" _
          & UCase(txtNombre) & "','" & UCase(txtObservacion) & "','" _
          & IIf((chkActivo.Value = vbChecked), "A", "I") & "','" & txtCodigoBarras.Text & "','" & txtCabys.Text _
          & "','" & txtUnidadCod & "','" & txtMarcaCod & "'," & txtLineaCod _
          & ",'" & Mid(cboTipo.Text, 1, 1) & "','" & txtCodFabricante _
          & "'," & CCur(txtExistenciaMin.Text) & "," & CCur(txtExistenciaMax.Text) & ",'" & IIf((chkExistencia.Value = vbChecked), "S", "N") _
          & "',0," & CCur(txtCostoRegular.Text) & "," & CCur(txtPrecioRegular) & "," & CCur(txtImpConsumo.Text) _
          & "," & CCur(txtImpVentas.Text) & "," & CCur(txtComMonto.Text) & "," & CCur(txtComUnidad.Text) _
          & ",1,'','M',0,'" & glogon.Usuario & "'," & CCur(txtUtilidadGeneral) & ",'" _
          & txtModelo & "'," & chkLotes.Value & ",'" & cboLineaSub.ItemData(cboLineaSub.ListIndex) _
          & "', " & chkStock.Value & ", " & chkVentaEnLinea.Value & ", " & txtV_ItemsFrecuencia.Text & ", " & txtV_ItemsCantidad.Text & ", dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Producto : " & vCodigo)
 
End If

vEdita = True

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

For i = 1 To tcMain.ItemCount - 1
  tcMain.Item(i).Enabled = True
Next i

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete pv_productos where cod_producto = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Producto : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer, vSQL As String

vSQL = ""

Select Case ButtonMenu.Key
  Case "Boleta"
     i = MsgBox("Desea visualizar solo la Boleta del Producto Actual", vbYesNo)
     If i = vbYes Then vSQL = "{PV_PRODUCTOS.COD_PRODUCTO} = '" & txtCodigo & "'"
     
     Call sbInvReportes("Productos", "Boleta de Prouctos", "Boleta de Ingreso", vSQL)
    
  Case "ListadoGeneral"
     Call sbInvReportes("ProductosGen", "Listado General de Productos", "Productos/Acticulos y Servicios", vSQL)
  
  Case "InventarioGeneral"
     
     Call sbInvReportes("ProductosInv", "Inventario General de Productos", "Productos/Acticulos y Servicios", vSQL)

End Select
End Sub



Private Sub txtCabys_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cod_ByS"
  gBusquedas.Orden = "Cod_ByS"
  gBusquedas.Consulta = "select Cod_ByS,Descripcion from vINV_Cabys"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCabys.Text = gBusquedas.Resultado
End If

End Sub

Private Sub txtCodFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtModelo.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_producto"
  gBusquedas.Orden = "cod_producto"
  gBusquedas.Consulta = "select cod_producto,descripcion from pv_productos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


'Private Sub txtComMonto_GotFocus()
'On Error GoTo vError
'txtComMonto = CCur(txtComMonto)
'vError:
'End Sub
'
'Private Sub txtComMonto_LostFocus()
'On Error GoTo vError
'txtComMonto = Format(txtComMonto, "Standard")
'vError:
'End Sub
'
'Private Sub txtComUnidad_GotFocus()
'On Error GoTo vError
'txtComUnidad = CCur(txtComUnidad)
'vError:
'End Sub
'
'Private Sub txtComUnidad_LostFocus()
'On Error GoTo vError
'txtComUnidad = Format(txtComUnidad, "Standard")
'vError:
'End Sub
'
'
'Private Sub txtCostoRegular_GotFocus()
'On Error GoTo vError
'txtCostoRegular = CCur(txtCostoRegular)
'vError:
'End Sub

'Private Sub txtCostoRegular_LostFocus()
'On Error GoTo vError
'txtCostoRegular = Format(txtCostoRegular, "Standard")
'vError:
'End Sub


Private Sub txtExistenciaMax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComMonto.SetFocus
End Sub

Private Sub txtLineaCod_LostFocus()
txtLineaDesc = fxSIFCCodigos("D", txtLineaCod, "LineaProducto")
Call sbLineaSub_Carga
End Sub

Private Sub txtMarcaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMarcaDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_marca"
  gBusquedas.Orden = "cod_marca"
  gBusquedas.Consulta = "select cod_marca,descripcion from pv_marcas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtMarcaCod = gBusquedas.Resultado
  txtMarcaDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtMarcaCod_LostFocus()
txtMarcaDesc = fxSIFCCodigos("D", txtMarcaCod, "Marcas")
End Sub

Private Sub txtMarcaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_marca,descripcion from pv_marcas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtMarcaCod = gBusquedas.Resultado
  txtMarcaDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtImpConsumo.SetFocus
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs  As New ADODB.Recordset, i As Integer

On Error GoTo vError

'Guardar el Precio
If KeyCode = vbKeyReturn And txtPrecioCod.Text <> "" Then
  strSQL = "select isnull(count(*),0) as existe from pv_producto_precios" _
         & " where cod_producto = '" & vCodigo & "' and cod_precio = '" _
         & txtPrecioCod.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
    strSQL = "insert into pv_producto_precios(cod_producto,cod_precio,monto,porc_utilidad)" _
           & " values('" & vCodigo & "','" & txtPrecioCod.Text & "'," _
           & CCur(txtPrecio) & "," & CCur(txtUtilidad) & ")"
  Else
    strSQL = "update pv_producto_precios set monto = " & CCur(txtPrecio) _
         & ",porc_utilidad = " & CCur(txtUtilidad) _
         & " where cod_producto = '" & vCodigo & "' and cod_precio = '" _
         & txtPrecioCod.Text & "'"
  End If
  rs.Close
  
  Call ConectionExecute(strSQL)
  
  For i = 1 To lswPrecios.ListItems.Count
    If Trim(lswPrecios.ListItems.Item(i).Text) = Trim(txtPrecioCod.Text) Then
      lswPrecios.ListItems.Item(i).SubItems(2) = Format(txtPrecio, "Standard")
      lswPrecios.ListItems.Item(i).SubItems(3) = Format(txtUtilidad, "Standard")
      Exit For
    End If
  Next i
  
  txtPrecioCod.Text = ""
  txtPrecioDesc.Text = ""
  txtPrecio = ""
  txtUtilidad = ""
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtPrecio_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
txtUtilidad.Text = Format((CCur(txtPrecio.Text) - CCur(txtCostoRegular.Text)) / CCur(txtCostoRegular.Text) * 100, "Standard")
vError:
End Sub

'Private Sub txtPrecioRegular_GotFocus()
'On Error GoTo vError
'txtPrecioRegular = CCur(txtPrecioRegular)
'vError:
'End Sub

Private Sub txtPrecioRegular_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
txtUtilidadGeneral.Text = Format((CCur(txtPrecioRegular.Text) - CCur(txtCostoRegular.Text)) / CCur(txtCostoRegular.Text) * 100, "Standard")
vError:
End Sub

'Private Sub txtPrecioRegular_LostFocus()
'On Error GoTo vError
'txtPrecioRegular = Format(txtPrecioRegular, "Standard")
'vError:
'End Sub
'
'Private Sub txtExistenciaMin_GotFocus()
'On Error GoTo vError
'txtExistenciaMin = CInt(txtExistenciaMin)
'vError:
'End Sub
'
'Private Sub txtExistenciaMin_LostFocus()
'On Error GoTo vError
'txtExistenciaMin = Format(txtExistenciaMin, "###,###,###,##0")
'vError:
'End Sub

'Private Sub txtExistenciaMax_GotFocus()
'On Error GoTo vError
'txtExistenciaMax = CInt(txtExistenciaMax)
'vError:
'End Sub
'
'Private Sub txtExistenciaMax_LostFocus()
'On Error GoTo vError
'txtExistenciaMax = Format(txtExistenciaMax, "###,###,###,##0")
'vError:
'End Sub
'
'Private Sub txtImpConsumo_GotFocus()
'On Error GoTo vError
'txtImpConsumo = CCur(txtImpConsumo)
'vError:
'End Sub
'
'Private Sub txtImpConsumo_LostFocus()
'On Error GoTo vError
'txtImpConsumo = Format(txtImpConsumo, "Standard")
'vError:
'End Sub

'Private Sub txtImpVentas_GotFocus()
'On Error GoTo vError
'txtImpVentas = CCur(txtImpVentas)
'vError:
'End Sub
'
'Private Sub txtImpVentas_LostFocus()
'On Error GoTo vError
'txtImpVentas = Format(txtImpVentas, "Standard")
'vError:
'End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigoBarras.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_producto,descripcion from pv_productos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigoBarras_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_barras"
  gBusquedas.Orden = "cod_barras"
  gBusquedas.Consulta = "select cod_producto,descripcion,cod_barras from pv_productos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtComMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComUnidad.SetFocus
End Sub

Private Sub txtComUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodFabricante.SetFocus
End Sub

Private Sub txtCostoRegular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUtilidadGeneral.SetFocus
End Sub


Private Sub txtExistenciaMin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtExistenciaMax.SetFocus
End Sub

Private Sub txtImpConsumo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtImpVentas.SetFocus
End Sub

Private Sub txtImpVentas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCostoRegular.SetFocus
End Sub

Private Sub sbLineaSub_Carga()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_LINEA_SUB as 'IdX',  DESCRIPCION as 'ItmX'" _
    & " From PV_PROD_CLASIFICA_SUB where COD_PRODCLAS = " & txtLineaCod.Text
Call sbCbo_Llena_New(cboLineaSub, strSQL, False, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub txtLineaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtLineaDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_prodclas"
  gBusquedas.Orden = "cod_prodclas"
  gBusquedas.Consulta = "select cod_prodclas,descripcion from pv_prod_clasifica"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtLineaCod = gBusquedas.Resultado
  txtLineaDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtLineaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboLineaSub.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_prodclas,descripcion from pv_prod_clasifica"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtLineaCod = gBusquedas.Resultado
  txtLineaDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtExistenciaMin.SetFocus
End Sub

Private Sub txtPrecioRegular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub

Private Sub txtSimCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Resultado3 = ""
  frmBusquedaArticulos.Show vbModal
  
  txtSimCodigo.Text = gBusquedas.Resultado
  txtSimDesc.Text = gBusquedas.Resultado2
  txtSimCabys.Text = gBusquedas.Resultado3
End If


End Sub

Private Sub txtSimDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Resultado3 = ""
  frmBusquedaArticulos.Show vbModal
  
  txtSimCodigo.Text = gBusquedas.Resultado
  txtSimDesc.Text = gBusquedas.Resultado2
  txtSimCabys.Text = gBusquedas.Resultado3
End If
End Sub

Private Sub txtUnidadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_unidad"
  gBusquedas.Orden = "cod_unidad"
  gBusquedas.Consulta = "select cod_unidad,descripcion from pv_unidades"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidadCod.Text = gBusquedas.Resultado
  txtUnidadDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtUnidadCod_LostFocus()
txtUnidadDesc = fxSIFCCodigos("D", txtUnidadCod, "Unidades")
End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMarcaCod.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_unidad,descripcion from pv_unidades"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidadCod.Text = gBusquedas.Resultado
  txtUnidadDesc.Text = gBusquedas.Resultado2
End If
End Sub


Private Sub txtUtilidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call txtPrecio_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtUtilidad_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
txtPrecio.Text = Format((CCur(txtCostoRegular.Text) + (CCur(txtCostoRegular.Text) * CCur(txtUtilidad.Text) / 100)), "Standard")
vError:
End Sub


Private Sub txtUtilidadGeneral_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPrecioRegular.SetFocus
End Sub

Private Sub txtUtilidadGeneral_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
txtPrecioRegular.Text = Format((CCur(txtCostoRegular.Text) + (CCur(txtCostoRegular.Text) * CCur(txtUtilidadGeneral.Text) / 100)), "Standard")
txtPrecioRegularIVA.Text = Format((CCur(txtPrecioRegular.Text) + (CCur(txtPrecioRegular.Text) * CCur(txtImpVentas.Text) / 100)), "Standard")
vError:
End Sub

Private Sub vGridBon_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

If vGridBon.ActiveCol = vGridBon.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  vGridBon.Row = vGridBon.ActiveRow
  vGridBon.Col = 1
  If vGridBon.Text = "" Then
     strSQL = "select (isnull(max(consec),0) + 1) as Consecutivo from pv_producto_bonif" _
            & " where cod_producto = '" & vCodigo & "'"
     Call OpenRecordSet(rs, strSQL)
         i = rs!Consecutivo
     rs.Close
     
     vGridBon.Col = 2
     strSQL = "insert pv_producto_bonif(consec,cod_producto,desde,hasta,bonificacion)" _
             & " values(" & i & ",'" & vCodigo & "'," & CCur(vGridBon.Text) & ","
     vGridBon.Col = 3
     strSQL = strSQL & CCur(vGridBon.Text) & ","
     vGridBon.Col = 4
     strSQL = strSQL & CCur(vGridBon.Text) & ")"
     
     vGridBon.Col = 1
     vGridBon.Text = CStr(i)
     
  Else
     vGridBon.Col = 2
     strSQL = "update pv_producto_bonif set desde = " & CCur(vGridBon.Text) & ",hasta = "
     vGridBon.Col = 3
     strSQL = strSQL & CCur(vGridBon.Text) & ",bonificacion = "
     vGridBon.Col = 4
     strSQL = strSQL & CCur(vGridBon.Text) & " where consec = "
     vGridBon.Col = 1
     strSQL = strSQL & vGridBon.Text & " and cod_producto = '" & vCodigo & "'"
  End If
  
  Call ConectionExecute(strSQL)
  
  If vGridBon.MaxRows <= vGridBon.ActiveRow Then
    vGridBon.MaxRows = vGridBon.MaxRows + 1
    vGridBon.Row = vGridBon.MaxRows
  End If
End If

If KeyCode = vbKeyDelete Then
  vGridBon.Row = vGridBon.ActiveRow
  vGridBon.Col = 1
  strSQL = "delete pv_producto_bonif where consec = " & vGridBon.Text _
         & " and cod_producto = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  strSQL = "select consec,desde,hasta,bonificacion from pv_producto_bonif" _
         & " where cod_producto = '" & vCodigo & "' order by desde"
  Call sbCargaGrid(vGridBon, 4, strSQL)
End If


End Sub


Private Sub vGridDes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

If vGridDes.ActiveCol = vGridDes.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  vGridDes.Row = vGridDes.ActiveRow
  vGridDes.Col = 1
  If vGridDes.Text = "" Then
     strSQL = "select (isnull(max(consec),0) + 1) as Consecutivo from pv_producto_desc" _
            & " where cod_producto = '" & vCodigo & "'"
     Call OpenRecordSet(rs, strSQL)
         i = rs!Consecutivo
     rs.Close
     
     vGridDes.Col = 2
     strSQL = "insert pv_producto_desc(consec,cod_producto,desde,hasta,porcentaje)" _
             & " values(" & i & ",'" & vCodigo & "'," & CCur(vGridDes.Text) & ","
     vGridDes.Col = 3
     strSQL = strSQL & CCur(vGridDes.Text) & ","
     vGridDes.Col = 4
     strSQL = strSQL & CCur(vGridDes.Text) & ")"
     
     vGridDes.Col = 1
     vGridDes.Text = CStr(i)
     
  Else
     vGridDes.Col = 2
     strSQL = "update pv_producto_desc set desde = " & CCur(vGridDes.Text) & ",hasta = "
     vGridDes.Col = 3
     strSQL = strSQL & CCur(vGridDes.Text) & ",porcentaje = "
     vGridDes.Col = 4
     strSQL = strSQL & CCur(vGridDes.Text) & " where consec = "
     vGridDes.Col = 1
     strSQL = strSQL & vGridDes.Text & " and cod_producto = '" & vCodigo & "'"
  End If
  
  Call ConectionExecute(strSQL)
  
  If vGridDes.MaxRows <= vGridDes.ActiveRow Then
    vGridDes.MaxRows = vGridDes.MaxRows + 1
    vGridDes.Row = vGridDes.MaxRows
  End If
End If

If KeyCode = vbKeyDelete Then
  vGridDes.Row = vGridDes.ActiveRow
  vGridDes.Col = 1
  strSQL = "delete pv_producto_desc where consec = " & vGridDes.Text _
         & " and cod_producto = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  strSQL = "select consec,desde,hasta,porcentaje from pv_producto_desc" _
         & " where cod_producto = '" & vCodigo & "' order by desde"
  Call sbCargaGrid(vGridDes, 4, strSQL)
End If


End Sub

