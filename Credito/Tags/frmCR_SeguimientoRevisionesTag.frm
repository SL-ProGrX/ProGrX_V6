VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_SeguimientoRevisionesTag 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Créditos"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   12135
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7815
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   12135
      _Version        =   1441793
      _ExtentX        =   21405
      _ExtentY        =   13785
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
      Item(0).Caption =   "Operaciones"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGridSolicitudes"
      Item(0).Control(1)=   "gbFiltros"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "SSTabDetalle"
      Item(2).Caption =   "Seguimiento"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridSeguimiento"
      Item(3).Caption =   "Revisión"
      Item(3).ControlCount=   7
      Item(3).Control(0)=   "txtObservacion"
      Item(3).Control(1)=   "lswErrores"
      Item(3).Control(2)=   "cboEtiquetas"
      Item(3).Control(3)=   "Label8"
      Item(3).Control(4)=   "Label2"
      Item(3).Control(5)=   "Label27"
      Item(3).Control(6)=   "btnAplicar"
      Begin XtremeSuiteControls.ListView lswErrores 
         Height          =   3495
         Left            =   -68320
         TabIndex        =   133
         Top             =   1080
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1441793
         _ExtentX        =   16536
         _ExtentY        =   6165
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbFiltros 
         Height          =   1935
         Left            =   0
         TabIndex        =   10
         Top             =   5880
         Width           =   12135
         _Version        =   1441793
         _ExtentX        =   21405
         _ExtentY        =   3413
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswBancos 
            Height          =   1455
            Left            =   5880
            TabIndex        =   132
            Top             =   480
            Width           =   6255
            _Version        =   1441793
            _ExtentX        =   11033
            _ExtentY        =   2566
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
         End
         Begin XtremeSuiteControls.CheckBox chkSoloEspera 
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   960
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Solo Créditos en Espera"
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
         End
         Begin XtremeSuiteControls.ComboBox cboEtiquetasFiltro 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   5655
            _Version        =   1441793
            _ExtentX        =   9975
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
         Begin XtremeSuiteControls.PushButton btnFiltros 
            Height          =   495
            Left            =   360
            TabIndex        =   15
            Top             =   1440
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Filtrar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":0000
         End
         Begin XtremeSuiteControls.Label Label58 
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   12
            Top             =   240
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Bancos:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label58 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Etiquetas:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin FPSpreadADO.fpSpread vGridSolicitudes 
         Height          =   5535
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   12135
         _Version        =   524288
         _ExtentX        =   21405
         _ExtentY        =   9763
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
         MaxCols         =   491
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":0708
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin TabDlg.SSTab SSTabDetalle 
         Height          =   7215
         Left            =   -69520
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12726
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   7
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Crédito y Patrimonio"
         TabPicture(0)   =   "frmCR_SeguimientoRevisionesTag.frx":12B8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label15"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label13(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblDiferenciaCuota"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblCuota"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblTotalCuotas"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblCuotaDesembolsos"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblCuotaRefundicion"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label40"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblMontoDesembolsos"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblMontoRefundicion"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblMonto_Girado"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lblMontoApr"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label17"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label16"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Line2"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label10"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label9"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label6"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label26"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label5"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "lblGarantia"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label1"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label7"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Label21"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Label20"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Label12"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Label19"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "lblAhorrosTotal"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "lblAhorrosExtra"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "lblCapitalizacion"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "lblAportePatronal"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "lblAporteObrero"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "lblAhorrosFCorte"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Label23"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Label18"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "lblAhorrosFecha"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "Label38"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "lblDisponibleBruto"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "Label39"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "lblSaldoPrestamos"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "Label42"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "Label44"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "lblDisponible"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).ControlCount=   45
         TabCaption(1)   =   "Información Personal"
         TabPicture(1)   =   "frmCR_SeguimientoRevisionesTag.frx":12D4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label67"
         Tab(1).Control(1)=   "lblDireccion"
         Tab(1).Control(2)=   "lblDistrito"
         Tab(1).Control(3)=   "Label36"
         Tab(1).Control(4)=   "lblCanton"
         Tab(1).Control(5)=   "Label35"
         Tab(1).Control(6)=   "lblProvincia"
         Tab(1).Control(7)=   "Label78"
         Tab(1).Control(8)=   "Label77"
         Tab(1).Control(9)=   "Label76"
         Tab(1).Control(10)=   "lblFiadorLiqSFianza"
         Tab(1).Control(11)=   "lblFiadorLiqCFianza"
         Tab(1).Control(12)=   "lblFiadorSalLiquido"
         Tab(1).Control(13)=   "Label72"
         Tab(1).Control(14)=   "lblFiadorInstitucion"
         Tab(1).Control(15)=   "lblFiadorIngreso"
         Tab(1).Control(16)=   "lblFiadorMembresia"
         Tab(1).Control(17)=   "Label61"
         Tab(1).Control(18)=   "Label55"
         Tab(1).Control(19)=   "lblFiadorLiqCFianzaPorc"
         Tab(1).Control(20)=   "Label22"
         Tab(1).Control(21)=   "Label30"
         Tab(1).Control(22)=   "lblFiadorLiqSFianzaPorc"
         Tab(1).Control(23)=   "Label32"
         Tab(1).Control(24)=   "Label57"
         Tab(1).Control(25)=   "Label11"
         Tab(1).Control(26)=   "lblPreanalisis"
         Tab(1).Control(27)=   "vGrid"
         Tab(1).Control(28)=   "lswFiadores"
         Tab(1).ControlCount=   29
         TabCaption(2)   =   "Deudas y Fianzas"
         TabPicture(2)   =   "frmCR_SeguimientoRevisionesTag.frx":12F0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label54"
         Tab(2).Control(1)=   "lblSuma"
         Tab(2).Control(2)=   "Label37"
         Tab(2).Control(3)=   "Label33"
         Tab(2).Control(4)=   "Label34"
         Tab(2).Control(5)=   "lblDeudasTotal"
         Tab(2).Control(6)=   "lblDeducciones"
         Tab(2).Control(7)=   "lblDeudasCuota"
         Tab(2).Control(8)=   "Label25"
         Tab(2).Control(9)=   "Label53"
         Tab(2).Control(10)=   "Label41"
         Tab(2).Control(11)=   "lblFianzasMonto"
         Tab(2).Control(12)=   "Label29"
         Tab(2).Control(13)=   "lblFianzasSaldo"
         Tab(2).Control(14)=   "Label31"
         Tab(2).Control(15)=   "lblFianzasCuota"
         Tab(2).Control(16)=   "Label51"
         Tab(2).Control(17)=   "Line1"
         Tab(2).Control(18)=   "vGridFianzas"
         Tab(2).Control(19)=   "vGridDeudas"
         Tab(2).ControlCount=   20
         TabCaption(3)   =   "Refundiciones y Desembolsos"
         TabPicture(3)   =   "frmCR_SeguimientoRevisionesTag.frx":130C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label24"
         Tab(3).Control(1)=   "Label52"
         Tab(3).Control(2)=   "lblRefundeCuota"
         Tab(3).Control(3)=   "Label43"
         Tab(3).Control(4)=   "Label50"
         Tab(3).Control(5)=   "Label45"
         Tab(3).Control(6)=   "lblRefundeMonto"
         Tab(3).Control(7)=   "Line3"
         Tab(3).Control(8)=   "Label48"
         Tab(3).Control(9)=   "Label49"
         Tab(3).Control(10)=   "lblDesembolsoCuota"
         Tab(3).Control(11)=   "lblDesembolsoMonto"
         Tab(3).Control(12)=   "Label47"
         Tab(3).Control(13)=   "Label46"
         Tab(3).Control(14)=   "vGridDesembolsos"
         Tab(3).Control(15)=   "vGridRefundiciones"
         Tab(3).ControlCount=   16
         Begin XtremeSuiteControls.ListView lswFiadores 
            Height          =   1455
            Left            =   -74640
            TabIndex        =   131
            Top             =   240
            Width           =   10335
            _Version        =   1441793
            _ExtentX        =   18230
            _ExtentY        =   2566
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
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   2055
            Left            =   -72000
            TabIndex        =   17
            Top             =   4320
            Width           =   7215
            _Version        =   524288
            _ExtentX        =   12726
            _ExtentY        =   3625
            _StockProps     =   64
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ScrollBars      =   0
            SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":1328
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridDeudas 
            Height          =   2700
            Left            =   -72600
            TabIndex        =   18
            Top             =   420
            Width           =   8415
            _Version        =   524288
            _ExtentX        =   14843
            _ExtentY        =   4762
            _StockProps     =   64
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   501
            SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":188D
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridFianzas 
            Height          =   2700
            Left            =   -72600
            TabIndex        =   19
            Top             =   3660
            Width           =   8415
            _Version        =   524288
            _ExtentX        =   14843
            _ExtentY        =   4762
            _StockProps     =   64
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   499
            SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":23DB
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridRefundiciones 
            Height          =   2700
            Left            =   -72600
            TabIndex        =   20
            Top             =   300
            Width           =   8415
            _Version        =   524288
            _ExtentX        =   14843
            _ExtentY        =   4762
            _StockProps     =   64
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   497
            SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":2ACA
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridDesembolsos 
            Height          =   2700
            Left            =   -72600
            TabIndex        =   21
            Top             =   3660
            Width           =   8415
            _Version        =   524288
            _ExtentX        =   14843
            _ExtentY        =   4762
            _StockProps     =   64
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   492
            SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":31AB
            AppearanceStyle =   1
         End
         Begin VB.Label Label46 
            Caption         =   "Desembolsos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   123
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label Label47 
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   122
            Top             =   4020
            Width           =   855
         End
         Begin VB.Label lblDesembolsoMonto 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -74880
            TabIndex        =   121
            Top             =   4260
            Width           =   1995
         End
         Begin VB.Label lblDesembolsoCuota 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -74880
            TabIndex        =   120
            Top             =   4980
            Width           =   1995
         End
         Begin VB.Label Label49 
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   119
            Top             =   4740
            Width           =   855
         End
         Begin VB.Label Label48 
            Caption         =   "Fuente: Preanálisis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -65280
            TabIndex        =   118
            Top             =   6480
            Width           =   1215
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   -74880
            X2              =   -64200
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblRefundeMonto 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -74880
            TabIndex        =   117
            Top             =   900
            Width           =   1995
         End
         Begin VB.Label Label45 
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   116
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label50 
            Caption         =   "Refundiciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   115
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label43 
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   114
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label lblRefundeCuota 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -74880
            TabIndex        =   113
            Top             =   1620
            Width           =   1995
         End
         Begin VB.Label Label52 
            Caption         =   "Fuente: Solicitud Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -65400
            TabIndex        =   112
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   -74880
            X2              =   -64200
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label51 
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   111
            Top             =   5460
            Width           =   855
         End
         Begin VB.Label lblFianzasCuota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   110
            Top             =   5700
            Width           =   1995
         End
         Begin VB.Label Label31 
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   109
            Top             =   4740
            Width           =   855
         End
         Begin VB.Label lblFianzasSaldo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   108
            Top             =   4980
            Width           =   1995
         End
         Begin VB.Label Label29 
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   107
            Top             =   4020
            Width           =   855
         End
         Begin VB.Label lblFianzasMonto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   106
            Top             =   4260
            Width           =   1995
         End
         Begin VB.Label Label41 
            Caption         =   "Fianzas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   105
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label53 
            Caption         =   "Fuente: Fianzas Activas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -68160
            TabIndex        =   104
            Top             =   6480
            Width           =   4095
         End
         Begin VB.Label Label25 
            Caption         =   "Total Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   103
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label lblDeudasCuota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   102
            Top             =   1620
            Width           =   1995
         End
         Begin VB.Label lblDeducciones 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   101
            Top             =   2340
            Width           =   1995
         End
         Begin VB.Label lblDeudasTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74880
            TabIndex        =   100
            Top             =   900
            Width           =   1995
         End
         Begin VB.Label Label34 
            Caption         =   "Deducciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   99
            Top             =   2100
            Width           =   1335
         End
         Begin VB.Label Label33 
            Caption         =   "Total Saldo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   98
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Deudas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblPreanalisis 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74640
            TabIndex        =   96
            ToolTipText     =   "Membresia"
            Top             =   4560
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "No. Expediente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74640
            TabIndex        =   95
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label Label57 
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74640
            TabIndex        =   94
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label32 
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74640
            TabIndex        =   93
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblFiadorLiqSFianzaPorc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -65400
            TabIndex        =   92
            ToolTipText     =   "Membresia"
            Top             =   2640
            Width           =   1005
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -65400
            TabIndex        =   91
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -69360
            TabIndex        =   90
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblFiadorLiqCFianzaPorc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -69360
            TabIndex        =   89
            ToolTipText     =   "Membresia"
            Top             =   2640
            Width           =   1005
         End
         Begin VB.Label lblDisponible 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7440
            TabIndex        =   88
            Top             =   6000
            Width           =   2160
         End
         Begin VB.Label Label44 
            Caption         =   "Disponible"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   87
            Top             =   6000
            Width           =   1935
         End
         Begin VB.Label Label42 
            Caption         =   "(-) Saldo préstamos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   86
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Label lblSaldoPrestamos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7440
            TabIndex        =   85
            Top             =   5400
            Width           =   2160
         End
         Begin VB.Label Label39 
            Caption         =   "Disponible bruto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   84
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label lblDisponibleBruto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7440
            TabIndex        =   83
            ToolTipText     =   "Ahorros a la fecha * Porcentaje para préstamos sobre ahorro"
            Top             =   5040
            Width           =   2160
         End
         Begin VB.Label Label38 
            Caption         =   "Ahorros a la fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   82
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label lblAhorrosFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7440
            TabIndex        =   81
            ToolTipText     =   "Aporte Obrero + Capitalización"
            Top             =   4680
            Width           =   2160
         End
         Begin VB.Label Label18 
            Caption         =   "Obrero"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   80
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Fecha Corte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   79
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label lblAhorrosFCorte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7440
            TabIndex        =   78
            Top             =   4320
            Width           =   2160
         End
         Begin VB.Label lblAporteObrero 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   77
            Top             =   4320
            Width           =   2160
         End
         Begin VB.Label lblAportePatronal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   76
            Top             =   4680
            Width           =   2160
         End
         Begin VB.Label lblCapitalizacion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   75
            Top             =   5040
            Width           =   2160
         End
         Begin VB.Label lblAhorrosExtra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   74
            Top             =   5400
            Width           =   2160
         End
         Begin VB.Label lblAhorrosTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   73
            Top             =   6000
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "Patronal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   72
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Capitalización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   71
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Ahorros Extra"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   70
            Top             =   5400
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   69
            Top             =   6000
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Patrimonio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   68
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label lblSuma 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -66120
            TabIndex        =   67
            ToolTipText     =   "Cédula"
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Cálculos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   66
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label55 
            Caption         =   "Fuente: Estado al Momento del Estudio de Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -68160
            TabIndex        =   65
            Top             =   6480
            Width           =   3135
         End
         Begin VB.Label lblGarantia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   64
            Top             =   3000
            Width           =   2100
         End
         Begin VB.Label Label5 
            Caption         =   "Garantía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   63
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   62
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   61
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   60
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Monto Solicitado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   59
            Top             =   960
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   360
            X2              =   10440
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label Label16 
            Caption         =   "Nueva Cuota"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   58
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label17 
            Caption         =   "Aumenta/Disminuye"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   57
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label Label24 
            Caption         =   "Deudas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74880
            TabIndex        =   56
            Top             =   -180
            Width           =   1935
         End
         Begin VB.Label lblMontoApr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   55
            Top             =   960
            Width           =   2100
         End
         Begin VB.Label lblMonto_Girado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   54
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblMontoRefundicion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   53
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label lblMontoDesembolsos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   2880
            TabIndex        =   52
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5400
            TabIndex        =   51
            Top             =   960
            Width           =   2100
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5400
            TabIndex        =   50
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblCuotaRefundicion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5400
            TabIndex        =   49
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label lblCuotaDesembolsos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5400
            TabIndex        =   48
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label lblTotalCuotas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5400
            TabIndex        =   47
            Top             =   2520
            Width           =   2100
         End
         Begin VB.Label lblCuota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7800
            TabIndex        =   46
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblDiferenciaCuota 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   7800
            TabIndex        =   45
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label Label28 
            Caption         =   "Fianzas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   -480
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "Monto Girado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   43
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Refundición"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   42
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Desembolsos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   41
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label61 
            Caption         =   "Membresía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74640
            TabIndex        =   40
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblFiadorMembresia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74640
            TabIndex        =   39
            ToolTipText     =   "Membresia"
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label lblFiadorIngreso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -66480
            TabIndex        =   38
            ToolTipText     =   "Membresia"
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label lblFiadorInstitucion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -71040
            TabIndex        =   37
            ToolTipText     =   "Membresia"
            Top             =   2040
            Width           =   4305
         End
         Begin VB.Label Label72 
            Caption         =   "Institución"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -71040
            TabIndex        =   36
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblFiadorSalLiquido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74640
            TabIndex        =   35
            ToolTipText     =   "Membresia"
            Top             =   2640
            Width           =   2400
         End
         Begin VB.Label lblFiadorLiqCFianza 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -72000
            TabIndex        =   34
            ToolTipText     =   "Membresia"
            Top             =   2640
            Width           =   2400
         End
         Begin VB.Label lblFiadorLiqSFianza 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -68040
            TabIndex        =   33
            ToolTipText     =   "Membresia"
            Top             =   2640
            Width           =   2400
         End
         Begin VB.Label Label76 
            Caption         =   "Salario Liquido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74640
            TabIndex        =   32
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label77 
            Caption         =   "Liq C/Fianzas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -72000
            TabIndex        =   31
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label78 
            Caption         =   "Liq S/Fianzas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -68040
            TabIndex        =   30
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label lblProvincia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74640
            TabIndex        =   29
            ToolTipText     =   "Membresia"
            Top             =   3240
            Width           =   3195
         End
         Begin VB.Label Label35 
            Caption         =   "Cantón"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -71160
            TabIndex        =   28
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblCanton 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -71160
            TabIndex        =   27
            ToolTipText     =   "Membresia"
            Top             =   3240
            Width           =   3195
         End
         Begin VB.Label Label36 
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -67560
            TabIndex        =   26
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblDistrito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -67560
            TabIndex        =   25
            ToolTipText     =   "Membresia"
            Top             =   3240
            Width           =   3195
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   -74640
            TabIndex        =   24
            ToolTipText     =   "Membresia"
            Top             =   3840
            Width           =   10290
         End
         Begin VB.Label Label67 
            Caption         =   "Ingreso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -66480
            TabIndex        =   23
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label54 
            Caption         =   "Fuente: Créditos Activos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   -68280
            TabIndex        =   22
            Top             =   3120
            Width           =   4095
         End
      End
      Begin FPSpreadADO.fpSpread vGridSeguimiento 
         Height          =   7335
         Left            =   -70000
         TabIndex        =   124
         Top             =   360
         Visible         =   0   'False
         Width           =   12135
         _Version        =   524288
         _ExtentX        =   21405
         _ExtentY        =   12938
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
         MaxCols         =   487
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoRevisionesTag.frx":3728
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEtiquetas 
         Height          =   330
         Left            =   -68320
         TabIndex        =   125
         Top             =   600
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
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
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   2055
         Left            =   -68320
         TabIndex        =   129
         Top             =   4680
         Visible         =   0   'False
         Width           =   9375
         _Version        =   1441793
         _ExtentX        =   16536
         _ExtentY        =   3625
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
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   -60640
         TabIndex        =   130
         Top             =   6960
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aplicar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoRevisionesTag.frx":3D09
      End
      Begin VB.Label Label27 
         Caption         =   "Omisiones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69760
         TabIndex        =   128
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Etiqueta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69760
         TabIndex        =   127
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69760
         TabIndex        =   126
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.GroupBox fraOperacion 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   12015
      _Version        =   1441793
      _ExtentX        =   21193
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   360
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label56 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Operación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.PushButton btnProceso 
      Height          =   495
      Index           =   0
      Left            =   8640
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Estudio de Crédito"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_SeguimientoRevisionesTag.frx":4430
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":4A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":B2B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":11B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":11C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":11D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":185AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":1EE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":25670
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":2BED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":32734
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   6840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":38F96
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":390B4
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":391DA
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":39304
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":39416
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":3952D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":3962E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":39765
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRevisionesTag.frx":3987A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnProceso 
      Height          =   495
      Index           =   1
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Trámite de Crédito"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_SeguimientoRevisionesTag.frx":3999E
   End
   Begin XtremeSuiteControls.Label lblNombreUsuario 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _Version        =   1441793
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   600
      Picture         =   "frmCR_SeguimientoRevisionesTag.frx":39FC2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmCR_SeguimientoRevisionesTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Dim vPaso As Boolean

Private mOperacion As String
Private mPreanalisis As String
Private mNuevoEstado As String
Private mFondoSolidario As Double
Private rslocal As New ADODB.Recordset
Private mTagRevision As String

Private Sub cboFiltroEstado_Click()
    Call sbCargarListaSolicitudes
End Sub

Private Sub btnAplicar_Click()

On Error GoTo vError
    
    Me.MousePointer = vbHourglass
    
    If Trim(cboEtiquetas.Text) = Empty Then
        MsgBox "Debe seleccionar la etiqueta que desea plicar"
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    txtObservacion.Text = fxSysCleanTxtInject(txtObservacion.Text)
    
    strSQL = "select isnull(Nota_Largo,0) as 'Nota_Largo' from Crd_Tags where Tag_Codigo = '" & cboEtiquetas.ItemData(cboEtiquetas.ListIndex) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs!Nota_Largo > Len(txtObservacion.Text) Then
        MsgBox "Este tipo de Etiqueta Requiere que la Nota sea de al menos " & rs!Nota_Largo & " caracteres!", vbExclamation
        Exit Sub
    End If


    If MsgBox("Está seguro que sea aplicar la etiqueta en las operaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Call sbIncluirEtiquetas
   
    
    Call sbAplicarErrores
    
    tcMain.Item(0).Selected = True
    
    Call sbLimpiarDatosCreditos(True)
    Call sbCargarListaSolicitudes
    
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFiltros_Click()
Call sbCargarListaSolicitudes
End Sub

Private Sub btnProceso_Click(Index As Integer)
    Dim frm As Form
    On Error GoTo vError

    If Len(Trim(mOperacion)) = 0 Then
        MsgBox "Debe seleccionar una solicitud", vbExclamation
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    Select Case Index
        Case 0 '"PREANALISIS"
        
            Dim strSQL As String, rs As New ADODB.Recordset
            Dim x As clsEstudioCrd
            
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = mOperacion
            x.xkey = glogon.ConectRPT
            
            strSQL = "select cod_preAnalisis from CRD_PREA_PREANALISIS" _
                       & " Where id_solicitud = " & mOperacion
            
            Call OpenRecordSet(rs, strSQL)
            If rs.EOF And rs.BOF Then
                Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                            , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
            
            Else
                x.vSolicitudPreanalisis = rs!cod_PreAnalisis
                Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                            , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
            End If
            rs.Close

            Set x = Nothing
            
        Case 1 '"SEGUIMIENTO"

            Call sbFormsCall("frmCR_SeguimientoTramites")
            For Each frm In Forms
                If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbConsultaExterna(Val(mOperacion))
                    Exit For
                End If
            Next frm

    End Select
    
     Me.MousePointer = vbDefault
    
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboEtiquetas_Click()
    Call sbCargarObservacion
End Sub



Private Sub Form_Activate()

vModulo = 8
End Sub

Private Sub Form_Load()


vModulo = 8

    vGridSolicitudes.AppearanceStyle = fxGridStyle
    vGridSeguimiento.AppearanceStyle = fxGridStyle
    
    vGridSolicitudes.MaxRows = 0
    vGridSolicitudes.MaxCols = 13
    vGridSeguimiento.MaxRows = 0
    vGridSeguimiento.MaxCols = 4
    
    
    tcMain.Item(0).Selected = True
    
    SSTabDetalle.Tab = 0
    
    lblNombreUsuario.Caption = fxNombreUsuario
    
    
    Call sbParametrosTagRevision
    mFondoSolidario = fxFondoSolPreanalisis
    Call sbCargarCombosEtiquetas
    Call sbCargarListaBancos
    Call sbCargarListaSolicitudes

End Sub

'Private Sub sbLlenarListaBancos()
'
'On Error GoTo Error
'
'    Dim vItem As MSComctlLib.ListItem
'    Dim vLvw As MSComctlLib.ListView
'    Dim vKey As String
'    Dim rs As New ADODB.Recordset
'
'
'    Me.lswBancos.ColumnHeaders.Clear
'    Me.lswBancos.ListItems.Clear
'
'    Set vLvw = Me.lswBancos
'    vLvw.ColumnHeaders.Add , , "Banco", 2100
'
'    strSQL = "select GARANTIA, DESCRIPCION from CRD_GARANTIA_TIPOS " & _
'             " order by DESCRIPCION "
'    Call OpenRecordSet(rs, strSQL)
'
'    While Not rs.EOF
'
'        vKey = Trim(rs.Fields("GARANTIA")) & "(GA)"
'
'        Set vItem = lswGarantias.ListItems.Add(, vKey, Trim(rs.Fields!Descripcion))
'
'        rs.MoveNext
'    Wend
'
'
'    Exit Sub
'Error:
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'
'End Sub

Private Sub sbCargarListaBancos()

On Error GoTo vError

With lswBancos.ColumnHeaders
    .Clear
    .Add , , "Cuenta Id", 1800
    .Add , , "Descripción", lswBancos.Width - 2000
End With

    With lswBancos
     .ListItems.Clear
     .Checkboxes = True
     
     strSQL = "select ID_BANCO,DESCRIPCION from BANCOS WHERE ESTADO = 'A'"
            
     Call OpenRecordSet(rs, strSQL, 0)
     
     Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , rs!Id_Banco)
          itmX.SubItems(1) = rs!Descripcion
      rs.MoveNext
     Loop
     rs.Close
    
    End With
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargarCombosEtiquetas()

On Error GoTo vError

    
    strSQL = "SELECT CT.TAG_CODIGO as 'IDX', '[' + rtrim(CT.TAG_CODIGO) + '] ' + CT.DESCRIPCION as 'ItmX'" _
       & " FROM CRD_TAGS CT INNER JOIN CRD_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
       & " INNER JOIN CRD_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
       & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
       & "' order by CT.TAG_CODIGO"
    
    Call sbCbo_Llena_New(cboEtiquetas, strSQL, False, True)
    Call sbCbo_Copia(cboEtiquetas, cboEtiquetasFiltro)
    
    cboEtiquetasFiltro.AddItem "Ninguna"
    cboEtiquetas.AddItem " "
    
    cboEtiquetas.Text = " "
    cboEtiquetasFiltro.Text = "Ninguna"
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbParametrosTagRevision()

On Error GoTo vError
    
    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '26'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRevision = rs.Fields(0)
    End If
    rs.Close
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxNombreUsuario() As String
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select DESCRIPCION from USUARIOS where Nombre = '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxNombreUsuario = rs.Fields(0)
    Else
        fxNombreUsuario = Empty
    End If

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function





Private Sub lswErrores_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
            If txtObservacion = Empty Then
              txtObservacion.Text = "-" & Item.SubItems(3)
            Else
              txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & Item.SubItems(3)
            End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert CRD_ANALISIS_ERRORESREG (ID_SOLICITUD,ID_ERROR) values(" & mOperacion _
             & "," & Item.Text & ")"
             
      If txtObservacion = Empty Then
        txtObservacion.Text = "-" & Item.SubItems(3)
      Else
        txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & Item.SubItems(3)
      End If
      
    Else
      strSQL = "delete CRD_ANALISIS_ERRORESREG where ID_SOLICITUD = " & mOperacion _
             & " and ID_ERROR = " & Item.Text
             
     Call sbCargarObservacion
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargarObservacion()
Dim i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from CRD_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtObservacion = rs.Fields(0) & vbNewLine
    Else
        txtObservacion = Empty
    End If
    
    For i = 1 To lswErrores.ListItems.Count
        If lswErrores.ListItems(i).Checked = True Then
            If lswErrores.ListItems(i).SubItems(2) = "N" Then
                If txtObservacion = Empty Then
                    txtObservacion.Text = "-" & lswErrores.ListItems(i).SubItems(3)
                Else
                    txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & lswErrores.ListItems(i).SubItems(3)
                End If
            End If
        End If
    Next
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargarListaErrores()

If mOperacion = Empty Then
    Exit Sub
End If

With lswErrores.ColumnHeaders
    .Clear
    .Add , , "Error Id", 1300
    .Add , , "Descripción", 3300
    .Add , , "Aplicado ?", 1500, vbCenter
    .Add , , "Mensaje", 5000
End With

With lswErrores
 .ListItems.Clear
 .Checkboxes = True
  
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE" _
        & " from CRD_ANALISIS_ERRORES E left join CRD_ANALISIS_ERRORESREG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.ID_SOLICITUD = " & mOperacion _
        & " where E.ACTIVO = '1'" _
        & " order by E.ID_ERROR"
        
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With


End Sub

Private Sub lswFiadores_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Call sbCargaDatosFiadores(Trim(Item.Text), Item.SubItems(2))
End Sub


Private Sub sbCargaPersonas()

    On Error GoTo vError
    
    lblFiadorMembresia.Caption = Empty
    lblFiadorIngreso.Caption = Empty
    lblFiadorInstitucion.Caption = Empty
    lblFiadorSalLiquido.Caption = Empty
    lblFiadorLiqCFianza.Caption = Empty
    lblFiadorLiqSFianza.Caption = Empty
    lblProvincia.Caption = Empty
    lblCanton.Caption = Empty
    lblDistrito.Caption = Empty
    lblDireccion.Caption = Empty
    lblPreanalisis = Empty
    
    
   strSQL = "SELECT S.CEDULA, S.NOMBRE, Est.Descripcion as 'EstadoActual', isnull(El.Descripcion,'') as 'EstadoLaboral'" _
          & "  FROM SOCIOS S" _
          & "    inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.Cod_Estado" _
          & "    left join AFI_ESTADO_LABORAL  El on S.ESTADOLABORAL  = EL.ESTADO_LABORAL" _
          & "  WHERE S.CEDULA = '" & Trim(txtCedula.Text) & "'"
 

    lswFiadores.ListItems.Clear
    
 With lswFiadores.ColumnHeaders
    .Clear
    .Add , , "Cédula", 1800
    .Add , , "Nombre", 3500
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Calidad", 1800, vbCenter
    .Add , , "Est.Lab.", 1800, vbCenter
 End With
    
    'Ingresa Deudor
    Call OpenRecordSet(rs, strSQL)
    With lswFiadores.ListItems
            Set itmX = .Add(, , rs!Cedula)
                itmX.SubItems(1) = rs!Nombre
                itmX.SubItems(2) = "Deudor"
                itmX.SubItems(3) = rs!EstadoActual
                itmX.SubItems(4) = rs!estadoLaboral
    End With
    
    Call sbCargaDatosFiadores(rs!Cedula, "Deudor")
    

   strSQL = "SELECT F.CEDULAF AS CEDULA, S.NOMBRE, Est.Descripcion as 'EstadoActual', isnull(El.Descripcion,'') as 'EstadoLaboral'" _
          & "  FROM FIADORES F INNER JOIN SOCIOS S ON F.CEDULAF = S.CEDULA" _
          & "    inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.Cod_Estado" _
          & "    left join AFI_ESTADO_LABORAL  El on S.ESTADOLABORAL  = EL.ESTADO_LABORAL" _
          & "  WHERE ID_SOLICITUD =" & mOperacion


    'Ingresa Fiadores
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswFiadores.ListItems
            Set itmX = .Add(, , rs!Cedula)
                itmX.SubItems(1) = rs!Nombre
                itmX.SubItems(2) = "Fiador"
                itmX.SubItems(3) = rs!EstadoActual
                itmX.SubItems(4) = rs!estadoLaboral
       End With
       rs.MoveNext
     Loop
     rs.Close
     
    lswFiadores.SetFocus
     
    Exit Sub
    
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
      
End Sub


Private Sub SSTabDetalle_Click(PreviousTab As Integer)
    lblSuma.Visible = False

    Select Case SSTabDetalle.Tab
    Case 0
        Call sbCargaDatosCredito
        Call sbCargaPatrimonio
    Case 1
        vGrid.MaxRows = 0
        Call sbCargaPersonas
    Case 2
        lblDeudasTotal.Caption = Format(0, "Standard")
        lblDeudasCuota.Caption = Format(0, "Standard")
        lblDeducciones.Caption = Format(0, "Standard")
        vGridDeudas.MaxRows = 0
        Call sbCargaDeudas
        
        lblFianzasMonto.Caption = Format(0, "Standard")
        lblFianzasSaldo.Caption = Format(0, "Standard")
        lblFianzasCuota.Caption = Format(0, "Standard")
        Call sbCargaFianzas

    Case 3
        lblRefundeMonto.Caption = Format(0, "Standard")
        lblRefundeCuota.Caption = Format(0, "Standard")
        Call sbCargaRefundiciones

        lblDesembolsoMonto.Caption = Format(0, "Standard")
        Call sbCargaDesembolsos

    End Select
End Sub



Private Sub sbIncluirEtiquetas()

Dim Linea As String

On Error GoTo vError
    If mOperacion = Empty Then
        Exit Sub
    End If
    
    strSQL = "select codigo from reg_creditos where id_solicitud = " & mOperacion
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        Linea = rs.Fields(0)
    Else
        MsgBox ("El número de operación no existe")
        Exit Sub
    End If
    
    Call sbCrdOperacionTags(CDbl(mOperacion), Linea, cboEtiquetas.ItemData(cboEtiquetas.ListIndex), "", txtObservacion.Text)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAplicarErrores()


On Error GoTo vError
    If mOperacion = Empty Then
        Exit Sub
    End If
    
    strSQL = "update CRD_ANALISIS_ERRORESREG SET APLICADO = 'S' WHERE ID_SOLICITUD = " & mOperacion
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCambiaCreditoRevisado()

On Error GoTo vError
    If mOperacion = Empty Then
        Exit Sub
    End If
    
    strSQL = "update REG_CREDITOS SET ANALISTAS_REVISION = '1' WHERE ID_SOLICITUD = " & mOperacion
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLimpiarDatosCreditos(ByVal Todo As Boolean)
On Error GoTo vError
    If Todo = True Then
        txtOperacion.Text = Empty
    End If

    vGridSeguimiento.MaxRows = 0
    'vGridSolicitudes.MaxRows = 0
    lswFiadores.ListItems.Clear
    
    txtObservacion.Text = Empty
    txtCedula.Text = Empty
    txtNombre.Text = Empty
    
    lblMontoApr.Caption = Empty
    lblMonto_Girado.Caption = Empty
    lblMontoRefundicion.Caption = Empty
    lblMontoDesembolsos.Caption = Empty
    lblCuotaRefundicion.Caption = Empty
    lblCuotaDesembolsos.Caption = Empty
    lblTotalCuotas.Caption = Empty
    lblCuota.Caption = Empty
    lblDiferenciaCuota.Caption = Empty
    txtObservacion.Text = Empty
    
    lblFiadorMembresia.Caption = Empty
    lblFiadorIngreso.Caption = Empty
    lblFiadorInstitucion.Caption = Empty
    lblFiadorSalLiquido.Caption = Empty
    lblFiadorLiqCFianza.Caption = Empty
    lblFiadorLiqSFianza.Caption = Empty
    lblPreanalisis.Caption = Empty
    lblFiadorLiqCFianzaPorc.Caption = Empty
    lblFiadorLiqSFianzaPorc.Caption = Empty
    
    cboEtiquetas.Text = " "
    txtObservacion.Text = Empty
    lblSuma.Visible = False
    mOperacion = Empty
    
    
    tcMain.Item(0).Selected = True
    
    SSTabDetalle.Tab = 0
    
    lblAporteObrero.Caption = Format(0, "Standard")
    lblAportePatronal.Caption = Format(0, "Standard")
    lblCapitalizacion.Caption = Format(0, "Standard")
    lblAhorrosExtra.Caption = Format(0, "Standard")
    lblAhorrosTotal.Caption = Format(0, "Standard")
    lblAhorrosFCorte.Caption = Format(0, "Standard")
    lblAhorrosFecha.Caption = Format(0, "Standard")
    lblDisponibleBruto.Caption = Format(0, "Standard")
    lblSaldoPrestamos.Caption = Format(0, "Standard")
    lblDisponible.Caption = Format(0, "Standard")
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaDatosCredito(Optional ByVal Num_Operacion As String = Empty)

On Error GoTo vError

    If mOperacion = Empty Then Exit Sub
    
    Me.MousePointer = vbHourglass

    strSQL = "select R.cedula,S.nombre,R.ID_SOLICITUD,G.DESCRIPCION AS GARANTIA," _
        & " R.MONTOAPR,R.CUOTA,R.MONTO_GIRADO, dbo.fxCrdDesembolsosOperacion(R.ID_SOLICITUD) as MONTODESEMBOLSOS, dbo.fxCrdRefundicionesOperacion(R.ID_SOLICITUD) as MONTOREFUNDICION, " _
        & " dbo.fxCrdRefundicionesCuotaOperacion(R.ID_SOLICITUD) as REFUNDICIONESCUOTA, 0 as DESEMBOLSOSCUOTA " _
        & " from reg_creditos R " _
        & " inner join socios S on S.cedula = R.cedula " _
        & " inner join AFI_ESTADOS_PERSONA E on S.ESTADOACTUAL = E.COD_ESTADO " _
        & " inner join CRD_GARANTIA_TIPOS G on R.GARANTIA = G.GARANTIA " _
        & " where R.id_solicitud = " & mOperacion
        
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
    
        txtCedula.Text = rs!Cedula
        txtNombre.Text = rs!Nombre
        
        lblGarantia.Caption = rs!Garantia
        txtOperacion = rs!Id_Solicitud
        lblMontoApr = Format(IIf(IsNull(rs!montoapr), "", rs!montoapr), "Standard")
        lblCuota = Format(IIf(IsNull(rs!Cuota), "", rs!Cuota), "Standard")
        lblMonto_Girado = Format(IIf(IsNull(rs!monto_girado), "", rs!monto_girado), "Standard")
        lblMontoDesembolsos = Format(IIf(IsNull(rs!MONTODESEMBOLSOS), "", rs!MONTODESEMBOLSOS), "Standard")
        lblMontoRefundicion = Format(IIf(IsNull(rs!MONTOREFUNDICION), "", rs!MONTOREFUNDICION), "Standard")
        lblCuotaRefundicion = Format(IIf(IsNull(rs!REFUNDICIONESCUOTA), "", rs!REFUNDICIONESCUOTA), "Standard")
        lblCuotaDesembolsos = Format(IIf(IsNull(rs!DESEMBOLSOSCUOTA), "", rs!DESEMBOLSOSCUOTA), "Standard")
        
        lblTotalCuotas = Format(IIf(IsNull(rs!REFUNDICIONESCUOTA + rs!DESEMBOLSOSCUOTA), "", rs!REFUNDICIONESCUOTA + rs!DESEMBOLSOSCUOTA), "Standard")
        
        lblDiferenciaCuota = Format(CDbl(lblCuota.Caption) - CDbl(lblTotalCuotas.Caption), "Standard")
        
        If lblDiferenciaCuota > 0 Then
            lblDiferenciaCuota.BackColor = &HFF&
            lblDiferenciaCuota.ForeColor = &H8000000F
        Else
            lblDiferenciaCuota.BackColor = &HC0FFC0
            lblDiferenciaCuota.ForeColor = &H80000012
        End If
        
    End If
    rs.Close
        
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaDatosFiadores(ByVal Cedula As String, ByVal Tipo As String)

On Error GoTo vError
    Me.MousePointer = vbHourglass

    mPreanalisis = fxPreanalisisOperacion

    strSQL = "select dbo.fxEC_Membresia(S.cedula,dbo.MyGetdate()) as Membresia, " _
        & " S.FECHAINGRESO, I.DESCRIPCION as INSTITUCION, ISNULL(P.SALARIO_LIQUIDO,0) AS SALARIO_LIQUIDO, " _
        & " ISNULL(LIQUIDEZ_SIMPLE,0) AS LIQUIDEZ_SIMPLE, ISNULL(LIQUIDEZ_CFIANZAS,0) as LIQUIDEZ_CFIANZAS," _
        & " PR.DESCRIPCION AS PROVINCIA, C.DESCRIPCION AS CANTON, D.DESCRIPCION AS DISTRITO, S.DIRECCION, " _
        & " ISNUll(DEVENGADO_MES,0) as DEVENGADO_MES, P.COD_PREANALISIS " _
        & " from Socios S inner join INSTITUCIONES I on  S.COD_INSTITUCION = I.COD_INSTITUCION " _
        & " left join PROVINCIAS PR ON S.PROVINCIA = PR.PROVINCIA left join CANTONES C ON S.CANTON = C.CANTON AND S.PROVINCIA = C.PROVINCIA" _
        & " left join DISTRITOS D ON S.DISTRITO = D.DISTRITO AND S.CANTON = D.CANTON AND S.PROVINCIA = D.PROVINCIA" _
        & " left join CRD_PREA_PREANALISIS P on S.cedula = P.cedula "
        
        If Tipo = "Fiador" Then
            strSQL = strSQL & "and P.COD_PREANALISIS_REF = '" & mPreanalisis
        Else
            strSQL = strSQL & "and P.COD_PREANALISIS = '" & mPreanalisis
        End If
        
        strSQL = strSQL & "' where S.CEDULA = '" & Trim(Cedula) & "'"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
    
        lblFiadorMembresia = IIf(IsNull(rs!Membresia), "", rs!Membresia)
        lblFiadorIngreso = IIf(IsNull(rs!FechaIngreso), "", rs!FechaIngreso)
        lblFiadorInstitucion = IIf(IsNull(rs!Institucion), "", rs!Institucion)
        lblFiadorSalLiquido = IIf(IsNull(rs!SALARIO_LIQUIDO), "", Format(rs!SALARIO_LIQUIDO, "Standard"))
        lblFiadorLiqCFianza = IIf(IsNull(rs!LIQUIDEZ_CFIANZAS), "", Format(rs!LIQUIDEZ_CFIANZAS, "Standard"))
        lblFiadorLiqSFianza = IIf(IsNull(rs!LIQUIDEZ_SIMPLE), "", Format(rs!LIQUIDEZ_SIMPLE, "Standard"))
        
        If rs!DEVENGADO_MES <> 0 Then
            lblFiadorLiqSFianzaPorc = Format(CDbl(rs!LIQUIDEZ_SIMPLE) / CDbl(rs!DEVENGADO_MES), "Standard") * 100
            lblFiadorLiqCFianzaPorc = Format(CDbl(rs!LIQUIDEZ_CFIANZAS) / CDbl(rs!DEVENGADO_MES), "Standard") * 100
        Else
            lblFiadorLiqSFianzaPorc = 0
            lblFiadorLiqCFianzaPorc = 0
        End If
        
        lblProvincia = IIf(IsNull(rs!Provincia), "", rs!Provincia)
        lblCanton = IIf(IsNull(rs!Canton), "", rs!Canton)
        lblDistrito = IIf(IsNull(rs!distrito), "", rs!distrito)
        lblDireccion = IIf(IsNull(rs!Direccion), "", rs!Direccion)
        
        lblPreanalisis = IIf(IsNull(rs!cod_PreAnalisis), "", rs!cod_PreAnalisis)
        
    End If
    rs.Close
    
    If Not lblPreanalisis = Empty Then
        Call CargarGridClasificacion(Cedula, lblFiadorLiqCFianzaPorc, lblPreanalisis)
    End If
        
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbCargarGridSeguimiento(ByVal Operacion As String)


On Error GoTo vError
    If mOperacion = Empty Then Exit Sub

    Me.MousePointer = vbHourglass

    strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO from CRD_OPERACION_TAGS OT" _
           & " inner join CRD_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO where OT.ID_SOLICITUD = " & Operacion
            
    vGridSeguimiento.MaxCols = 4
    vGridSeguimiento.MaxRows = 0


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!Registro_Usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!NOTAS), "", rs!NOTAS)
    
    vGridSeguimiento.col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridSeguimiento.col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Usuario), "", rs!Registro_Usuario)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
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
    Select Case Item.Index
    Case 2
        Call sbCargarGridSeguimiento(mOperacion)
    Case 3
        Call sbCargarListaErrores
    End Select
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then tcMain.SetFocus
End Sub

Private Sub txtOperacion_LostFocus()
    If Not txtOperacion.Text = Empty Then
        If fxExisteOperacion = 0 Then
            Call sbLimpiarDatosCreditos(True)
            vGridSolicitudes.MaxRows = 0
            Exit Sub
        Else
            Call sbLimpiarDatosCreditos(False)
            mOperacion = txtOperacion
        End If
        If fxValidaRevision = True Then
            MsgBox ("La operación digitada ya ha sido revisada")
        End If
        Call sbCargarListaSolicitudes(mOperacion)
        Call sbCargaDatosCredito
        Call sbCargaPatrimonio

        tcMain.Item(1).Selected = True
    End If
End Sub

Private Function fxExisteOperacion() As Integer

On Error GoTo vError

    If Not IsNumeric(txtOperacion) Then
        fxExisteOperacion = 0
        Exit Function
    End If
    
    strSQL = "select isnull(count(*),0) from REG_CREDITOS WHERE ID_SOLICITUD = " & txtOperacion
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxExisteOperacion = rs.Fields(0)
    Else
        fxExisteOperacion = 0
    End If
    rs.Close

    Exit Function
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Function


Private Function fxValidaRevision() As Boolean

On Error GoTo vError

    If Not IsNumeric(txtOperacion) Then
        fxValidaRevision = False
        Exit Function
    End If
    
    strSQL = "select isnull(ANALISTAS_REVISION,0) from REG_CREDITOS WHERE ID_SOLICITUD = " & txtOperacion
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        If rs.Fields(0) = 1 Then
            fxValidaRevision = True
        Else
            fxValidaRevision = False
        End If
    Else
        fxValidaRevision = False
    End If
    rs.Close

    Exit Function
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Function


Private Sub vGridDeudas_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Integer, TotalSaldo As Double

    vGridDeudas.col = 2
    TotalSaldo = 0

    For i = 1 To vGridDeudas.MaxCols
        vGridDeudas.Row = i
        If vGridDeudas.Value = vbChecked Then
            vGridDeudas.col = 7
            TotalSaldo = TotalSaldo + vGridDeudas.Value
        End If
        vGridDeudas.col = 2
    Next i
    
    If TotalSaldo > 0 Then
        lblSuma = Format(TotalSaldo, "Standard")
        lblSuma.Visible = True
    Else
        lblSuma.Visible = False
    End If
    
End Sub

Private Sub vGridSolicitudes_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    vGridSolicitudes.col = 2
    vGridSolicitudes.Row = Row
    
    Call sbLimpiarDatosCreditos(True)
    
    mOperacion = vGridSolicitudes.Text
    tcMain.Item(0).Selected = True
    
    If Len(Trim(mOperacion)) > 0 Then
        
        SSTabDetalle.Tab = 0
        
        Call sbCargaDatosCredito
        Call sbCargaPatrimonio
         
        tcMain.Item(1).Selected = True
         
    End If
    
End Sub

Public Sub sbCargaGridCheckIni(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
'Procedimiento para cargar grids con el check en la primera columna
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

    vGrid.MaxCols = vGridMaxCol + 1
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      vGrid.Row = vGrid.MaxRows
      For i = 2 To vGrid.MaxCols
        vGrid.col = i
        vGrid.Text = CStr(rs.Fields(i - 2).Value & "")
      Next i
      vGrid.MaxRows = vGrid.MaxRows + 1
      rs.MoveNext
    Loop
    rs.Close
    Exit Sub

vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Function fxFondoSolPreanalisis() As Double

On Error GoTo vError

    strSQL = "select isnull(valor,0) from CRD_PREA_PARAMETROS WHERE COD_PARAMETRO = '13'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxFondoSolPreanalisis = rs.Fields(0)
    Else
        fxFondoSolPreanalisis = 0
    End If
    rs.Close

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Function fxPreanalisisOperacion() As String

On Error GoTo vError

    strSQL = "select isnull(COD_PREANALISIS,0) from CRD_PREA_PREANALISIS WHERE TIPO_PREANALISIS = 'E' and ID_SOLICITUD = " & mOperacion
    
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxPreanalisisOperacion = rs.Fields(0)
    Else
        fxPreanalisisOperacion = 0
    End If
    rs.Close
    
Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Sub sbCargaPatrimonio()

Dim porc_OBR As Double, porc_CAP As Double

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    ''Carga porcentaje del aporte obrero
    strSQL = "select isnull(PORCENTAJE,0) from CRD_GARANTIAS_PATRIMONIO where TIPO = 'OBR'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        porc_OBR = rs.Fields(0)
    End If
    rs.Close

    ''Carga porcentaje de la capitalizacion
    strSQL = "select isnull(PORCENTAJE,0) from CRD_GARANTIAS_PATRIMONIO where TIPO = 'CAP'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        porc_CAP = rs.Fields(0)
    End If
    rs.Close

    strSQL = "SELECT isnull(A.APORTE,0) as APORTE, isnull(A.AHORRO,0) as AHORRO, isnull(A.CAPITALIZA,0)as CAPITALIZA, A.FECAPORTE, A.FECAHORRO, " _
        & " A.FECEXTRA, A.FECCAPITALIZA,(isnull(SUM(F.APORTES),0) + isnull(SUM(F.RENDIMIENTO),0)) AS FND_AHORROS " _
        & " FROM AHORRO_CONSOLIDADO A LEFT JOIN FND_CONTRATOS F ON A.CEDULA = F.CEDULA AND F.ESTADO = 'A' WHERE A.CEDULA = '" & txtCedula.Text & "'" _
        & " GROUP BY A.APORTE, A.AHORRO, A.CAPITALIZA, A.FECAPORTE, A.FECAHORRO, A.FECEXTRA, A.FECCAPITALIZA"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
    
        lblAportePatronal = Format(IIf(IsNull(rs!Aporte), 0, rs!Aporte), "Standard")
        lblAporteObrero = Format(IIf(IsNull(rs!ahorro), 0, rs!ahorro), "Standard")
        lblAhorrosExtra = Format(IIf(IsNull(rs!FND_AHORROS), 0, rs!FND_AHORROS), "Standard")
        lblCapitalizacion = Format(IIf(IsNull(rs!capitaliza), 0, rs!capitaliza), "Standard")
        lblAhorrosFCorte = IIf(IsNull(rs!fecaporte), "", rs!fecaporte)
        
        lblAhorrosTotal = Format(CDbl(lblAportePatronal.Caption) + CDbl(lblAporteObrero) + CDbl(lblAhorrosExtra) + CDbl(lblCapitalizacion), "Standard")
        
    End If
    rs.Close
    
    lblAhorrosFecha = Format(CDbl(lblAporteObrero) + CDbl(lblCapitalizacion), "Standard")
    lblDisponibleBruto = Format((CDbl(lblAporteObrero) * (porc_OBR / 100)) + (CDbl(lblCapitalizacion) * (porc_CAP / 100)), "Standard")
    
    strSQL = "Select isnull(sum(saldo),0) from reg_creditos where cedula = '" & Trim(txtCedula.Text) _
        & "' and saldo > 0 and proceso = 'N' and estado = 'A' and garantia = 'A'"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblSaldoPrestamos = Format(rs.Fields(0), "Standard")
    Else
        lblSaldoPrestamos = Format(0, "Standard")
    End If
    rs.Close
    
    lblDisponible = Format(CDbl(lblDisponibleBruto) - CDbl(lblSaldoPrestamos), "Standard")
    
    If lblDisponible < 0 Then
        lblDisponible.BackColor = &HFF&
        lblDisponible.ForeColor = &H8000000F
    Else
        lblDisponible.BackColor = &HC0FFC0
        lblDisponible.ForeColor = &H80000012
    End If
            
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaDeudas()

On Error GoTo vError
    Me.MousePointer = vbHourglass

        strSQL = "select SUM(saldo) as SaldoTotal, SUM(cuota) as CuotaTotal from reg_creditos " _
            & " where saldo > 0 and estado = 'A' and cedula = '" & txtCedula.Text & "'"
            
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            lblDeudasTotal = Format(IIf(IsNull(rs!SaldoTotal), 0, rs!SaldoTotal), "Standard")
            lblDeudasCuota = Format(IIf(IsNull(rs!CuotaTotal), 0, rs!CuotaTotal), "Standard")
        End If
        rs.Close
        
        mPreanalisis = fxPreanalisisOperacion
        
        If CDbl(mPreanalisis) > 0 Then
            strSQL = "select ISNULL(SUM(CUOTA_MENSUAL),0) AS DEDUCCIONES from CRD_PREA_DETALLE_DEDUC WHERE COD_PREANALISIS = '" & mPreanalisis & "'"
            Call OpenRecordSet(rs, strSQL)
            If Not rs.EOF Then
                lblDeducciones = Format(IIf(IsNull(rs!DEDUCCIONES), 0, rs!DEDUCCIONES), "Standard")
            End If
            rs.Close
        End If

'        strSQL = "select R.id_solicitud,R.codigo,R.PLAZO,R.montoapr,R.saldo,R.cuota,R.PRIDEDUC,isnull(M.intc+M.intm+M.amortiza,0) as MoraMnt,G.DESCRIPCION " _
'               & "from reg_creditos R left join vista_morosidad M on R.id_solicitud = M.id_solicitud inner join CRD_GARANTIA_TIPOS G on R.GARANTIA = G.GARANTIA " _
'               & "where R.saldo > 0 and R.estado = 'A' and R.cedula ='" & lblCedula & "'"
'
'        Call sbCargaGrid(vGridDeudas, 9, strSQL)
'        vGridDeudas.MaxRows = vGridDeudas.MaxRows - 1

        Call sbCargaGridDeudas
        
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaFianzas()

On Error GoTo vError
    Me.MousePointer = vbHourglass

    strSQL = "select isnull(sum(R.montoapr),0) as Monto, isnull(sum(R.cuota),0) as Cuota, isnull(sum(R.saldo),0) as Saldo " _
        & "  from reg_creditos R where R.saldo > 0 and R.estado = 'A' and R.id_solicitud " _
        & " in(select id_solicitud from fiadores where cedulaf = '" & txtCedula.Text & "' and firma = 'S')"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblFianzasMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
        lblFianzasSaldo = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "Standard")
        lblFianzasCuota = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
    End If
    rs.Close
    
    strSQL = "select R.id_solicitud,R.codigo,dbo.fxCRDNumFiadores(R.id_solicitud) as NFiadores" _
       & ",R.montoapr,R.saldo,R.cuota,S.cedula,S.nombre,isnull(M.cuota,0) as MoraCta,M.intc+M.intm+M.amortiza as MoraMnt ,R.proceso" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " left join vista_morosidad M on R.id_solicitud = M.id_solicitud" _
       & " Where R.estado = 'A' and R.id_solicitud in(select id_solicitud from fiadores where cedulaf = '" & Trim(txtCedula.Text) & "')"
           
    vGridFianzas.MaxRows = 0
    Call sbCargaGrid(vGridFianzas, 8, strSQL)
    vGridFianzas.MaxRows = vGridFianzas.MaxRows - 1
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaRefundiciones()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    Me.MousePointer = vbHourglass

    strSQL = "select isnull(sum(RE.monto),0) as Monto, isnull(sum(R.Cuota),0) as Cuota from refundiciones RE " _
        & " inner join reg_creditos R on RE.id_solicitud = R.id_solicitud " _
        & " where RE.id_solicitudr = '" & Trim(mOperacion) & "'"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblRefundeMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
        lblRefundeCuota = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
    End If
    rs.Close
    
    strSQL = "select R.id_solicitud,R.codigo,R.plazo,R.montoapr,RE.monto,R.cuota,R.garantia" _
       & " from reg_creditos R inner join refundiciones RE on R.id_solicitud = RE.id_solicitud" _
       & " where RE.id_solicitudr = '" & Trim(mOperacion) & "'"

    vGridRefundiciones.MaxRows = 0
    Call sbCargaGrid(vGridRefundiciones, 7, strSQL)
    vGridRefundiciones.MaxRows = vGridRefundiciones.MaxRows - 1
        
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaDesembolsos()

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    mPreanalisis = Empty
    mPreanalisis = fxPreanalisisOperacion
    
    If mPreanalisis = Empty Then
        Exit Sub
    End If

    strSQL = "SELECT ISNULL(SUM(MONTO),0) AS MONTO, ISNULL(SUM(CUOTA),0) AS CUOTA FROM CRD_PREA_DETALLE_DESEMBOLSOS " _
        & "WHERE COD_PREANALISIS = '" & Trim(mPreanalisis) & "'"
        
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        lblDesembolsoMonto = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
        lblDesembolsoCuota = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
    End If
    rs.Close
    
    strSQL = "SELECT DESCRIPCION,MONTO,CUOTA FROM CRD_PREA_DETALLE_DESEMBOLSOS WHERE COD_PREANALISIS = '" & Trim(mPreanalisis) & "'"

    vGridDesembolsos.MaxRows = 0
    Call sbCargaGrid(vGridDesembolsos, 3, strSQL)
    vGridDesembolsos.MaxRows = vGridDesembolsos.MaxRows - 1
        
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

'Inicio Tab Clasificacion ********************************************************************
Private Sub CargarGridClasificacion(ByVal Cedula As String, ByVal LiquidezCFianzas As Double, Expediente As String)
On Error GoTo error
Dim strSQL As String
vGrid.MaxCols = 3
        
    If Val(LiquidezCFianzas) = 0 Then
        LiquidezCFianzas = 0
    End If
    strSQL = "exec spCRDPreaClasificacion '" & Trim(Cedula) & "'," & CDbl(LiquidezCFianzas) & ",'" & Expediente & "'"
    rslocal.Open strSQL, glogon.Conection, adOpenStatic
    Call sbCargaGridLocal(vGrid, 3)
    vGrid.MaxRows = vGrid.MaxRows - 1
    rslocal.Close
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer)
Dim i As Integer
Dim RsNextCapacidad As New ADODB.Recordset
Dim RsNextEndeudamiento As New ADODB.Recordset
Dim RsNextHistorial As New ADODB.Recordset
Dim RsNextMorosidad As New ADODB.Recordset
Dim RsNextGarantia As New ADODB.Recordset
Dim vcolorcell As String


vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i
' --capacidad de pago
Set RsNextCapacidad = rslocal
Do While Not (RsNextCapacidad.EOF Or RsNextCapacidad.BOF)
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If vGrid.col = 2 Then
        vGrid.Text = "Capacidad de pago"
        vcolorcell = CStr(RsNextCapacidad.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(RsNextCapacidad.Fields(i - 1).Value)
    End If
  Next i
  RsNextCapacidad.MoveNext
  Call fxColorCell(vGrid, 1, 1, vcolorcell)
Loop
' --Evalua la capacidad de Endeudamiento
vGrid.MaxRows = 2
Set RsNextEndeudamiento = rslocal.NextRecordset
Do While Not (RsNextEndeudamiento.EOF Or RsNextEndeudamiento.BOF)
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If vGrid.col = 2 Then
        vGrid.Text = "Endeudamiento"
        vcolorcell = CStr(RsNextEndeudamiento.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(RsNextEndeudamiento.Fields(i - 1).Value)
    End If
  Next i
  RsNextEndeudamiento.MoveNext
  Call fxColorCell(vGrid, 2, 1, vcolorcell)
Loop
'--Evalua Historia laboral
vGrid.MaxRows = 3
Set RsNextHistorial = rslocal.NextRecordset
Do While Not (RsNextHistorial.EOF Or RsNextHistorial.BOF)
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If vGrid.col = 2 Then
        vGrid.Text = "Historia laboral"
        vcolorcell = CStr(RsNextHistorial.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(RsNextHistorial.Fields(i - 1).Value)
    End If
  Next i
  RsNextHistorial.MoveNext
  Call fxColorCell(vGrid, 3, 1, vcolorcell)
Loop

'--Evalua morosidad

vGrid.MaxRows = 4
Set RsNextMorosidad = rslocal.NextRecordset
Do While Not (RsNextMorosidad.EOF Or RsNextMorosidad.BOF)
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If vGrid.col = 2 Then
        vGrid.Text = "Morosidad"
        vcolorcell = CStr(RsNextMorosidad.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(RsNextMorosidad.Fields(i - 1).Value)
    End If
  Next i
  RsNextMorosidad.MoveNext
  Call fxColorCell(vGrid, 4, 1, vcolorcell)
Loop

'--Evalua Garantia
vGrid.MaxRows = 5
Set RsNextGarantia = rslocal.NextRecordset
Do While Not (RsNextGarantia.EOF Or RsNextGarantia.BOF)
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    If vGrid.col = 2 Then
        vGrid.Text = "Garantía"
        vcolorcell = CStr(RsNextGarantia.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(RsNextGarantia.Fields(i - 1).Value)
    End If
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  RsNextGarantia.MoveNext
  Call fxColorCell(vGrid, 5, 1, vcolorcell)
Loop



End Sub

Private Function fxColorCell(ByRef vGrid As Object, _
                             ByVal Row As Integer, _
                             ByVal col As Integer, _
                             ByVal strcolor As String) As String
vGrid.Row = Row
vGrid.col = col
Select Case LCase(strcolor)
    Case "rojo"
         vGrid.BackColor = &HFF&
    Case "verde"
         vGrid.BackColor = &H80FF80
    Case "amarillo"
        vGrid.BackColor = &HFFFF&
End Select
End Function

Private Sub sbCargarListaSolicitudes(Optional ByVal Num_Operacion As String = Empty)
' Carga Lista de operaciones
    Dim strSQL As String, BancosSeleccionados As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    
    Me.MousePointer = vbHourglass
    
    If Num_Operacion = Empty Then
    
        strSQL = "select  top 3000 R.id_solicitud, R.cedula,S.nombre,R.codigo, R.MONTOSOL,R.CUOTA,R.PLAZO,R.INT,case " _
                & " R.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else R.ESTADOSOL end,FECHASOL, " _
                & " RA.remesa, ISNULL(RE.USUARIO,'') AS USUARIO_REMESA " _
                & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.poliza = 'N' and C.retencion = 'N'" _
                & " inner join socios S on S.cedula = R.cedula left join CRD_REMESA_ASG RA on R.id_solicitud = RA.id_solicitud " _
                & " left join CRD_REMESAS RE on RE.REMESA = RA.REMESA " _
                & " where isnull(R.ANALISTAS_REVISION,0) = 0 and R.ESTADOSOL = 'F' and R.REFERENCIA IS NULL "
                
        If cboEtiquetasFiltro.Text <> "Ninguna" Then
           strSQL = strSQL & " and dbo.fxCRDValidaTag('" & Trim(cboEtiquetasFiltro.ItemData(cboEtiquetasFiltro.ListIndex)) & "',R.id_solicitud) > 0 "
        End If
        
        
        
        If chkSoloEspera.Value = vbChecked Then
           strSQL = strSQL & " and not R.EN_ESPERA_FECHA is null "
        End If
        
        BancosSeleccionados = fxBancosSeleccionados
        If BancosSeleccionados <> Empty Then
            strSQL = strSQL & " and R.COD_BANCO IN (" & BancosSeleccionados & ") "
        End If
        strSQL = strSQL & " order by R.id_solicitud"

    Else
        strSQL = "select R.id_solicitud, R.cedula,S.nombre,R.codigo, R.MONTOSOL,R.CUOTA,R.PLAZO,R.INT,case " _
                & " R.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else R.ESTADOSOL end,FECHASOL, " _
                & " RA.remesa, ISNULL(RE.USUARIO,'') AS USUARIO_REMESA " _
                & " from reg_creditos R " _
                & " inner join socios S on S.cedula = R.cedula left join CRD_REMESA_ASG RA on R.id_solicitud = RA.id_solicitud " _
                & " left join CRD_REMESAS RE on RE.REMESA = RA.REMESA " _
                & " where R.ESTADOSOL = 'F' and R.ID_Solicitud = " & Trim(Num_Operacion)
    End If
            
    Call sbCargaGridCheckIni(vGridSolicitudes, 12, strSQL)
    vGridSolicitudes.MaxRows = vGridSolicitudes.MaxRows - 1
    
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub

Public Function fxBancosSeleccionados() As String
'' Función busca los bancos seleccionados en la lista de bancos
    On Error GoTo vError
    Dim i As Long
        fxBancosSeleccionados = Empty
        For i = lswBancos.ListItems.Count To 1 Step -1
            If lswBancos.ListItems(i).Checked Then
                If fxBancosSeleccionados = Empty Then
                    fxBancosSeleccionados = "'" & Trim(lswBancos.ListItems(i).Text) & "'"
                Else
                    fxBancosSeleccionados = fxBancosSeleccionados & ",'" & Trim(lswBancos.ListItems(i).Text) & "'"
                End If
            End If
        Next i
        Exit Function
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Function

Private Sub sbCargaGridDeudas()

Dim vMora As Boolean
Dim i As Integer

'On Error Resume Next

Me.MousePointer = vbHourglass

vMora = False

With vGridDeudas
 .MaxRows = 0
 .MaxCols = 11
 strSQL = "exec spSIFEstadoCreditos '" & Trim(txtCedula.Text) & "'"
 
 Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .col = i
      Select Case i
        Case 1 'Status

              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
             Select Case rs!ProcesoCod
              Case "N"
       
                If Not IsNull(rs!Referencia) Then
                    If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                    .FontBold = True
                End If
        
                If rs!IndicadorCbr > 0 Then
                   .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                  .TextTip = TextTipFixed
                  .TextTipDelay = 1000
                
                  .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
                  .CellNoteIndicatorColor = vbRed
                  
                  .CellNote = "!!! Esta Operación fue Reversada de Cobro Judicial, Revise el Tab de Cobros para mayor información..!!!"
                            
                End If
              
              Case "J"
                  .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
                   vMora = True
                       
                  .TextTip = TextTipFixed
                  .TextTipDelay = 1000
                
                  .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                  .CellNoteIndicatorColor = vbRed
                  
                  .CellNote = ">> Cobro Judicial <<" & vbCrLf _
                            & "Fecha : " & Format(rs!fecha_enviaproceso, "dd/mm/yyyy") & vbCrLf _
                            & "Nota  : " & rs!observacion_proceso & ""
              
              Case "T"
                    If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                    
                    If rs!IndicadorCbr > 0 Then
                       .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                    End If
        
             End Select
             
             
             
             If Mid(rs!Estado, 1, 1) = "C" Then
                .TypePictPicture = imgSemaforos.ListImages.Item(6).Picture
             End If

            ' Si esta moroso indicar Mora siempre y cuando no este en cobro Judicial
            If rs!MoraCuota > 0 And rs!ProcesoCod <> "J" Then
              
              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
              vMora = True
            
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbBlue
              
              .CellNote = "Morosidad:  Cuotas: " & rs!MoraCuota & vbCrLf _
                        & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                        & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                        & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                        & "   Cta.+ Vieja : " & Format(rs!MoraAntigua, "####-##") & vbCrLf _
                        & "   Cta. Ultima : " & Format(rs!MoraUltima, "####-##") & vbCrLf & vbCrLf _
                        & "   Total Mora  : " & Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal, "Standard") & vbCrLf
            
            End If
        
        Case 3 'Operacion
            .Text = CStr(rs!Id_Solicitud)
        
        Case 4 'Linea
            .Text = rs!Codigo
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
            .CellNote = Trim(rs!LineaX) & vbCrLf & vbCrLf & "Formaliza: " & Format(rs!FechaForp, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!Userfor) & vbCrLf & "Oficina:" & rs!OficinaX & ""
        Case 5 'Plazo
            .Text = CStr(rs!Plazo)
        Case 6 'Monto
            .Text = Format(rs!montoapr, "Standard")
        Case 7 'Saldo
            .Text = Format(rs!Saldo, "Standard")
        Case 8 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        Case 9 'Primer Movimiento
            .Text = Format(rs!PriDeduc, "####-##")
        Case 10 'Mora
            .Text = Format(rs!MoraPrincipal + rs!MoraInt, "Standard")
        Case 11 'Garantia
            .Text = rs!Garantia
'        Case 10 'Termina
'            .Text = Format((Year(rs!Termina) & Format(Month(rs!Termina), "00")), "####-##")
'        Case 11 'Estado
'            .Text = rs!Estado
'        Case 12 'Proceso
'            .Text = rs!Proceso
'        Case 13 'Documento
'            .Text = rs!documento_referido & ""
'        Case 14 'Referencia
'            .Text = rs!referencia & ""
'        Case 15 'Tasa Original
'            .Text = Format(rs!TasaOriginal, "Standard")
'        Case 16 'Tasa Actual
'            .Text = Format(rs!Interesv, "Standard")
'        Case 17 'Plazo
'
'
'        Case 18 'Cuotas Atrasadas
'            .Text = CStr(rs!MoraCuota)
      
      End Select
    Next i

    rs.MoveNext
  Loop
  rs.Close
  
End With

''Actualiza Etiqueta del nombre con el estado de la mora
''If vEtiquetas Then
'    If vMora Then
'        lblCreditos.Caption = "Créditos en Mora"
'        Set imgCreditos.Picture = imgNo.Picture
'    Else
'        lblCreditos.Caption = "Créditos al Día"
'        Set imgCreditos.Picture = imgSi.Picture
'    End If
''End If

Me.MousePointer = vbDefault

End Sub



