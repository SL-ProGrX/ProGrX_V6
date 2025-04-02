VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_APA_ControlPagos 
   Caption         =   "Control Pagos"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Operaciones"
      TabPicture(0)   =   "frmCR_APA_ControlPagos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtpFecha_Venc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "gridBuscar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnSolicitaPago"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraFiltros"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraBuscaOperacion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Asignación Bancos"
      TabPicture(1)   =   "frmCR_APA_ControlPagos.frx":0700
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblBancos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTipo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblTotalMonto"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblMonto"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblCantidadCasos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblCasos"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btnSolicitudPagoReversa"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboTipoDesembolso"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "gridAsignacion"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lswAsignados"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "optNuevos"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cboBancos"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "tlbAsignar"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "optAsignados"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Traslados"
      TabPicture(2)   =   "frmCR_APA_ControlPagos.frx":0E08
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDesglosePago"
      Tab(2).Control(1)=   "fraBusqOptrasladas"
      Tab(2).Control(2)=   "txtCantidadRegistros"
      Tab(2).Control(3)=   "gridTraslados"
      Tab(2).Control(4)=   "btnGeneraPago"
      Tab(2).Control(5)=   "lblRegistros"
      Tab(2).ControlCount=   6
      Begin XtremeSuiteControls.GroupBox fraBuscaOperacion 
         Height          =   2775
         Left            =   120
         TabIndex        =   95
         Top             =   4080
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   4895
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtOperacionBusqueda 
            Height          =   330
            Left            =   120
            TabIndex        =   97
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBuscarOperacion 
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   2160
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Buscar"
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
            Picture         =   "frmCR_APA_ControlPagos.frx":16D9
         End
         Begin XtremeSuiteControls.ComboBox cboEstadoOpeBusq 
            Height          =   330
            Left            =   120
            TabIndex        =   100
            Top             =   1680
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   99
            Top             =   1440
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   96
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox fraFiltros 
         Height          =   3615
         Left            =   120
         TabIndex        =   87
         Top             =   480
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   6376
         _StockProps     =   79
         Caption         =   "         Filtros"
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
         Begin XtremeSuiteControls.CheckBox ckFiltrar 
            Height          =   255
            Left            =   0
            TabIndex        =   111
            Top             =   0
            Width           =   270
            _Version        =   1441793
            _ExtentX        =   476
            _ExtentY        =   450
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   330
            Left            =   120
            TabIndex        =   89
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.ComboBox cboTipoBusqueda 
            Height          =   330
            Left            =   120
            TabIndex        =   91
            Top             =   1560
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.FlatEdit txtVariacion 
            Height          =   330
            Left            =   120
            TabIndex        =   93
            Top             =   2520
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   3120
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Buscar"
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
            Picture         =   "frmCR_APA_ControlPagos.frx":1DD9
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   2160
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Variación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   90
            Top             =   1320
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Busqueda"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton btnSolicitaPago 
         Height          =   375
         Left            =   9960
         TabIndex        =   81
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Solicita Pago"
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
         Picture         =   "frmCR_APA_ControlPagos.frx":24D9
      End
      Begin VB.Frame fraDesglosePago 
         Height          =   4695
         Left            =   -71160
         TabIndex        =   17
         Top             =   1440
         Width           =   7455
         Begin MSComctlLib.Toolbar tlbInfo 
            Height          =   264
            Left            =   6120
            TabIndex        =   80
            Top             =   348
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   476
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Info"
                  Object.ToolTipText     =   "Información Adicional"
                  ImageIndex      =   15
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReActivar"
                  Object.ToolTipText     =   "ReActivar"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.Frame fraDetalleTesoreria 
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   120
            TabIndex        =   61
            Top             =   4440
            Visible         =   0   'False
            Width           =   7095
            Begin MSComctlLib.ImageList ImageList2 
               Left            =   360
               Top             =   3240
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   1
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmCR_APA_ControlPagos.frx":2C00
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblInfoTesoreriaEmision 
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
               Height          =   375
               Left            =   5040
               TabIndex        =   79
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha Emisión"
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
               Index           =   15
               Left            =   3600
               TabIndex        =   78
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label lblInfoTesoreriaDocumento 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   77
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "No. Documento"
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
               Index           =   14
               Left            =   360
               TabIndex        =   76
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label lblInfoTesoreriaEstado 
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
               Height          =   375
               Left            =   5040
               TabIndex        =   75
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Estado"
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
               Index           =   13
               Left            =   3600
               TabIndex        =   74
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label lblInfoTrasladoFecha 
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
               Height          =   375
               Left            =   5040
               TabIndex        =   73
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha Traslado"
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
               Index           =   12
               Left            =   3600
               TabIndex        =   72
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label lblInfoTrasladoUsuario 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   71
               Top             =   2880
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Usuario Traslada"
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
               Index           =   11
               Left            =   360
               TabIndex        =   70
               Top             =   2880
               Width           =   1335
            End
            Begin VB.Label lblInfoTesoreriaId 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   69
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "No. Tesorería"
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
               Index           =   10
               Left            =   360
               TabIndex        =   68
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lblInfoBanco 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   67
               Top             =   1320
               Width           =   4935
            End
            Begin VB.Label Label2 
               Caption         =   "Banco"
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
               Index           =   9
               Left            =   360
               TabIndex        =   66
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label lblInfoTipo 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   65
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo"
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
               Index           =   8
               Left            =   360
               TabIndex        =   64
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblInfoRemesa 
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
               Height          =   375
               Left            =   1680
               TabIndex        =   63
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "No. Remesa"
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
               Index           =   4
               Left            =   360
               TabIndex        =   62
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDocumentoPago 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1800
            MaxLength       =   38
            TabIndex        =   58
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtComisionPago 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1800
            TabIndex        =   39
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox txtInteresesPago 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1800
            TabIndex        =   38
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtCargosPago 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1800
            TabIndex        =   37
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox txtAbono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1800
            TabIndex        =   36
            Top             =   2280
            Width           =   1695
         End
         Begin MSComctlLib.Toolbar tlbDetallePagos 
            Height          =   312
            Left            =   1440
            TabIndex        =   18
            Top             =   4080
            Width           =   5748
            _ExtentX        =   10134
            _ExtentY        =   556
            ButtonWidth     =   2249
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgMenu"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Guardar"
                  Key             =   "Guardar"
                  Object.ToolTipText     =   "Guardar desglose de pago"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Imprime Boleta del Traslado"
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Reimprime"
                  Key             =   "Reimprime"
                  Object.ToolTipText     =   "Reimprime Boleta de Pago"
                  ImageIndex      =   15
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Revisar"
                  Key             =   "Revisar"
                  Object.ToolTipText     =   "Revisa y/o Corrige Asiento1"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cancelar"
                  Key             =   "Cancelar"
                  Object.ToolTipText     =   "Cancelar desglose del pago"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   480
            Top             =   4080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   17
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":2D24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":9586
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":FDE8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":1664A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":1CEAC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":2370E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":29F70
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":307D2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":37034
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":3D896
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":440F8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":4A95A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":4A9B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":4C5BB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":52E1D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":5967F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCR_APA_ControlPagos.frx":5FEE1
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaPago 
            Height          =   330
            Left            =   1800
            TabIndex        =   85
            Top             =   1920
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpVencePago 
            Height          =   330
            Left            =   5520
            TabIndex        =   86
            Top             =   1920
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   582
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
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   120
            X2              =   7200
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Label Label2 
            Caption         =   "Documento"
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
            Index           =   30
            Left            =   360
            TabIndex        =   59
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo al Corte"
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
            Index           =   26
            Left            =   3960
            TabIndex        =   57
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Amortización"
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
            Index           =   24
            Left            =   3960
            TabIndex        =   56
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo Hoy"
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
            Index           =   2
            Left            =   3960
            TabIndex        =   55
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Número Cuota"
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
            Index           =   22
            Left            =   3960
            TabIndex        =   54
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblSaldoPago 
            Alignment       =   1  'Right Justify
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
            Height          =   320
            Left            =   5520
            TabIndex        =   53
            Top             =   3360
            Width           =   1692
         End
         Begin VB.Label lblAmortizacionPago 
            Alignment       =   1  'Right Justify
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
            Height          =   320
            Left            =   5520
            TabIndex        =   52
            Top             =   3000
            Width           =   1692
         End
         Begin VB.Label lblNPago 
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
            Height          =   320
            Left            =   5520
            TabIndex        =   51
            Top             =   960
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Fec Movimiento"
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
            Index           =   48
            Left            =   360
            TabIndex        =   50
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Comisiones"
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
            Index           =   47
            Left            =   360
            TabIndex        =   49
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Intereses"
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
            Index           =   29
            Left            =   360
            TabIndex        =   48
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Cargos"
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
            Index           =   28
            Left            =   360
            TabIndex        =   47
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Abono"
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
            Index           =   23
            Left            =   360
            TabIndex        =   46
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Fec Vence"
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
            Index           =   38
            Left            =   3960
            TabIndex        =   45
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblUltimoSaldo 
            Alignment       =   1  'Right Justify
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
            Height          =   320
            Left            =   5520
            TabIndex        =   44
            Top             =   1560
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Tasa"
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
            Index           =   1
            Left            =   360
            TabIndex        =   43
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblTasaPago 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   1800
            TabIndex        =   42
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblSaldoAnterior 
            Alignment       =   1  'Right Justify
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
            Height          =   320
            Left            =   5520
            TabIndex        =   41
            Top             =   2640
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo anterior"
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
            Index           =   3
            Left            =   3960
            TabIndex        =   40
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   120
            X2              =   7200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Desglose de Pago a Operación :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   3012
         End
         Begin VB.Label lblNumeroOperacion 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   34
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame fraBusqOptrasladas 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   13335
         Begin VB.CheckBox chkPendientesDetalle 
            Appearance      =   0  'Flat
            Caption         =   "Mostrar solo pendientes de detallar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   11760
            TabIndex        =   60
            Top             =   0
            Width           =   1455
         End
         Begin MSComctlLib.ImageCombo cboBancoAsig 
            Height          =   345
            Left            =   6600
            TabIndex        =   26
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   609
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
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
         End
         Begin MSComctlLib.Toolbar tlbBuscaTraslados 
            Height          =   660
            Left            =   10320
            TabIndex        =   31
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1164
            ButtonWidth     =   1720
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgMenu"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Buscar"
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Busca las operaciones trasladadas"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Bancos"
                  Key             =   "Bancos"
                  Object.ToolTipText     =   "Revisar Bancos Asociados"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInicioTrasl 
            Height          =   330
            Left            =   840
            TabIndex        =   83
            Top             =   120
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpFechaFinTrasl 
            Height          =   330
            Left            =   840
            TabIndex        =   84
            Top             =   480
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.ComboBox cboEstadosAsignacion 
            Height          =   330
            Left            =   6600
            TabIndex        =   103
            Top             =   120
            Width           =   3615
            _Version        =   1441793
            _ExtentX        =   6376
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.ComboBox cboDesembolso 
            Height          =   330
            Left            =   3840
            TabIndex        =   104
            Top             =   120
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
         Begin XtremeSuiteControls.FlatEdit txtOperacion 
            Height          =   330
            Left            =   3840
            TabIndex        =   106
            Top             =   480
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Corte"
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
            Index           =   7
            Left            =   0
            TabIndex        =   33
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio"
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
            Index           =   6
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Operación"
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
            Index           =   5
            Left            =   2520
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblEstadoTraslados 
            Caption         =   "Estado"
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
            Left            =   5760
            TabIndex        =   29
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label6 
            Caption         =   "Documentos"
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
            Left            =   2520
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Bancos"
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
            Left            =   5760
            TabIndex        =   27
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton optOrdinario 
            Caption         =   "Ordinario"
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
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Value           =   -1  'True
            Width           =   1212
         End
         Begin VB.OptionButton optExtraordinario 
            Caption         =   "Extraordinario"
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
            Left            =   1440
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox txtCantidadRegistros 
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
         Left            =   -63000
         TabIndex        =   11
         Top             =   6600
         Width           =   1455
      End
      Begin VB.OptionButton optAsignados 
         Caption         =   "Asignados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   10
         Top             =   4800
         Width           =   2175
      End
      Begin MSComctlLib.Toolbar tlbAsignar 
         Height          =   312
         Left            =   -64680
         TabIndex        =   5
         Top             =   4680
         Width           =   2052
         _ExtentX        =   3625
         _ExtentY        =   556
         ButtonWidth     =   2275
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   ">> Asignar"
               Key             =   "Asignar"
               Object.ToolTipText     =   "Asigna Operaciones a Bancos Acreedores"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboBancos 
         Height          =   345
         Left            =   -72480
         TabIndex        =   3
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.OptionButton optNuevos 
         Caption         =   "Nuevos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.ListView lswAsignados 
         Height          =   2145
         Left            =   -74880
         TabIndex        =   15
         Top             =   5160
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin FPSpreadADO.fpSpread gridBuscar 
         Height          =   6132
         Left            =   2040
         TabIndex        =   22
         Top             =   1080
         Width           =   11292
         _Version        =   524288
         _ExtentX        =   19918
         _ExtentY        =   10816
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   493
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_ControlPagos.frx":66743
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gridAsignacion 
         Height          =   3612
         Left            =   -74880
         TabIndex        =   23
         Top             =   1080
         Width           =   13452
         _Version        =   524288
         _ExtentX        =   23728
         _ExtentY        =   6371
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         SpreadDesigner  =   "frmCR_APA_ControlPagos.frx":670BC
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gridTraslados 
         Height          =   4812
         Left            =   -74760
         TabIndex        =   24
         Top             =   1560
         Width           =   13212
         _Version        =   524288
         _ExtentX        =   23305
         _ExtentY        =   8488
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   492
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_ControlPagos.frx":67AF1
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha_Venc 
         Height          =   330
         Left            =   8040
         TabIndex        =   82
         Top             =   480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboTipoDesembolso 
         Height          =   330
         Left            =   -67560
         TabIndex        =   101
         Top             =   480
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.PushButton btnSolicitudPagoReversa 
         Height          =   375
         Left            =   -64440
         TabIndex        =   102
         Top             =   480
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reversa Solicitud de  Pago"
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
         Picture         =   "frmCR_APA_ControlPagos.frx":6885F
      End
      Begin XtremeSuiteControls.PushButton btnGeneraPago 
         Height          =   375
         Left            =   -67200
         TabIndex        =   105
         Top             =   6600
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Genera Pago en Bancos"
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
         Picture         =   "frmCR_APA_ControlPagos.frx":68F5F
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencimientos al:"
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
         Left            =   6240
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Pago:"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblRegistros 
         Caption         =   "Registros"
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
         Left            =   -63840
         TabIndex        =   12
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label lblCasos 
         Caption         =   "Casos"
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
         Left            =   -71160
         TabIndex        =   9
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label lblCantidadCasos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   -70320
         TabIndex        =   8
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto"
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
         Left            =   -69120
         TabIndex        =   7
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label lblTotalMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   -68280
         TabIndex        =   6
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo Documentos"
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
         Left            =   -68520
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblBancos 
         Caption         =   "Bancos"
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
         Left            =   -73320
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   13320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":69830
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":70092
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":768F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":7D156
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":839B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":8A21A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":90A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":972DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":9DB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":A43A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":AAC04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":B1466
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":B14C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":B30C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":B9929
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":C018B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_ControlPagos.frx":C69ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCod_Acreedor 
      Height          =   330
      Left            =   240
      TabIndex        =   107
      Top             =   480
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtNombreAcreedor 
      Height          =   330
      Left            =   2040
      TabIndex        =   108
      Top             =   480
      Width           =   4575
      _Version        =   1441793
      _ExtentX        =   8070
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.PushButton btnSeachMain 
      Height          =   375
      Left            =   6720
      TabIndex        =   109
      Top             =   480
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_APA_ControlPagos.frx":CD24F
   End
   Begin XtremeSuiteControls.PushButton btnActualizaSaldos 
      Height          =   375
      Left            =   7200
      TabIndex        =   110
      Top             =   480
      Width           =   2895
      _Version        =   1441793
      _ExtentX        =   5106
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actualiza Saldos y Variación"
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
      Picture         =   "frmCR_APA_ControlPagos.frx":CD94F
      ImageAlignment  =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acreedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1212
   End
   Begin VB.Menu mnuMant 
      Caption         =   "Mantenimiento"
      Visible         =   0   'False
      Begin VB.Menu mnuMantsub 
         Caption         =   "Modificar"
         Index           =   0
      End
      Begin VB.Menu mnuMantsub 
         Caption         =   "Eliminar"
         Index           =   1
      End
      Begin VB.Menu mnuMantsub 
         Caption         =   "Detalle"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCR_APA_ControlPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vCambios As Boolean, vOperacion As String
Dim vUsuario As String, vLinea As Long
Dim vEstado As String, itmX As ListItem
Dim vCambio As Boolean, vSaldoAnterior As Double

Private Sub dtpFechaMovimiento_Change()
   vCambios = True
End Sub

Private Sub dtpFechaVencimiento_Change()
   vCambios = True
End Sub


Private Sub btnActualizaSaldos_Click()
   Dim strSQL As String
   Dim i As Integer
   Dim vOperacion As String
   
   With gridBuscar
       
       For i = 1 To .MaxRows
          .Row = i
          .Col = 7
          If .Text > 0 Then
            .Col = 2
            vOperacion = Trim(.Text)
            strSQL = "exec sbActualizaVariacion '" & txtCod_Acreedor & "', '" & vOperacion & "'"
            Call ConectionExecute(strSQL)
          End If
       Next i
   
   End With
     
   Call sbGridLimpiar(gridBuscar)
   Call sbCargarOperaciones
End Sub

Private Sub btnBuscar_Click()
  Call sbCargarOperaciones
End Sub

Private Sub btnGeneraPago_Click()
  With gridTraslados
     .Row = .ActiveRow
     .Col = 17
      If .Text = "T" Then Exit Sub
      Call sbGeneraRemesa
  End With
  Call sbCargaAsignados(gridTraslados, 16, Mid(cboEstadosAsignacion.Text, 1, 1), 3)
  Call sbCargaAsignadosBancos(CInt(lswAsignados.SelectedItem.Text))
End Sub

Private Sub btnSeachMain_Click()
  Call sbCargarOperaciones
End Sub

Private Sub btnSolicitaPago_Click()
 Dim i As Long

With gridBuscar
  For i = 1 To .MaxRows
    .Col = 1
    .Row = i
    If .Value = vbChecked Then
      Call sbGuardaPago(i)
    End If
  Next i
End With

Call sbCargarOperaciones

End Sub

Private Sub btnSolicitudPagoReversa_Click()
 Dim i As Integer
 
 vOperacion = Empty
 vLinea = 0
   
   With gridAsignacion
     For i = 1 To .MaxRows
       .Row = i
       .Col = 1
       If .Value = 1 Then
       .Col = 2
       vLinea = .Text
       
       .Col = 4
       vOperacion = .Text
         Call sbReversaSolicitudPago(vOperacion, Trim(txtCod_Acreedor), vLinea)
       End If
     Next i
   End With
 Call sbCargaPagos(gridAsignacion, 15, "P", 2)
 Call sbBancosPagosAsignados
 Call sbCargarOperaciones
End Sub

Private Sub cboDesembolso_Click()
If vPaso Then Exit Sub
  Call sbCargaBancosAsignados
End Sub

Private Sub cboEstadosAsignacion_Click()
If vPaso Then Exit Sub
  If Mid(cboEstadosAsignacion.Text, 1, 1) = "T" Then
     chkPendientesDetalle.Enabled = True
  Else
     chkPendientesDetalle.Enabled = False
  End If
  Call sbCargaBancosAsignados
End Sub

Private Sub chkPendientesDetalle_Click()
   Call sbCargaAsignados(gridTraslados, 16, Mid(cboEstadosAsignacion.Text, 1, 1), 3)
End Sub


Private Sub Form_Activate()
   vModulo = 14
End Sub

Private Sub sbCargaCombo()
vPaso = True

    cboEstadoOpeBusq.Clear
    cboEstadoOpeBusq.AddItem "Activa"
    cboEstadoOpeBusq.AddItem "Cancelada"
    cboEstadoOpeBusq.Text = "Activa"

    cboTipoDesembolso.Clear
    cboTipoDesembolso.AddItem "Cheque"
    cboTipoDesembolso.AddItem "Debito Cuenta"
    cboTipoDesembolso.AddItem "Transferencia"
    cboTipoDesembolso.Text = "Debito Cuenta"
    
    cboDesembolso.Clear
    cboDesembolso.AddItem "Cheque"
    cboDesembolso.AddItem "Debito Cuenta"
    cboDesembolso.AddItem "Transferencia"
    cboDesembolso.Text = "Debito Cuenta"

    cboEstado.Clear
    cboEstado.AddItem "Cancelada"
    cboEstado.AddItem "Activa"
    cboEstado.Text = "Activa"
    
    cboEstadosAsignacion.Clear
    cboEstadosAsignacion.AddItem "Gestionadas"
    cboEstadosAsignacion.AddItem "Trasladadas"
    cboEstadosAsignacion.Text = "Gestionadas"
        
    cboTipoBusqueda.Clear
    cboTipoBusqueda.AddItem "="
    cboTipoBusqueda.AddItem "<"
    cboTipoBusqueda.AddItem ">"
    cboTipoBusqueda.AddItem "<>"
    cboTipoBusqueda.AddItem ">="
    cboTipoBusqueda.AddItem "<="
    cboTipoBusqueda.Text = "="
        
vPaso = False
        
End Sub

Private Sub sbCargarOperaciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim Fecha_Pago As String, i As Integer


On Error GoTo error

    'Consulta la lista de las Operaciones
    strSQL = "select OP.COD_ACREEDOR,OP.OPERACION, OP.CUOTA, OP.PORC_RESPONSABILIDAD, " _
           & "dbo.fxCRDAPASaldoResponsabilidad (OP.Cod_Acreedor,OP.Operacion) as RESP, " _
           & "dbo.fxCRDAPASaldoGarantias(OP.COD_ACREEDOR,OP.OPERACION) as SALDO_GARANT," _
           & "(dbo.fxCRDAPASaldoResponsabilidad(OP.COD_ACREEDOR,OP.OPERACION) - dbo.fxCRDAPASaldoGarantias(OP.COD_ACREEDOR,OP.OPERACION)) as DESV, " _
           & "OP.VARIACION,OP.FECHA_FORMALIZA, OP.DIA_DE_PAGO, OP.RESPONSABILIDAD_BASE, OP.FECHA_PROX_PAGO, " _
           & "OP.PERIOCIDAD_PAGO,case when OP.ESTADO = 'A' then 'Activa' when OP.ESTADO = 'C' then  'Cancelado' else '' end as ESTADO " _
           & "from CRD_APA_OPERACIONES OP " _
           & "left join CRD_APA_CONTROL_PAGOS CP on OP.COD_ACREEDOR = CP.COD_ACREEDOR and OP.OPERACION = CP.OPERACION " _
           & "and datepart(M,OP.FECHA_PROX_PAGO) = DATEPART(m,CP.FECHA_PAGO) and datepart(YYYY,OP.FECHA_PROX_PAGO) = DATEPART(yyyy,CP.FECHA_PAGO) " _
           & "where op.COD_ACREEDOR = '" & txtCod_Acreedor.Text & "' and op.FECHA_PROX_PAGO <= '" & Format(dtpFecha_Venc, "yyyymmdd") & "' and CP.FECHA_PAGO is null"

        If ckFiltrar.Value = 1 Then
           strSQL = strSQL & " and OP.VARIACION "
           strSQL = strSQL & cboTipoBusqueda
           strSQL = strSQL & (txtVariacion) / 100
           strSQL = strSQL & " and OP.Estado = '" & Mid(cboEstado, 1, 1) & "' "
        End If
        
       
        gridBuscar.MaxCols = 12
        gridBuscar.MaxRows = 1
        gridBuscar.Row = gridBuscar.MaxRows
        For i = 1 To gridBuscar.MaxCols
         gridBuscar.Col = i
         gridBuscar.Text = ""
        Next i
        
        Call OpenRecordSet(rs, strSQL)

        
        Do While Not rs.EOF
          
          gridBuscar.Row = gridBuscar.MaxRows
          
          
          gridBuscar.Col = 2
          gridBuscar.Text = rs!Operacion
          
          gridBuscar.Col = 3
          gridBuscar.Text = Format(rs!Cuota, "standard")
          
          gridBuscar.Col = 4
          gridBuscar.Text = Format(rs!Cuota, "standard")
          
          gridBuscar.Col = 5
          gridBuscar.Text = Format(rs!PORC_RESPONSABILIDAD, "standard")
          
          gridBuscar.Col = 6
          gridBuscar.Text = Format(rs!Resp, "standard")
          
          gridBuscar.Col = 7
          gridBuscar.Text = Format(rs!SALDO_GARANT, "standard")
          
          gridBuscar.Col = 8
          gridBuscar.Text = Format(rs!DESV, "standard")
          
          gridBuscar.Col = 9
          gridBuscar.Text = Format(rs!VARIACION, "standard")
          
          gridBuscar.Col = 10
          gridBuscar.Text = Format(rs!Fecha_Formaliza, "short date")
          
          gridBuscar.Col = 11
          gridBuscar.Text = Format(rs!Fecha_Prox_Pago, "short date")
          
          gridBuscar.Col = 12
          gridBuscar.Text = rs!Estado
          
          gridBuscar.MaxRows = gridBuscar.MaxRows + 1
          rs.MoveNext
        Loop
        rs.Close
        
    gridBuscar.MaxRows = gridBuscar.MaxRows - 1
Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbBuscaOperacionesTrasl()

End Sub

Private Sub sbBuscaOperacion()
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim Fecha_Pago As String
Dim i As Integer
  
On Error GoTo error
  
strSQL = "select OP.COD_ACREEDOR,OP.OPERACION, OP.CUOTA, OP.PORC_RESPONSABILIDAD, " _
    & "dbo.fxCRDAPASaldoResponsabilidad (OP.Cod_Acreedor,OP.Operacion) as RESP, " _
    & "dbo.fxCRDAPASaldoGarantias(OP.COD_ACREEDOR,OP.OPERACION) as SALDO_GARANT," _
    & "(dbo.fxCRDAPASaldoResponsabilidad(OP.COD_ACREEDOR,OP.OPERACION) - dbo.fxCRDAPASaldoGarantias(OP.COD_ACREEDOR,OP.OPERACION)) as DESV, " _
    & "OP.VARIACION,OP.FECHA_FORMALIZA, OP.DIA_DE_PAGO, OP.RESPONSABILIDAD_BASE, OP.FECHA_PROX_PAGO, " _
    & "OP.PERIOCIDAD_PAGO,case when OP.ESTADO = 'A' then 'Activa' when OP.ESTADO = 'C' then  'Cancelado' else '' end as ESTADO " _
    & "from CRD_APA_OPERACIONES OP " _
    & "left join CRD_APA_CONTROL_PAGOS CP on OP.COD_ACREEDOR = CP.COD_ACREEDOR and OP.OPERACION = CP.OPERACION " _
    & "and datepart(M,OP.FECHA_PROX_PAGO) = DATEPART(m,CP.FECHA_PAGO) and datepart(YYYY,OP.FECHA_PROX_PAGO) = DATEPART(yyyy,CP.FECHA_PAGO) " _
    & "where op.COD_ACREEDOR = '" & txtCod_Acreedor.Text & "' and op.FECHA_PROX_PAGO <= '" & Format(dtpFecha_Venc, "yyyymmdd") & "' and " _
    & "OP.OPERACION = '" & Trim(txtOperacionBusqueda) & "' and OP.ESTADO = '" & Mid(cboEstadoOpeBusq.Text, 1, 1) & "' "
 

 gridBuscar.MaxRows = 1
 gridBuscar.Row = gridBuscar.MaxRows
 
 For i = 1 To gridBuscar.MaxCols
  gridBuscar.Col = i
  gridBuscar.Text = ""
 Next i
 
 Call OpenRecordSet(rs, strSQL)

 
 Do While Not rs.EOF
   
   gridBuscar.Row = gridBuscar.MaxRows
   
   
   gridBuscar.Col = 2
   gridBuscar.Text = rs!Operacion
   
   gridBuscar.Col = 3
   gridBuscar.Text = Format(rs!Cuota, "standard")
   
   gridBuscar.Col = 4
   gridBuscar.Text = Format(rs!Cuota, "standard")
   
   gridBuscar.Col = 5
   gridBuscar.Text = Format(rs!PORC_RESPONSABILIDAD, "standard")
   
   gridBuscar.Col = 6
   gridBuscar.Text = Format(rs!Resp, "standard")
   
   gridBuscar.Col = 7
   gridBuscar.Text = Format(rs!SALDO_GARANT, "standard")
   
   gridBuscar.Col = 8
   gridBuscar.Text = Format(rs!DESV, "standard")
   
   gridBuscar.Col = 9
   gridBuscar.Text = Format(rs!VARIACION, "standard")
   
   gridBuscar.Col = 10
   gridBuscar.Text = Format(rs!Fecha_Formaliza, "short date")
   
   gridBuscar.Col = 11
   gridBuscar.Text = Format(rs!Fecha_Prox_Pago, "short date")
   
   gridBuscar.Col = 12
   gridBuscar.Text = rs!Estado
   
   gridBuscar.MaxRows = gridBuscar.MaxRows + 1
   rs.MoveNext
 Loop
 rs.Close
  
 gridBuscar.MaxRows = gridBuscar.MaxRows - 1
  
Exit Sub
    
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbCargaBancos()
    Dim strSQL As String
    Dim rs As New ADODB.Recordset

    strSQL = "Select ID_BANCO, DESCRIPCION" _
            & " from TES_BANCOS"

    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
        cboBancos.ComboItems.Add , rs.Fields("ID_BANCO") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
        rs.MoveNext
    Loop

    rs.Close
End Sub

Private Sub sbCargaBancosAsignados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDocumento As String

On Error GoTo vError

If txtCod_Acreedor.Text = Empty Then Exit Sub

Me.MousePointer = vbHourglass
  
Select Case cboDesembolso.Text
   Case "Cheque"
       vDocumento = "Ck"
   Case "Debito Cuenta"
       vDocumento = "Dc"
   Case "Transferencia"
       vDocumento = "Tf"
End Select
  
strSQL = " select B.ID_BANCO, B.DESCRIPCION" _
       & " from TES_BANCOS B inner join CRD_APA_CONTROL_PAGOS P on B.id_Banco = P.cod_Banco" _
       & " where P.COD_ACREEDOR = '" & txtCod_Acreedor & "'  and P.ESTADO = '" & Mid(cboEstadosAsignacion.Text, 1, 1) & "'" _
       & " and P.FECHA_REGISTRO between '" & Format(dtpFechaInicioTrasl, "yyyymmdd 00:00") & "' and '" & Format(dtpFechaFinTrasl, "yyyymmdd 23:59") & "'"
       
If txtOperacion.Text <> Empty Then
  strSQL = strSQL & " and P.OPERACION = '" & Trim(txtOperacion.Text) & "' "
End If

If vDocumento <> Empty Then
  strSQL = strSQL & " and P.TIPO_DESEMBOLSO = '" & vDocumento & "'"
End If

strSQL = strSQL & " group by B.ID_BANCO, B.DESCRIPCION"
 
 
cboBancoAsig.ComboItems.Clear
cboBancoAsig.Text = ""
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   cboBancoAsig.ComboItems.Add , rs.Fields("ID_BANCO") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
   rs.MoveNext
Loop

'If rs.RecordCount > 0 Then
'   rs.MoveFirst
'   cboBancoAsig.Text = UCase(Trim(rs!DESCRIPCION))
'End If
rs.Close
      
gridTraslados.MaxRows = 0

Me.MousePointer = vbDefault
      
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
  vModulo = 14
    
  gridBuscar.MaxCols = 12
  gridBuscar.MaxRows = 1
  gridTraslados.MaxRows = 1
  
  lblTotalMonto = 0
  txtVariacion = 0
  
  ssTab.Tab = 0
  
  Call sbCargaCombo
  Call sbCargaBancos
  Call sbCargaBancosAsignados
  
  optNuevos.Value = True
  optOrdinario.Value = True
    
  dtpFecha_Venc.Value = fxFechaServidor
  
  gridTraslados.Col = 1
  gridTraslados.Row = gridTraslados.ActiveRow
  
  fraDesglosePago.Top = 55000
  txtCantidadRegistros.Text = 100
  dtpFechaInicioTrasl = fxFechaServidor
  dtpFechaFinTrasl = fxFechaServidor
 
End Sub

Private Sub Form_Resize()
  On Error Resume Next
        
'    lblMonto.Visible = False
'    lblCantidadCasos.Visible = False
'    txtCantidadRegistros.Visible = False
'    lblEstadoTraslados.Visible = False
'    lblRegistros.Visible = False
'    btnGeneraPago.Visible = False
'    tlbAsignar.Visible = False
'    ckFiltrar.Visible = False
'    cboEstadosAsignacion.Visible = False
'    fraDesglosePago.Visible = False
'    fraBusqOptrasladas.Visible = False

    ssTab.Width = Me.Width - 500
    ssTab.Height = Me.Height - 1600
    fraFiltros.Top = gridBuscar.Top
    fraBuscaOperacion.Top = ssTab.Top + fraFiltros.Height + 100
    gridBuscar.Top = ssTab.Top
    gridBuscar.Width = ssTab.Width - fraFiltros.Width - 500
    gridBuscar.Height = ssTab.Height - 1500
    
    gridTraslados.Top = fraBusqOptrasladas.Top + fraBusqOptrasladas.Height + 100
    gridTraslados.Width = ssTab.Width - 500
    gridTraslados.Height = ssTab.Height - fraBusqOptrasladas.Height - 1000
    
    txtCantidadRegistros.Left = ssTab.Width - txtCantidadRegistros.Width - 200
    txtCantidadRegistros.Top = gridTraslados.Top + gridTraslados.Height + 100
    lblRegistros.Top = txtCantidadRegistros.Top
    lblRegistros.Left = ssTab.Width - txtCantidadRegistros.Width - lblRegistros.Width - 500
    btnGeneraPago.Left = ssTab.Width - txtCantidadRegistros.Width - lblRegistros.Width - btnGeneraPago.Width - 2500
    btnGeneraPago.Top = txtCantidadRegistros.Top
    fraDesglosePago.Top = Me.Height + 250
    
    
    gridAsignacion.Top = ssTab.Top
    gridAsignacion.Width = ssTab.Width - 200
    gridAsignacion.Height = ssTab.Height - 5000
        
    lblCasos.Top = ssTab.Top + gridAsignacion.Height + 100
    lblCasos.Left = ssTab.Width - tlbAsignar.Width - lblTotalMonto.Width - lblMonto.Width - lblCantidadCasos.Width - lblCasos.Width - 1200
    lblCantidadCasos.Top = ssTab.Top + gridAsignacion.Height + 100
    lblCantidadCasos.Left = ssTab.Width - tlbAsignar.Width - lblTotalMonto.Width - lblMonto.Width - lblCantidadCasos.Width - 800

    lblTotalMonto.Top = ssTab.Top + gridAsignacion.Height + 100
    lblTotalMonto.Left = ssTab.Width - tlbAsignar.Width - lblTotalMonto.Width - 400
    lblMonto.Top = ssTab.Top + gridAsignacion.Height + 100
    lblMonto.Left = ssTab.Width - tlbAsignar.Width - lblTotalMonto.Width - lblMonto.Width - 600

    tlbAsignar.Top = ssTab.Top + gridAsignacion.Height + 100
    tlbAsignar.Left = ssTab.Width - tlbAsignar.Width - 200

    optAsignados.Top = ssTab.Top + gridAsignacion.Height + 100

    lswAsignados.Top = gridAsignacion.Top + gridAsignacion.Height + 1000
    lswAsignados.Width = gridAsignacion.Width
    lswAsignados.Height = ssTab.Height - (gridAsignacion.Height + 2500)

    
    Select Case ssTab.Tab
       Case 0 'Buscar
           lblTotalMonto.Visible = False
           lblCantidadCasos.Visible = False
           txtCantidadRegistros.Visible = False
           lblEstadoTraslados.Visible = False
           lblRegistros.Visible = False
           btnGeneraPago.Visible = False
           tlbAsignar.Visible = False
           ckFiltrar.Visible = True
           cboEstadosAsignacion.Visible = False
           lblEstadoTraslados.Visible = False
           fraDesglosePago.Visible = False
           fraBuscaOperacion.Visible = True
           fraBusqOptrasladas.Visible = False
        Case 1 ' Asignacion de bancos
           ckFiltrar.Visible = False
           tlbAsignar.Visible = True
           lblTotalMonto.Visible = True
           lblMonto.Visible = True
           lblCantidadCasos.Visible = True
           txtCantidadRegistros.Visible = False
           lblEstadoTraslados.Visible = False
           lblRegistros.Visible = False
           btnGeneraPago.Visible = False
           cboEstadosAsignacion.Visible = False
           lblEstadoTraslados.Visible = False
           fraDesglosePago.Visible = False
           fraBuscaOperacion.Visible = False
           fraBusqOptrasladas.Visible = False
        Case 2 ' Traslados
           ckFiltrar.Visible = False
           lblEstadoTraslados.Visible = True
           lblRegistros.Visible = True
           btnGeneraPago.Visible = True
           txtCantidadRegistros.Visible = True
           tlbAsignar.Visible = False
           cboEstadosAsignacion.Visible = True
           lblEstadoTraslados.Visible = True
           fraDesglosePago.Visible = True
           fraDesglosePago.Left = gridTraslados.Left + 3800
           fraBuscaOperacion.Visible = False
           fraBusqOptrasladas.Visible = True
    End Select
        
End Sub

Private Sub gridAsignacion_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long, curTotal As Currency

If vPaso Or Col > 1 Then Exit Sub

i = IIf(IsNumeric(lblCantidadCasos.Caption), lblCantidadCasos.Caption, 0)
curTotal = IIf(IsNumeric(lblTotalMonto.Caption), CCur(lblTotalMonto.Caption), 0)

With gridAsignacion
   .Row = Row
   .Col = 1
    If .Value = vbChecked Then
       .Col = 8
       i = i + 1
       curTotal = curTotal + CCur(.Text)
    Else
       .Col = 8
       i = i - 1
       curTotal = curTotal - CCur(.Text)
    End If
End With
  
lblCantidadCasos.Caption = i
lblTotalMonto.Caption = Format(curTotal, "Standard")

End Sub

Private Sub gridAsignacion_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  If Button = 2 Then
'    Call PopupMenu(mnuMant, , x, y * 2)
'  End If
End Sub

Private Sub sbOperacionDetalle(pOperacion As String, pLinea As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 strSQL = "Select CP.*, dbo.MyGetdate() as 'FechaServer',Op.SALDO as 'SaldoHoy'" _
        & ",Ban.Descripcion as 'Banco',Tes.NDocumento,Tes.Fecha_Emision,Tes.Estado as 'Tes_Estado'" _
        & " from CRD_APA_CONTROL_PAGOS CP inner join CRD_APA_OPERACIONES Op on CP.cod_acreedor = Op.cod_acreedor and CP.Operacion = Op.Operacion" _
        & " left join Tes_Bancos Ban on CP.COD_BANCO  = Ban.id_Banco" _
        & " left join Tes_Transacciones Tes on CP.Tesoreria_Solicitud = Tes.NSolicitud" _
        & " where CP.COD_ACREEDOR = '" & Trim(txtCod_Acreedor.Text) & "' and CP.OPERACION='" & pOperacion & "' and CP.Linea =" & pLinea
 Call OpenRecordSet(rs, strSQL)
 
 lblNumeroOperacion.Caption = rs!Operacion
 lblNPago.Caption = rs!N_CUOTA
 lblNPago.Tag = rs!LINEA
 
 txtAbono = Format(rs!ABONO, "standard")
 txtDocumentoPago.Text = IIf(IsNull(rs!Documento), Empty, rs!Documento)
 lblUltimoSaldo.Caption = Format(rs!SaldoHoy, "standard")
 
 If IsNull(rs!Documento) Then
    lblSaldoAnterior.Caption = Format(rs!SaldoHoy, "standard")
    dtpVencePago.Value = rs!Fecha_Pago
    dtpFechaPago.Value = rs!FechaServer
    txtInteresesPago.Text = "0.00"
    txtComisionPago.Text = "0.00"
    lblAmortizacionPago.Caption = "0.00"
    txtCargosPago.Text = "0.00"
    lblSaldoPago.Caption = IIf(IsNull(Format(rs!SaldoHoy, "standard")), 0, Format(rs!SaldoHoy, "standard"))
    lblTasaPago.Caption = IIf(IsNull(rs!Tasa), 0, rs!Tasa)
 Else
    lblSaldoAnterior.Caption = Format(rs!Saldo + rs!Amortizacion, "standard")
    dtpVencePago.Value = rs!Fecha_Pago
    dtpFechaPago.Value = rs!Fecha_Pago
    txtInteresesPago.Text = IIf(IsNull(Format(rs!INTERESES, "standard")), 0, Format(rs!INTERESES, "standard"))
    txtComisionPago.Text = IIf(IsNull(Format(rs!COMISION, "standard")), 0, Format(rs!COMISION, "standard"))
    lblAmortizacionPago.Caption = IIf(IsNull(Format(rs!Amortizacion, "standard")), 0, Format(rs!Amortizacion, "standard"))
    txtCargosPago.Text = IIf(IsNull(Format(rs!CARGOS, "standard")), 0, Format(rs!CARGOS, "standard"))
    lblSaldoPago.Caption = IIf(IsNull(Format(rs!Saldo, "standard")), 0, Format(rs!Saldo, "standard"))
    lblTasaPago.Caption = IIf(IsNull(rs!Tasa), 0, rs!Tasa)
 End If
        
 fraDesglosePago.Left = gridTraslados.Left + 3900
 fraDesglosePago.Top = 1300
              
 If txtDocumentoPago <> Empty Then
   tlbDetallePagos.Buttons.Item(1).Enabled = False
 Else
   tlbDetallePagos.Buttons.Item(1).Enabled = True
 End If
 
 'Datos Adicionales
 lblInfoRemesa.Caption = rs!REMESA & ""
 lblInfoBanco.Caption = rs!Banco & ""
 lblInfoTesoreriaDocumento.Caption = rs!Ndocumento & ""
 lblInfoTesoreriaEmision.Caption = Format(rs!Fecha_Emision & "", "dd/mm/yyyy")
 lblInfoTesoreriaEstado.Caption = rs!Tes_Estado & ""
 lblInfoTesoreriaId.Caption = rs!Tesoreria_Solicitud & ""
 lblInfoTipo.Caption = rs!Tipo_Desembolso & ""
 lblInfoTrasladoFecha.Caption = Format(rs!Tesoreria_Fecha & "", "dd/mm/yyyy")
 lblInfoTrasladoUsuario.Caption = rs!Tesoreria_Usuario & ""
              
 rs.Close

      

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub gridTraslados_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim vEstado As String

                
                
On Error GoTo vError
                
                
With gridTraslados
  .Row = .ActiveRow
  .Col = 17
  vEstado = .Value
  
  If ButtonDown = 0 Then
    .Col = 2
    If .Value = 1 And vEstado = "T" Then
      .Col = 5
      vOperacion = .Value
      .Col = 3
      vLinea = .Text
      Call sbOperacionDetalle(vOperacion, vLinea)
    End If
  End If
      
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswAsignados_ItemClick(ByVal Item As MSComctlLib.ListItem)

   Call sbCargaAsignadosBancos(CInt(lswAsignados.SelectedItem.Text))
   tlbAsignar.Buttons.Item(1).Caption = "Desasignar Bancos"
   tlbAsignar.Buttons.Item(1).Key = "Desasignar"
   optAsignados.Value = True
   
   lblCantidadCasos.Caption = lswAsignados.SelectedItem.SubItems(2)
   lblTotalMonto.Caption = lswAsignados.SelectedItem.SubItems(3)
End Sub

Private Sub mnuMantsub_Click(Index As Integer)
Dim strSQL As String
Dim rs As New ADODB.Recordset

On Error GoTo error
  
  With gridAsignacion
    .Col = 1
    .Row = .ActiveRow
    
    If .Value = 1 Then
      .Col = 2
      vLinea = .Text
      
      .Col = 4
      vOperacion = .Text
      
        Select Case Index
          Case 0 'Modificar
          Case 1 'Eliminar
             MsgBox "Eliminar"
          Case 2 'Detalles
             MsgBox "Detalles"
        End Select
    
    End If
  End With
  
  Exit Sub
  
error:
  MsgBox fxSys_Error_Handler(Err.Description)
  
End Sub

Private Sub optAsignados_Click()
  If optAsignados.Value = True Then
     tlbAsignar.Buttons.Item(1).Caption = "Desasignar Bancos"
     tlbAsignar.Buttons.Item(1).Key = "Desasignar"
     tlbAsignar.Buttons.Item(1).Image = 6
     gridAsignacion.MaxRows = 0
  End If
End Sub

Private Sub optNuevos_Click()
  If optNuevos.Value = True Then
     tlbAsignar.Buttons.Item(1).Caption = "Asignar Bancos"
     tlbAsignar.Buttons.Item(1).Key = "Asignar"
     tlbAsignar.Buttons.Item(1).Image = 1
     Call sbCargaPagos(gridAsignacion, 15, "P", 2)
  End If
End Sub



Private Sub btnBuscarOperacion_Click()
  Call sbBuscaOperacion
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case ssTab.Tab
       Case 0 'Buscar
           lblTotalMonto.Visible = False
           lblCantidadCasos.Visible = False
           txtCantidadRegistros.Visible = False
           lblEstadoTraslados.Visible = False
           lblRegistros.Visible = False
           btnGeneraPago.Visible = False
           tlbAsignar.Visible = False
           ckFiltrar.Visible = True
           cboEstadosAsignacion.Visible = False
           lblEstadoTraslados.Visible = False
           fraDesglosePago.Visible = False
           fraBuscaOperacion.Visible = True
           fraBusqOptrasladas.Visible = False
        Case 1 ' Asignacion de bancos
           ckFiltrar.Visible = False
           tlbAsignar.Visible = True
           lblTotalMonto.Visible = True
           lblMonto.Visible = True
           lblCantidadCasos.Visible = True
           txtCantidadRegistros.Visible = False
           cboEstadosAsignacion.Visible = False
           lblEstadoTraslados.Visible = False
           lblRegistros.Visible = False
           btnGeneraPago.Visible = False
           optNuevos.Value = True
           cboEstadosAsignacion.Visible = False
           lblEstadoTraslados.Visible = False
           fraDesglosePago.Visible = False
           fraBuscaOperacion.Visible = False
           fraBusqOptrasladas.Visible = False
           
           Call sbCargaPagos(gridAsignacion, 15, "P", 2)
           Call sbBancosPagosAsignados
           
        Case 2 ' Traslados
           ckFiltrar.Visible = False
           cboEstadosAsignacion.Visible = True
           lblEstadoTraslados.Visible = True
           lblRegistros.Visible = True
           btnGeneraPago.Visible = True
           txtCantidadRegistros.Visible = True
           tlbAsignar.Visible = False
           cboEstadosAsignacion.Visible = True
           lblEstadoTraslados.Visible = True
           fraDesglosePago.Visible = True
           fraBuscaOperacion.Visible = False
           fraBusqOptrasladas.Visible = True
           gridTraslados.MaxRows = 0
    End Select
End Sub



Private Sub tlbAsignar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case tlbAsignar.Buttons.Item(1).Key
     Case "Asignar"
       If MsgBox("Desea asignar el banco a las operaciones seleccionadas", vbYesNo, "SIF : Adm. Garantías") = vbYes Then
           Call sbAsignaBanco
           Call sbCargaPagos(gridAsignacion, 15, "P", 2)
           Call sbBancosPagosAsignados

       End If
     Case "Desasignar"
       If MsgBox("Desea desasignar el banco de las operaciones seleccionadas", vbYesNo, "SIF : Adm. Garantías") = vbYes Then
           Call sbDesAsignarBanco
           Call sbCargaPagos(gridAsignacion, 15, "P", 2)
           Call sbBancosPagosAsignados
           Call optNuevos_Click
       End If
  End Select
End Sub
Private Sub sbCargaAcreedor()
 Dim strSQL As String
 Dim rs As New ADODB.Recordset
   
   strSQL = "select COD_ACREEDOR, DESCRIPCION " _
          & " from CRD_APA_ACREEDORES where COD_ACREEDOR='" & txtCod_Acreedor & "'"
   
   Call OpenRecordSet(rs, strSQL)

   If rs.EOF Then
     Exit Sub
   Else
     txtNombreAcreedor = rs!Descripcion
   End If

   rs.Close
     
End Sub

Private Sub tlbBusca_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call sbCargaAcreedor
    Call sbCargarOperaciones
    Call sbCargaPagos(gridAsignacion, 15, "P", 2)
    Call sbBancosPagosAsignados
End Sub


Private Sub sbGuardaPago(vFilas As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vAbono As Currency, vTipoPago As String
Dim vFecha_Ultimo_Pago As Date, vCuota As Currency

On Error GoTo vError
     
     
    If Not optOrdinario.Value And Not optExtraordinario Then
       MsgBox "No tiene un tipo de pago seleccionado"
       Exit Sub
    End If
     
    Select Case True
       Case optOrdinario.Value
           vTipoPago = "O"
       Case optExtraordinario.Value
           vTipoPago = "E"
    End Select
     
    With gridBuscar
       .Row = vFilas
              
       .Col = 2
       vOperacion = .Text
       
       .Col = 3
       vCuota = CCur(.Text)
       
       .Col = 4
       vAbono = CCur(.Text)
       
       .Col = 11
       If .Text = "" Then
        vFecha_Ultimo_Pago = fxFechaServidor
       Else
        vFecha_Ultimo_Pago = CDate(.Text)
       End If
       
    End With
     
     vUsuario = glogon.Usuario
     
     strSQL = "exec spCrdAPAControlPagos_A '" & vOperacion & "','" & txtCod_Acreedor.Text & "'," _
            & vAbono & ",'P','" & vUsuario & "','" & vTipoPago & "','" & Format(vFecha_Ultimo_Pago, "yyyymmdd") _
            & "'," & vCuota
     Call ConectionExecute(strSQL)
     
     If vTipoPago = "O" Then
        strSQL = "exec spAPA_ActualizaFechaProxPago '" & vOperacion & "','" & txtCod_Acreedor.Text & "'"
        Call ConectionExecute(strSQL)
     End If
     
Exit Sub

vError:

   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCrearEncabezado()
        lswAsignados.ListItems.Clear
        lswAsignados.ColumnHeaders.Clear
        lswAsignados.ColumnHeaders.Add , , "ID Banco", 500 '0
        lswAsignados.ColumnHeaders.Add , , "Banco", 3000, 0 ' 2
        lswAsignados.ColumnHeaders.Add , , "Cantidad Casos", 2000, 2 '0
        lswAsignados.ColumnHeaders.Add , , "Monto", 2000, 1 '1
End Sub

Private Sub sbDesAsignarBanco()
 Dim i As Integer
 Dim strSQL As String
 Dim rs As New ADODB.Recordset
    
    
    With gridAsignacion
       For i = 1 To .MaxRows
         .Col = 1
         .Row = i
         If .Value <> 0 Then
            .Col = 2
            vLinea = .Text
            .Col = 4
            vOperacion = .Text
            
            strSQL = "Update CRD_APA_CONTROL_PAGOS set COD_BANCO = Null, TIPO_DESEMBOLSO = Null ,ESTADO='P' " _
                   & "where COD_ACREEDOR= '" & txtCod_Acreedor.Text & "' and OPERACION='" & vOperacion & "' and LINEA='" & vLinea & "'"
            
            Call ConectionExecute(strSQL)
            
         End If
       Next i
    End With


End Sub

Private Sub sbAsignaBanco()
 'Asigna el banco del que se va a generar el pago
 Dim strSQL As String, rs As New ADODB.Recordset
 Dim vBanco As Integer, vDocumento As String
 Dim vLinea As Integer, i As Integer
 
 On Error GoTo vError
    
   If cboBancos.Text = Empty Then
      MsgBox "No se tiene ningún banco para asignar"
      Exit Sub
   Else
    vBanco = DeCodificaPrimaryKey(cboBancos.SelectedItem.Key, 1, "(id)")
   End If
      
    Select Case cboTipoDesembolso.Text
       Case "Cheque"
          vDocumento = "Ck"
       Case "Debito Cuenta"
          vDocumento = "Dc"
       Case "Transferencia"
          vDocumento = "Tf"
    End Select
  
    With gridAsignacion
       For i = 1 To .MaxRows
         .Col = 1
         .Row = i
         If .Value <> 0 Then
            .Col = 2
            vLinea = .Text
            .Col = 4
            vOperacion = .Text
            
            strSQL = "Update CRD_APA_CONTROL_PAGOS set COD_BANCO = '" & vBanco & "', TIPO_DESEMBOLSO = '" & vDocumento & "',ESTADO='G' " _
                   & "where COD_ACREEDOR= '" & txtCod_Acreedor.Text & "' and OPERACION='" & vOperacion & "' and LINEA='" & vLinea & "'"
            
            Call ConectionExecute(strSQL)
            
         End If
       Next i
    End With
    
 Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaAsignados(vGrid As Object, vGridMaxCol As Integer, vEstado As String, vColumna As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBanco As Integer, vDocumento As String
Dim vCantidadRegistros As Integer

On Error GoTo vError
  
  vCantidadRegistros = CInt(txtCantidadRegistros)
  gridTraslados.MaxRows = 0
  
  If txtCod_Acreedor.Text = Empty Or cboBancoAsig.Text = Empty Then Exit Sub
  If cboBancoAsig.ComboItems.Count <= 0 Then Exit Sub
  
  vBanco = DeCodificaPrimaryKey(cboBancoAsig.SelectedItem.Key, 1, "(id)")
  
  Select Case cboDesembolso.Text
     Case "Cheque"
         vDocumento = "Ck"
     Case "Debito Cuenta"
         vDocumento = "Dc"
     Case "Transferencia"
         vDocumento = "Tf"
  End Select
  
  strSQL = " select top(" & vCantidadRegistros & ") CP.LINEA, ACR.DESCRIPCION, CP.OPERACION, case when CP.TIPO_PAGO = 'O' then 'Ordinario' " _
         & " when CP.TIPO_PAGO = 'E' then  'ExtraOrdinario' else '' end as Pago,CP.N_CUOTA, CP.CUOTA, CP.ABONO,Op.SALDO, BANC.DESCRIPCION, " _
         & " CP.TIPO_DESEMBOLSO , CP.INTERESES, CP.COMISION, CP.CARGOS, CP.AMORTIZACION, CP.Estado " _
         & " from CRD_APA_CONTROL_PAGOS CP inner join   CRD_APA_OPERACIONES Op on CP.COD_ACREEDOR = Op.COD_ACREEDOR and CP.OPERACION = Op.OPERACION " _
         & " left join TES_BANCOS Banc on CP.COD_BANCO = Banc.ID_BANCO " _
         & " left join CRD_APA_ACREEDORES ACR on CP.COD_ACREEDOR = ACR.COD_ACREEDOR " _
         & " where CP.COD_ACREEDOR = '" & txtCod_Acreedor & "'  and CP.ESTADO = '" & vEstado & "'" _
         & " and CP.FECHA_REGISTRO between '" & Format(dtpFechaInicioTrasl, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpFechaFinTrasl, "yyyy/mm/dd") & " 23:59:59'"
         
         If vBanco <> Empty Then
           strSQL = strSQL & " and CP.COD_BANCO = '" & vBanco & "'"
         End If
         
         If txtOperacion.Text <> Empty Then
           strSQL = strSQL & " and CP.OPERACION = '" & Trim(txtOperacion.Text) & "' "
         End If
         
         If vDocumento <> Empty Then
           strSQL = strSQL & " and CP.TIPO_DESEMBOLSO = '" & vDocumento & "'"
         End If
         
        If vEstado = "T" And chkPendientesDetalle.Value = vbChecked Then
           strSQL = strSQL & " and CP.DOCUMENTO is null"
        End If
          
  Call sbCargaGridCheck(vGrid, vGridMaxCol, strSQL, vColumna)
  vGrid.MaxRows = vGrid.MaxRows - 1
      
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbCargaPagos(vGrid As Object, vGridMaxCol As Integer, vEstado As String, vColumna As Integer)
Dim strSQL As String

On Error GoTo vError

  If txtCod_Acreedor.Text = Empty Then Exit Sub
  
  lblCantidadCasos.Caption = 0
  lblTotalMonto.Caption = Format(0, "Standard")
  
  strSQL = " select 0 as 'Check',CP.LINEA, ACR.DESCRIPCION, CP.OPERACION, case when CP.TIPO_PAGO = 'O' then 'Ordinario' " _
         & " when CP.TIPO_PAGO = 'E' then  'ExtraOrdinario' else '' end as Pago,CP.N_CUOTA, CP.CUOTA, CP.ABONO, " _
         & " CP.SALDO, BANC.DESCRIPCION, " _
         & " CP.TIPO_DESEMBOLSO , CP.INTERESES, CP.COMISION, CP.CARGOS, CP.AMORTIZACION, CP.Estado " _
         & " from CRD_APA_CONTROL_PAGOS CP " _
         & " left join TES_BANCOS Banc on CP.COD_BANCO = Banc.ID_BANCO " _
         & " left join CRD_APA_ACREEDORES ACR on CP.COD_ACREEDOR = ACR.COD_ACREEDOR " _
         & " where CP.COD_ACREEDOR = '" & txtCod_Acreedor & "' and CP.ESTADO = '" & vEstado & "' "
          
'  Call sbCargaGridCheck(vGrid, vGridMaxCol, strSQL, vColumna)
  vPaso = True
     Call sbCargaGrid(vGrid, vGridMaxCol, strSQL)
  vPaso = False
  
  vGrid.MaxRows = vGrid.MaxRows - 1
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub tlbBuscaTraslados_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbCargaAsignados(gridTraslados, 16, Mid(cboEstadosAsignacion.Text, 1, 1), 3)
  Case "Bancos"
    Call sbCargaBancosAsignados
End Select

End Sub


Private Sub tlbDetallePagos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa As Integer

      
vOperacion = lblNumeroOperacion.Caption
vLinea = lblNPago.Tag
  
Select Case Button.Key
   Case "Guardar"
      Call sbModificaDesglose(txtCod_Acreedor.Text, vOperacion, vLinea)
   
   Case "Reimprime"
      
      strSQL = "select remesa From CRD_APA_CONTROL_PAGOS where COD_ACREEDOR = '" & txtCod_Acreedor & "' and OPERACION = '" & vOperacion & "' and LINEA = " & vLinea & ""
      Call OpenRecordSet(rs, strSQL)
          vRemesa = rs.Fields(0)
      rs.Close
        
      Call sbGeneraBoletaDesglose(txtCod_Acreedor.Text, vRemesa)
      Call sbLimpiar
      
      fraDesglosePago.Top = 55000
                
   Case "Cancelar"
      Call sbLimpiar
      

   Case "Revisar"
      Call sbDesgloseRevisaAsiento(txtCod_Acreedor.Text, vOperacion, vLinea)
   
End Select
  

  
End Sub

Private Sub sbReActivarLinea()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

vOperacion = lblNumeroOperacion.Caption
vLinea = lblNPago.Tag

strSQL = "exec spCrdAPA_ControlPago_ReActivaDetalle '" & txtCod_Acreedor.Text & "','" & vOperacion & "'," & vLinea
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "La Reactivación fue exitosa!", vbInformation

Call sbOperacionDetalle(vOperacion, vLinea)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub tlbInfo_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case "Info"

        If fraDetalleTesoreria.Visible Then
            fraDetalleTesoreria.Visible = False
            fraDetalleTesoreria.Top = 4000
        Else
            fraDetalleTesoreria.Top = 840
            fraDetalleTesoreria.Visible = True
        End If

    Case "ReActivar"
        Call sbReActivarLinea
End Select

End Sub


Private Sub txtCargosPago_GotFocus()
On Error GoTo vError
  txtCargosPago.Text = CCur(txtCargosPago.Text)
vError:
End Sub

Private Sub txtCargosPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionPago.SetFocus
End Sub

Private Sub txtCod_Acreedor_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   ssTab.Tab = 0
   Call sbCargaAcreedor
   Call sbCargarOperaciones
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Cod_Acreedor"
    gBusquedas.Orden = "Cod_Acreedor"
    gBusquedas.Filtro = ""
    gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
    frmBusquedas.Show vbModal
    txtCod_Acreedor = gBusquedas.Resultado
    txtNombreAcreedor = gBusquedas.Resultado2

End If

Call sbCargarOperaciones
    
End Sub

Private Sub txtCod_Acreedor_LostFocus()
    Call sbCargaAcreedor
    Call sbCargarOperaciones
    Call sbCargaAsignados(gridAsignacion, 14, "P", 2)
    Call sbBancosPagosAsignados
End Sub

Private Sub sbCargaAsignadosBancos(vCodBanco As Integer)
Dim strSQL As String

On Error GoTo vError
  
  
  If txtCod_Acreedor.Text = Empty Then Exit Sub

  lblCantidadCasos.Caption = 0
  lblTotalMonto.Caption = Format(0, "Standard")

  strSQL = "Select 0 as 'Check',CP.LINEA, ACR.DESCRIPCION, CP.OPERACION, case when CP.TIPO_PAGO = 'O' then 'Ordinario' " _
         & "when CP.TIPO_PAGO = 'E' then  'ExtraOrdinario' else '' end as Pago, CP.N_CUOTA, CP.CUOTA, CP.ABONO,CP.SALDO, " _
         & "BANC.DESCRIPCION, CP.TIPO_DESEMBOLSO , CP.INTERESES, CP.COMISION, CP.CARGOS, CP.AMORTIZACION, " _
         & "CP.Estado from CRD_APA_CONTROL_PAGOS CP " _
         & "left join TES_BANCOS Banc on CP.COD_BANCO = Banc.ID_BANCO " _
         & "left join CRD_APA_ACREEDORES ACR on CP.COD_ACREEDOR = ACR.COD_ACREEDOR " _
         & "where Banc.ID_BANCO = '" & vCodBanco & "' and CP.ESTADO ='G'"
  vPaso = True
    Call sbCargaGrid(gridAsignacion, 15, strSQL)
  vPaso = False
'  Call sbCargaGridCheck(gridAsignacion, 15, strSQL, 2)
  gridAsignacion.MaxRows = gridAsignacion.MaxRows - 1


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

'Carga lswAsignados
Private Sub sbBancosPagosAsignados()
 Dim strSQL As String, rs As New ADODB.Recordset
 
 On Error GoTo error
    
  If txtCod_Acreedor.Text = Empty Then Exit Sub
  
  strSQL = "select CP.COD_BANCO,Banc.DESCRIPCION,COUNT(CP.OPERACION)as Operaciones,SUM(CP.ABONO) as Abono " _
         & "from dbo.CRD_APA_CONTROL_PAGOS CP " _
         & "left join TES_BANCOS Banc on CP.COD_BANCO = Banc.ID_BANCO " _
         & "where CP.ESTADO='G' and CP.COD_ACREEDOR='" & txtCod_Acreedor.Text & "'" _
         & "group by cp.COD_BANCO,Banc.DESCRIPCION"
  
  Call OpenRecordSet(rs, strSQL)
     
    lswAsignados.ListItems.Clear
    Call sbCrearEncabezado
    
    Do While Not rs.EOF
     Set itmX = lswAsignados.ListItems.Add(1, , rs!COD_BANCO)
         itmX.SubItems(1) = rs!Descripcion
         itmX.SubItems(2) = rs!Operaciones
         itmX.SubItems(3) = Format(rs!ABONO, "Standard")
         
     rs.MoveNext
    Loop

 rs.Close
  
  
 Exit Sub

error:
   MsgBox fxSys_Error_Handler(Err.Description)
  
End Sub

Private Sub sbGeneraRemesa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa, vCuota As Integer, i As Integer
Dim vTesoreria As Long, vFecha As Date
Dim vMontoTotal As Double, vBanco As Integer

On Error GoTo vError

If cboBancoAsig.ComboItems.Count <= 0 Then Exit Sub

vBanco = DeCodificaPrimaryKey(cboBancoAsig.SelectedItem.Key, 1, "(id)")

strSQL = "select  isnull(max(NSOLICITUD),0) + 1 as 'NSolicitud',dbo.MyGetdate() as 'FechaServer' from TES_TRANSACCIONES"
Call OpenRecordSet(rs, strSQL)
    vTesoreria = rs!NSolicitud
    vFecha = rs!FechaServer
rs.Close

vMontoTotal = 0

'Crea la remesa
strSQL = "select isnull(max(remesa),0) + 1 as Ultimo from CRD_APA_REMESASTESORERIA"
Call OpenRecordSet(rs, strSQL)
    vRemesa = rs!Ultimo
rs.Close
       
strSQL = "insert CRD_APA_REMESASTESORERIA(REMESA,REGISTRO_USUARIO,REGISTRO_FECHA,ESTADO,FECHA_INICIO,FECHA_CORTE,NOTAS)" _
       & " values(" & vRemesa & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd") & "','A','" & Format(vFecha, "yyyymmdd") _
       & "','" & Format(vFecha, "yyyymmdd") & "','Generado por Pago Multiple')"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Remesa de administración pagarés: " & vRemesa)


With gridTraslados

   For i = 1 To .MaxRows
    .Row = i
    .Col = 2
    
    If .Value = 1 Then
        
        .Col = 3
        vLinea = .Text
        
        .Col = 5
        vOperacion = .Text
        
        .Col = 7
        vCuota = CInt(.Text)
                          
        .Col = 9
        vMontoTotal = vMontoTotal + CDbl(.Text)
        
        strSQL = "Update CRD_APA_CONTROL_PAGOS set REMESA=" & vRemesa & ",FECHA_CORTE_REMESA='" & Format(vFecha, "yyyymmdd") & "', " _
               & "ESTADO = 'T' where OPERACION ='" & vOperacion & "' and COD_ACREEDOR = '" & txtCod_Acreedor & "'  and LINEA= " & vLinea & ""
        Call ConectionExecute(strSQL)
        
        
        strSQL = "Update CRD_APA_CONTROL_PAGOS set TESORERIA_FECHA = '" & Format(vFecha, "yyyymmdd") & "', " _
               & "TESORERIA_SOLICITUD = " & vTesoreria & ", TESORERIA_USUARIO = '" & glogon.Usuario & "' where  REMESA = " & vRemesa & " " _
               & "and COD_ACREEDOR = '" & txtCod_Acreedor & "' and OPERACION = '" & vOperacion & "' and N_CUOTA  = " & vCuota & "" _
               & "and TESORERIA_FECHA is null"
        Call ConectionExecute(strSQL)
        
               
    End If ' Value = 1
  
   Next i
            
   strSQL = " exec spCRDAPATesoreriaPagoMultiple " & vRemesa & ",'" & txtCod_Acreedor.Text & "', '" & vOperacion & "'," & vCuota & ",'" & glogon.Usuario & "', " _
          & "'" & Format(vFecha, "yyyymmdd hh:mm:ss") & "', " & vMontoTotal & "," & vBanco & ""
   Call ConectionExecute(strSQL)
        
End With
   


'Actualiza el Estado de la Remesa como cerrada
strSQL = "update CRD_APA_REMESASTESORERIA set ESTADO = 'C' where REMESA = " & vRemesa
Call ConectionExecute(strSQL)
         
Call Bitacora("Aplica", "Cierre de Remesa de administración pagarés: " & vRemesa)
   
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbModificaDesglose(vAcreedor As String, vOperacion As String, vLinea As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vFechaCorte As Date
Dim vRemesa As Integer, vBanco As String, vModifica As Boolean
Dim curTotal As Currency

On Error GoTo vError

If Len(Trim(txtDocumentoPago.Text)) = 0 Then
   MsgBox "Especifique un número de documento?", vbExclamation
   Exit Sub
End If

curTotal = CCur(txtInteresesPago.Text) + CCur(txtComisionPago) + CCur(txtCargosPago) + CCur(lblAmortizacionPago)
If curTotal <> CCur(txtAbono.Text) Then
   MsgBox "Los montos del detalle no concuerdan con el Abono?", vbExclamation
   Exit Sub
End If

If CCur(lblAmortizacionPago.Caption) < 0 Then
   MsgBox "La amortización resultante es negativa! Verifique...", vbExclamation
   Exit Sub
End If

If CCur(lblAmortizacionPago.Caption) > CCur(lblUltimoSaldo.Caption) Then
   MsgBox "La amortización es mayor al Saldo Actual! Verifique...", vbExclamation
   Exit Sub
End If

vModifica = False
  
strSQL = "Select isnull(DOCUMENTO,'Null') as 'Documento', dbo.MyGetdate() as 'FechaServer'" _
       & ", dbo.fxAPA_OperacionValidaFechaMov(COD_ACREEDOR,OPERACION,'" & Format(dtpFechaPago.Value, "yyyy/mm/dd") & "','" & txtDocumentoPago.Text & "') as 'Resultado'" _
       & " from CRD_APA_CONTROL_PAGOS where ESTADO = 'T'" _
       & " and COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'" _
       & " and LINEA = " & vLinea & ""
       
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!FechaServer
    
    If rs.Fields("DOCUMENTO") = Trim(txtDocumentoPago.Text) Then
      vModifica = True
      MsgBox "La Operación no puede ser modificada porque ya fue procesada!", vbExclamation
      Exit Sub
    End If
    
    Select Case rs!Resultado
       Case 0 'No hay probleas
       Case 1 'Ya Existen movimientos
            vModifica = True
            MsgBox "La Operación tienen movimientos posteriores a esta fecha!", vbExclamation
            Exit Sub
       
       Case 2 'Ya Existe el Documento
            vModifica = True
            MsgBox "Ya existe registrado el documento!", vbExclamation
            Exit Sub
       Case Else
            vModifica = True
            MsgBox "La Operación no puede ser modificada verifique sus datos o consulte a su administrador!", vbExclamation
            Exit Sub
    End Select
    
rs.Close


vBanco = DeCodificaPrimaryKey(cboBancoAsig.SelectedItem.Key, 1, "(id)")

strSQL = "Select isnull(max(fecha_Corte), dbo.MyGetdate()) from CRD_APA_GARANTIAS_CORTES " _
      & "where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'"
      
Call OpenRecordSet(rs, strSQL)
    vFechaCorte = rs.Fields(0)
rs.Close

'Actualiza los datos en control e pagos
strSQL = "update CRD_APA_CONTROL_PAGOS set N_CUOTA = " & lblNPago & ", DOCUMENTO ='" & txtDocumentoPago.Text & "', SALDO = " & CCur(lblSaldoPago.Caption) _
       & ",INTERESES = " & CCur(txtInteresesPago) & ",COMISION = " & CCur(txtComisionPago) & ",CARGOS = " & CCur(txtCargosPago) _
       & ",AMORTIZACION = " & CCur(lblAmortizacionPago) & ",FECHA_DESGLOSE = '" & Format(dtpFechaPago.Value, "yyyymmdd") & "',TASA=" & CCur(lblTasaPago.Caption) _
       & " where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "' and LINEA = " & vLinea & ""
Call ConectionExecute(strSQL)

strSQL = "update CRD_APA_GARANTIAS_CORTES set SALDO_OPERACION = " & CCur(lblSaldoPago.Caption) _
       & ",SALDO_RESPONSABILIDAD = dbo.fxCRDAPASaldoResponsabilidad ('" & vAcreedor & "','" & vOperacion & "')" _
       & " where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "' and FECHA_CORTE ='" & Format(vFechaCorte, "yyyymmdd") & "'"
Call ConectionExecute(strSQL)


'Actualiza el saldo de la operacion Aplancada
strSQL = "update CRD_APA_OPERACIONES set SALDO = isnull(Saldo,0) - " & CCur(lblAmortizacionPago.Caption) _
       & " where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'"
Call ConectionExecute(strSQL)


strSQL = "select remesa From CRD_APA_CONTROL_PAGOS where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'and LINEA = " & vLinea & ""
Call OpenRecordSet(rs, strSQL)
    vRemesa = rs.Fields(0)
rs.Close
           
'Genera los asientos y el desglose del pago
strSQL = "exec spCRDAPATesoreriaDesglosePago " & vRemesa & ", '" & vAcreedor & "','" & vOperacion & "', " _
       & "" & lblNPago & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd hh:mm:ss") & "'," & vLinea & ",'" & vBanco & "'"
Call ConectionExecute(strSQL)

Call sbGeneraBoletaDesglose(vAcreedor, vRemesa)

Call Bitacora("Modifica Desglose", "Remesa de administración garantías: " & vRemesa)

Call sbLimpiar
fraDesglosePago.Top = 55000

       
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

               
End Sub


Private Sub sbDesgloseRevisaAsiento(vAcreedor As String, vOperacion As String, vLinea As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vFechaCorte As Date
Dim vRemesa As Integer, vBanco As String

On Error GoTo vError

  
strSQL = "Select DOCUMENTO, dbo.MyGetdate() as 'FechaServer' from CRD_APA_CONTROL_PAGOS" _
       & " where ESTADO = 'T' and COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'" _
       & "and LINEA = " & vLinea & ""
       
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!FechaServer
    
    If IsNull(rs!Documento) Then
      rs.Close
      MsgBox "La Operación ha sido detallado todavía!", vbExclamation
      Exit Sub
    End If
rs.Close


vBanco = DeCodificaPrimaryKey(cboBancoAsig.SelectedItem.Key, 1, "(id)")

strSQL = "Select max(fecha_Corte) " _
       & "  from CRD_APA_GARANTIAS_CORTES " _
       & " where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'"
      
Call OpenRecordSet(rs, strSQL)
    vFechaCorte = rs.Fields(0)
rs.Close

           
strSQL = "select remesa From CRD_APA_CONTROL_PAGOS where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'and LINEA = " & vLinea & ""
Call OpenRecordSet(rs, strSQL)
    vRemesa = rs.Fields(0)
rs.Close
           
'Revisa Asiento de Detalle del pago
strSQL = "exec spCRDAPADesglosePagoAjusteAsiento " & vRemesa & ", '" & vAcreedor & "','" & vOperacion & "', " _
       & "" & lblNPago & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd hh:mm:ss") & "'," & vLinea & ",'" & vBanco & "'"
Call ConectionExecute(strSQL)

Call sbGeneraBoletaDesglose(vAcreedor, vRemesa)
Call sbLimpiar

fraDesglosePago.Top = 55000

       
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
               
End Sub



'Private Sub sbEliminaDesglose(vAcreedor As String, vOperacion As String, vLinea As Integer)
'Dim strSQL As String
'Dim rs As New ADODB.Recordset
'Dim vFecha As Date
'Dim vFechaCorte As Date
'Dim vRemesa As Integer
'Dim vBanco As String
'Dim vModifica As Boolean
'
'On Error GoTo error
'
'  vFecha = fxFechaServidor
'
'  strSQL = "select LINEA,REMESA,COD_BANCO From CRD_APA_CONTROL_PAGOS where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'and LINEA = " & vLinea & ""
'  Call OpenRecordSet(rs, strSQL)
'
'  vRemesa = rs!REMESA
'  vBanco = rs!COD_BANCO
'
'  rs.Close
'
'  strSQL = " exec spCRDAPATesoreriaEliminaDesglosePago " & vRemesa & ", '" & vAcreedor & "','" & vOperacion & "', " _
'          & "" & lblNPago & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd hh:mm:ss") & "'," & vLinea & ",'" & vBanco & "'"
'
'  Call ConectionExecute(strSQL)
'
'  vSaldoAnterior = CDbl(lblSaldoPago) + CDbl(lblAmortizacionPago)
'
'  'Actualiza el salgo de la operacion
'  strSQL = "update CRD_APA_OPERACIONES set SALDO=" & Ccur(vSaldoAnterior) & "" _
'         & "where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'"
'
'  Call ConectionExecute(strSQL)
'
'  strSQL = "update CRD_APA_CONTROL_PAGOS set SALDO=" & vSaldoAnterior & ",INTERESES=0,COMISION=0,CARGOS=0,AMORTIZACION=0 " _
'         & "where COD_ACREEDOR = '" & vAcreedor & "' and OPERACION = '" & vOperacion & "'"
'
'  Call ConectionExecute(strSQL)
'
'  Call sbGeneraBoletaDesglose(vAcreedor, vRemesa)
'
'  Call Bitacora("Elimina Desglose", "Remesa de administración pagarés: " & vRemesa)
'
'Exit Sub
'
'error:
'  MsgBox fxSys_Error_Handler(Err.Description)
'
'End Sub

Private Sub sbLimpiar()
   
  fraDesglosePago.Top = 55000
  fraDetalleTesoreria.Visible = False
  fraDetalleTesoreria.Top = 4000
   
  vSaldoAnterior = 0
  lblSaldoAnterior.Caption = 0
  lblUltimoSaldo = Empty
  lblAmortizacionPago = Empty
  txtAbono = Empty
  txtInteresesPago = Empty
  txtCargosPago = Empty
  txtComisionPago = Empty
  lblNumeroOperacion.Caption = Empty
  vLinea = Empty
  vOperacion = Empty

End Sub

Private Sub sbCalcularSaldos()

On Error GoTo vError

    If lblUltimoSaldo = Empty Then lblUltimoSaldo = 0
    If lblAmortizacionPago = Empty Then lblAmortizacionPago = 0
    If txtAbono = Empty Then txtAbono = 0
    If txtInteresesPago = Empty Then txtInteresesPago = 0
    If txtCargosPago = Empty Then txtCargosPago = 0
    If txtComisionPago = Empty Then txtComisionPago = 0
    
    lblAmortizacionPago = CStr(CDbl(txtAbono.Text) - (CCur(txtInteresesPago.Text) + CCur(txtCargosPago.Text) + CCur(txtComisionPago.Text)))
    lblAmortizacionPago = Format(lblAmortizacionPago.Caption, "Standard")
    
    If CCur(lblSaldoAnterior.Caption) <> 0 Then
        lblTasaPago.Caption = CCur(txtInteresesPago.Text) / CCur(lblSaldoAnterior.Caption) * 12 * 100
        lblTasaPago.Caption = Format(lblTasaPago.Caption, "Standard")
    Else
        lblTasaPago = Format(0, "Standard")
    End If
    
    lblSaldoPago.Caption = CCur(lblSaldoAnterior.Caption) - CCur(lblAmortizacionPago.Caption)
    lblSaldoPago.Caption = Format(lblSaldoPago.Caption, "Standard")
    
    Exit Sub
    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub txtAbono_Change()
  vCambio = True
End Sub

Private Sub txtAbono_LostFocus()
  If txtAbono.Text <> Empty Then
    If Not IsNumeric(txtAbono) Then
        MsgBox "El campo solo permite valores numéricos"
        txtAbono.SetFocus
    End If
    txtAbono = Format(txtAbono, "Standard")
  End If
  Call sbCalcularSaldos
End Sub

Private Sub txtCargosPago_Change()
  vCambio = True
End Sub

Private Sub txtCargosPago_LostFocus()
  If txtCargosPago.Text <> Empty Then
    If Not IsNumeric(txtCargosPago) Then
        MsgBox "El campo solo permite valores numéricos"
        txtCargosPago.SetFocus
    End If
    txtCargosPago = Format(txtCargosPago, "Standard")
  End If
  Call sbCalcularSaldos
End Sub

Private Sub txtComisionPago_Change()
  vCambio = True
End Sub

Private Sub txtComisionPago_GotFocus()
On Error GoTo vError
  txtComisionPago.Text = CCur(txtComisionPago.Text)
vError:
End Sub

Private Sub txtComisionPago_LostFocus()
  If txtComisionPago.Text <> Empty Then
    If Not IsNumeric(txtComisionPago) Then
        MsgBox "El campo solo permite valores numéricos"
        txtComisionPago.SetFocus
    End If
    txtComisionPago = Format(txtCargosPago, "Standard")
  End If
  Call sbCalcularSaldos
End Sub

Private Sub txtInteresesPago_Change()
  vCambio = True
End Sub

Private Sub txtInteresesPago_GotFocus()
On Error GoTo vError
  txtInteresesPago.Text = CCur(txtInteresesPago.Text)
vError:
End Sub

Private Sub txtInteresesPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargosPago.SetFocus
End Sub

Private Sub txtInteresesPago_LostFocus()
    If txtInteresesPago.Text <> Empty Then
        If Not IsNumeric(txtInteresesPago) Then
            MsgBox "El campo solo permite valores numéricos"
            txtInteresesPago.SetFocus
        End If
        txtInteresesPago = Format(txtInteresesPago, "Standard")
    End If
    Call sbCalcularSaldos
End Sub

Private Sub sbCargaSaldoAnterior(vAcreedor As String, vOperacion As String, vLinea As Integer)
Dim strSQL As String
Dim rs As New ADODB.Recordset

On Error GoTo error

   strSQL = "select SALDO from CRD_APA_CONTROL_PAGOS " _
          & "where COD_ACREEDOR= '" & vAcreedor & "' and OPERACION= '" & vOperacion & "' " _
          & "and max(LINEA) "
       
   Call OpenRecordSet(rs, strSQL)
   
   lblUltimoSaldo = rs!Saldo
    
   rs.Close
   
Exit Sub

error:
  MsgBox fxSys_Error_Handler(Err.Description)
  
End Sub

Private Sub txtNombreAcreedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbCargaAcreedor
   Call sbCargarOperaciones
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Cod_Acreedor"
    gBusquedas.Orden = "Cod_Acreedor"
    gBusquedas.Filtro = ""
    gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
    frmBusquedas.Show vbModal
    txtCod_Acreedor = gBusquedas.Resultado
    txtNombreAcreedor = gBusquedas.Resultado2
End If

Call sbCargarOperaciones
End Sub

Private Sub sbReversaSolicitudPago(vOperacion As String, vCod_Acreedor As String, vLinea As Long)
Dim strSQL As String


strSQL = "spReversaSolicitudPago  '" & vOperacion & "','" & vCod_Acreedor & "'," & vLinea
Call ConectionExecute(strSQL)
    
End Sub

Private Sub sbGeneraBoletaDesglose(vAcreedor As String, vRemesa As Integer)
Me.MousePointer = vbHourglass
Dim vTransaccion As String

On Error GoTo vError

'If lblBoleta.Tag = "" Then Exit Sub

vTransaccion = vAcreedor & "_" & CStr(vRemesa)

With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Pasivos"
    
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_DesglosePago.rpt")
  .Formulas(1) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  
  .SelectionFormula = "{CRD_APA_CONTROL_PAGOS.OPERACION}='" & Trim(lblNumeroOperacion.Caption) _
                    & "' and {CRD_APA_CONTROL_PAGOS.N_CUOTA} = " & lblNPago.Caption _
                    & " and {CRD_APA_CONTROL_PAGOS.COD_ACREEDOR} = '" & vAcreedor & "'"
  
  .SubreportToChange = "sbAsiento"
  .StoredProcParam(0) = "APA"
  .StoredProcParam(1) = vTransaccion
  .StoredProcParam(2) = 1

  
  .Action = 1
 ' .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

