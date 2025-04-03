VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCxC_FacturasMonitoreo 
   Caption         =   "Monitoreo de Facturas"
   ClientHeight    =   8910
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   450
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox gbDetalle 
      Height          =   2292
      Left            =   3480
      TabIndex        =   1
      Top             =   6120
      Width           =   8892
      _Version        =   1310722
      _ExtentX        =   15684
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Factura No."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.TabControl tcDetalle 
         Height          =   2052
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8652
         _Version        =   1310722
         _ExtentX        =   15261
         _ExtentY        =   3619
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
         Item(0).Caption =   "Gestión"
         Item(0).ControlCount=   15
         Item(0).Control(0)=   "Label2(0)"
         Item(0).Control(1)=   "Label2(1)"
         Item(0).Control(2)=   "Label2(2)"
         Item(0).Control(3)=   "Label2(3)"
         Item(0).Control(4)=   "Label2(4)"
         Item(0).Control(5)=   "Label2(5)"
         Item(0).Control(6)=   "feG_Operacion"
         Item(0).Control(7)=   "feG_Cliente"
         Item(0).Control(8)=   "feG_Factura"
         Item(0).Control(9)=   "feG_Estado"
         Item(0).Control(10)=   "feG_Monto"
         Item(0).Control(11)=   "feG_Pendiente"
         Item(0).Control(12)=   "btnTramita"
         Item(0).Control(13)=   "btnCancela"
         Item(0).Control(14)=   "btnSustituye"
         Item(1).Caption =   "Historial"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "lswHistorial"
         Item(1).Control(1)=   "btnExport(0)"
         Item(2).Caption =   "Desembolsos"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "lswDesembolsos"
         Item(2).Control(1)=   "btnExport(1)"
         Item(3).Caption =   "Cancelación"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "lswCancelacion"
         Item(3).Control(1)=   "btnExport(2)"
         Begin XtremeSuiteControls.ListView lswHistorial 
            Height          =   1575
            Left            =   -69760
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   8415
            _Version        =   1310722
            _ExtentX        =   14838
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
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswDesembolsos 
            Height          =   1575
            Left            =   -69760
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   8415
            _Version        =   1310722
            _ExtentX        =   14838
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
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswCancelacion 
            Height          =   1575
            Left            =   -69760
            TabIndex        =   43
            Top             =   360
            Visible         =   0   'False
            Width           =   8415
            _Version        =   1310722
            _ExtentX        =   14838
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
            View            =   3
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnTramita 
            Height          =   504
            Left            =   3960
            TabIndex        =   38
            Top             =   1320
            Width           =   1452
            _Version        =   1310722
            _ExtentX        =   2561
            _ExtentY        =   889
            _StockProps     =   79
            Caption         =   "Tramita"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":0000
         End
         Begin XtremeSuiteControls.FlatEdit feG_Cliente 
            Height          =   312
            Left            =   3960
            TabIndex        =   33
            Top             =   480
            Width           =   4572
            _Version        =   1310722
            _ExtentX        =   8064
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
         Begin XtremeSuiteControls.FlatEdit feG_Estado 
            Height          =   312
            Left            =   3960
            TabIndex        =   35
            Top             =   840
            Width           =   4572
            _Version        =   1310722
            _ExtentX        =   8064
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCancela 
            Height          =   504
            Left            =   5400
            TabIndex        =   39
            Top             =   1320
            Width           =   1572
            _Version        =   1310722
            _ExtentX        =   2773
            _ExtentY        =   889
            _StockProps     =   79
            Caption         =   "Cancela"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":07D8
         End
         Begin XtremeSuiteControls.PushButton btnSustituye 
            Height          =   504
            Left            =   6960
            TabIndex        =   40
            Top             =   1320
            Width           =   1572
            _Version        =   1310722
            _ExtentX        =   2773
            _ExtentY        =   889
            _StockProps     =   79
            Caption         =   "Sustituye"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":0E9F
         End
         Begin XtremeSuiteControls.FlatEdit feG_Monto 
            Height          =   312
            Left            =   1200
            TabIndex        =   36
            Top             =   1200
            Width           =   1932
            _Version        =   1310722
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feG_Pendiente 
            Height          =   312
            Left            =   1200
            TabIndex        =   37
            Top             =   1560
            Width           =   1932
            _Version        =   1310722
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feG_Operacion 
            Height          =   312
            Left            =   1200
            TabIndex        =   32
            Top             =   480
            Width           =   1932
            _Version        =   1310722
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feG_Factura 
            Height          =   312
            Left            =   1200
            TabIndex        =   34
            Top             =   840
            Width           =   1932
            _Version        =   1310722
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   255
            Index           =   0
            Left            =   -70000
            TabIndex        =   50
            ToolTipText     =   "Exportar a Excel"
            Top             =   360
            Visible         =   0   'False
            Width           =   255
            _Version        =   1310722
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":182C
         End
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   255
            Index           =   1
            Left            =   -70000
            TabIndex        =   51
            ToolTipText     =   "Exportar a Excel"
            Top             =   360
            Visible         =   0   'False
            Width           =   255
            _Version        =   1310722
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":20FD
         End
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   255
            Index           =   2
            Left            =   -70000
            TabIndex        =   52
            ToolTipText     =   "Exportar a Excel"
            Top             =   360
            Visible         =   0   'False
            Width           =   255
            _Version        =   1310722
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "frmCxC_FacturasMonitoreo.frx":29CE
         End
         Begin VB.Label Label2 
            Caption         =   "Pendiente"
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
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   972
         End
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   30
            Top             =   1200
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente"
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
            Left            =   3240
            TabIndex        =   29
            Top             =   480
            Width           =   972
         End
         Begin VB.Label Label2 
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
            Index           =   2
            Left            =   3240
            TabIndex        =   28
            Top             =   840
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Factura"
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
            TabIndex        =   27
            Top             =   840
            Width           =   972
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
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   972
         End
      End
   End
   Begin XtremeSuiteControls.FlatEdit feFactura 
      Height          =   312
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.GroupBox gbContratos 
      Height          =   2652
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   3132
      _Version        =   1310722
      _ExtentX        =   5524
      _ExtentY        =   4678
      _StockProps     =   79
      Caption         =   "Contratos:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswContratos 
         Height          =   2172
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   3831
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkContratos 
         Height          =   252
         Left            =   2160
         TabIndex        =   44
         Top             =   0
         Width           =   972
         _Version        =   1310722
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
   End
   Begin XtremeSuiteControls.GroupBox gbConceptos 
      Height          =   3372
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3132
      _Version        =   1310722
      _ExtentX        =   5524
      _ExtentY        =   5948
      _StockProps     =   79
      Caption         =   "Conceptos:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkConceptos 
         Height          =   252
         Left            =   2160
         TabIndex        =   45
         Top             =   0
         Width           =   972
         _Version        =   1310722
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
      Begin XtremeSuiteControls.ListView lswConceptos 
         Height          =   2772
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   4890
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2412
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   3132
      _Version        =   1310722
      _ExtentX        =   5524
      _ExtentY        =   4254
      _StockProps     =   79
      Caption         =   "Filtros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkAdelantadas 
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
         _Version        =   1310722
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Facturas Adelantadas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstados 
         Height          =   312
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2778
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboFecha 
         Height          =   312
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2778
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   1080
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   556
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
         Left            =   1560
         TabIndex        =   25
         Top             =   1440
         Width           =   1572
         _Version        =   1310722
         _ExtentX        =   2773
         _ExtentY        =   556
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Base"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1212
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
         Height          =   312
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1212
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   3480
      TabIndex        =   14
      Top             =   1200
      Width           =   10692
      _Version        =   524288
      _ExtentX        =   18860
      _ExtentY        =   8488
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
      MaxCols         =   16
      SpreadDesigner  =   "frmCxC_FacturasMonitoreo.frx":329F
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit feCliente 
      Height          =   312
      Left            =   5160
      TabIndex        =   15
      Top             =   360
      Width           =   4092
      _Version        =   1310722
      _ExtentX        =   7218
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
   Begin XtremeSuiteControls.FlatEdit fePagador 
      Height          =   312
      Left            =   9240
      TabIndex        =   16
      Top             =   360
      Width           =   4092
      _Version        =   1310722
      _ExtentX        =   7218
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_FacturasMonitoreo.frx":3CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_FacturasMonitoreo.frx":4713
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_FacturasMonitoreo.frx":4ECF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit feClienteId 
      Height          =   312
      Left            =   6480
      TabIndex        =   21
      Top             =   720
      Width           =   2772
      _Version        =   1310722
      _ExtentX        =   4890
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
   Begin XtremeSuiteControls.FlatEdit fePagadorId 
      Height          =   312
      Left            =   10560
      TabIndex        =   23
      Top             =   720
      Width           =   2772
      _Version        =   1310722
      _ExtentX        =   4890
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   13440
      TabIndex        =   48
      Top             =   360
      Width           =   1215
      _Version        =   1310722
      _ExtentX        =   2138
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCxC_FacturasMonitoreo.frx":56D4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   14640
      TabIndex        =   49
      Top             =   360
      Width           =   1575
      _Version        =   1310722
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCxC_FacturasMonitoreo.frx":60F2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagador Id...:"
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
      Left            =   9360
      TabIndex        =   22
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Id...:"
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
      Left            =   5160
      TabIndex        =   20
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura...:"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente...:"
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
      Left            =   5160
      TabIndex        =   18
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pagador...:"
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
      Index           =   8
      Left            =   9360
      TabIndex        =   17
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmCxC_FacturasMonitoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOperacion As Long, mFactura As String


Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnCancela_Click()

On Error GoTo vError

GLOBALES.gTag = feG_Operacion.Text
GLOBALES.gTag2 = feG_Factura.Text
GLOBALES.gTag3 = feG_Cliente.Tag

Call sbFormsCall("frmCxC_Facturas_Cancela", vbModal, , , False, Me)

Call sbFactura_Detalle

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExport_Click(Index As Integer)
Select Case Index
    Case 0 'Historial
        Call Excel_Exportar_Lsw(lswHistorial)
    Case 1 'Desembolsos
        Call Excel_Exportar_Lsw(lswDesembolsos)
    Case 2 'Cancelacion
        Call Excel_Exportar_Lsw(lswCancelacion)
End Select
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 16
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "No Operación"
    vHeaders.Headers(3) = "No. Factura"
    vHeaders.Headers(4) = "Monto"
    vHeaders.Headers(5) = "Estado"
    vHeaders.Headers(6) = "Fec.Emision"
    vHeaders.Headers(7) = "Divisa"
    vHeaders.Headers(8) = "Tipo Cambio"
    vHeaders.Headers(9) = "(%) Adelanto"
    vHeaders.Headers(10) = "Monto Adelanto"
    vHeaders.Headers(11) = "Pendiente Girar"
    vHeaders.Headers(12) = "Cliente"
    vHeaders.Headers(13) = "Pagador"
    vHeaders.Headers(14) = "Fecha Cancela"
    vHeaders.Headers(15) = "I Remesa Desembolso"
    vHeaders.Headers(16) = "II Remesa Desembolso"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_CxC_Facturas_Monitoreo")
End Sub

Private Sub btnSustituye_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pFacturaEstado As String

On Error GoTo vError


pFacturaEstado = fxCxC_FacturaEstadoDefault("Sustitución")

strSQL = "exec spCxC_Operacion_Factura_Estado " & feG_Operacion.Text & ",'" & feG_Factura.Text & "','" & pFacturaEstado _
       & "','" & glogon.Usuario & "',1,0"
Call ConectionExecute(strSQL)

Call sbFactura_Detalle

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnTramita_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pFacturaEstado As String

On Error GoTo vError


pFacturaEstado = fxCxC_FacturaEstadoDefault("Confirmación")

strSQL = "exec spCxC_Operacion_Factura_Estado " & feG_Operacion.Text & ",'" & feG_Factura.Text & "','" & pFacturaEstado _
       & "','" & glogon.Usuario & "',1,0"
Call ConectionExecute(strSQL)

Call sbFactura_Detalle

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkConceptos_Click()
Dim i As Integer

If lswConceptos.ListItems.Count = 0 Then
    Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

    strSQL = "select COD_CONCEPTO as 'IdX', descripcion as 'ItmX' from CXC_CONCEPTOS where PROCESO_DESCUENTO = 1 order by descripcion"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswConceptos.ListItems.Add(, , rs!itmX)
         itmX.Tag = rs!idX
         itmX.Checked = chkConceptos.Value
     rs.MoveNext
    Loop
    rs.Close
    
    lswConceptos.BackColor = vbWhite


End If

For i = 1 To lswConceptos.ListItems.Count
  lswConceptos.ListItems.Item(i).Checked = chkConceptos.Value
Next i
End Sub

Private Sub chkContratos_Click()
Dim i As Integer

If lswContratos.ListItems.Count = 0 Then
    Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

    strSQL = "select COD_CONTRATO as 'IdX', descripcion as 'ItmX' from CXC_CONTRATOS where ACTIVO = 1 order by descripcion"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswContratos.ListItems.Add(, , rs!itmX)
         itmX.Tag = rs!idX
         itmX.Checked = chkContratos.Value
     rs.MoveNext
    Loop
    rs.Close
    
    lswContratos.BackColor = vbWhite


End If

For i = 1 To lswContratos.ListItems.Count
  lswContratos.ListItems.Item(i).Checked = chkContratos.Value
Next i
End Sub

Private Sub feCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      feClienteId.Text = Trim(gBusquedas.Resultado)
      feCliente.Text = Trim(gBusquedas.Resultado2)
   End If
End If

End Sub



Private Sub feClienteId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If


If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      feClienteId.Text = Trim(gBusquedas.Resultado)
      feCliente.Text = Trim(gBusquedas.Resultado2)
   End If
End If
End Sub



Private Sub feFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If
End Sub

Private Sub fePagador_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   gBusquedas.Filtro = " and cedula in(select cedula from cxc_contratos_pagadores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      fePagadorId.Text = Trim(gBusquedas.Resultado)
      fePagador.Text = Trim(gBusquedas.Resultado2)
   End If
End If

End Sub

Private Sub fePagadorId_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   gBusquedas.Filtro = " and cedula in(select cedula from cxc_contratos_pagadores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      fePagadorId.Text = Trim(gBusquedas.Resultado)
      fePagador.Text = Trim(gBusquedas.Resultado2)
   End If
End If

End Sub

Private Sub Form_Activate()
 vModulo = 31
End Sub

Private Sub Form_Load()
vModulo = 31

lswConceptos.ColumnHeaders.Add , , "", 3150
lswConceptos.HideColumnHeaders = True

lswContratos.ColumnHeaders.Add , , "", 3150
lswContratos.HideColumnHeaders = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
Dim pHeight As Long, pWidth As Long

On Error Resume Next

pWidth = Me.Width - 150
pHeight = Me.Height - (StatusBarX.Height + 550)

gbConceptos.Height = (pHeight - (gbFiltros.Height + 250)) / 2
gbContratos.Height = gbConceptos.Height

gbContratos.top = gbConceptos.top + gbConceptos.Height + 150
gbFiltros.top = gbContratos.top + gbContratos.Height + 150

lswConceptos.Height = gbConceptos.Height - 450
lswContratos.Height = gbContratos.Height - 450

vGrid.Height = pHeight - (vGrid.top + gbDetalle.Height + 200)
vGrid.Width = pWidth - (vGrid.Left + 250)

gbDetalle.top = gbFiltros.top
gbDetalle.Left = vGrid.Left
gbDetalle.Width = vGrid.Width

tcDetalle.Width = gbDetalle.Width - 150

lswHistorial.Width = tcDetalle.Width - 250
lswCancelacion.Width = lswHistorial.Width
lswDesembolsos.Width = lswHistorial.Width

End Sub

Private Sub lswDesembolsos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Dim frm As Form

If Not IsNumeric(Item.SubItems(2)) Then Exit Sub
If CCur(Item.SubItems(2)) <= 0 Then Exit Sub
 

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(Item.SubItems(2))
     Exit For
   End If
 Next frm

End Sub


Private Sub tcDetalle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call sbFactura_Detalle
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

TimerX.Interval = 0
TimerX.Enabled = False

vGrid.AppearanceStyle = fxGridStyle

strSQL = "select FACTURA_ESTADO as 'IdX', DESCRIPCION as 'ItmX' from CXC_FACTURAS_ESTADOS"
Call sbCbo_Llena_New(cboEstados, strSQL, True, True)


cboFecha.Clear
cboFecha.AddItem "Registro"
cboFecha.AddItem "Emisión"
cboFecha.AddItem "Pago"
cboFecha.AddItem "Libera"
cboFecha.AddItem "Cancela"
cboFecha.AddItem "Activación"
cboFecha.AddItem "Desembolso 1"
cboFecha.AddItem "Desembolso 2"
cboFecha.AddItem "[TODAS]"

cboFecha.Text = "Registro"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

Call sbBuscar

'Call sbLimpia

End Sub




Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

strSQL = "select '',Operacion,cod_Factura,Monto, Factura_Estado_Desc, Fecha_Emision, Cod_Divisa, Tipo_Cambio " _
       & ", Adelanto_Porc, Adelanto_Monto,Pendiente" _
       & ", rtrim(Cedula) + ' - ' + Cliente_Nombre, rtrim(Cedula_Pagador) + ' - ' + Pagador_Nombre, Cancela_Fecha" _
       & ", '[' + convert(varchar(30),Pago_Principal_Remesa) + '] ' + convert(varchar(30), Pago_Principal_Fecha) as 'RemesaI'" _
       & ", '[' + convert(varchar(30),Pago_Segundo_Remesa) + '] ' + convert(varchar(30), Pago_Segundo_Fecha) as 'RemesaII'" _
       & "  From vCxC_Facturas_Control" _
       & " where cod_Factura like '%" & feFactura.Text & "%'" _

If Len(feClienteId.Text) > 0 Then
   strSQL = strSQL & " and cedula like '%" & feClienteId.Text & "%'"
End If

If Len(feCliente.Text) > 0 Then
   strSQL = strSQL & " and Cliente_Nombre like '%" & feCliente.Text & "%'"
End If

If Len(fePagadorId.Text) > 0 Then
   strSQL = strSQL & " and Cedula_Pagador like '%" & fePagadorId.Text & "%'"
End If


If Len(fePagador.Text) > 0 Then
   strSQL = strSQL & " and Pagador_Nombre like '%" & fePagador.Text & "%'"
End If


'Lista de Contratos
If lswContratos.ListItems.Count > 0 Then
    vCadena = " and Cod_Contrato in('"
    For i = 1 To lswContratos.ListItems.Count
      If lswContratos.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswContratos.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If

iCantidad = 0
'Lista de Conceptos
If lswConceptos.ListItems.Count > 0 Then
    vCadena = " and Cod_Concepto in('"
    For i = 1 To lswConceptos.ListItems.Count
      If lswConceptos.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswConceptos.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


Select Case cboFecha.Text
  Case "Registro"
    strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Emisión"
    strSQL = strSQL & " and fecha_Emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Pago"
    strSQL = strSQL & " and fecha_Pago between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Libera"
    strSQL = strSQL & " and Liberado_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Cancela"
    strSQL = strSQL & " and Cancela_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Activación"
    strSQL = strSQL & " and Activa_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Desembolso 1"
    strSQL = strSQL & " and Pago_Principal_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Desembolso 2"
    strSQL = strSQL & " and Pago_Secundario_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

End Select

If cboEstados.Text <> "TODOS" Then
     strSQL = strSQL & " and factura_estado = '" & cboEstados.ItemData(cboEstados.ListIndex) & "'"
End If

If chkAdelantadas.Value = xtpChecked Then
    strSQL = strSQL & " and ADELANTO_INDICA = " & chkAdelantadas.Value
End If

Call sbCargaGridLocal(vGrid, 16, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i

'    If rs.Fields(i - 1).Type = 135 Then
'        If Year(rs.Fields(i - 1).Value) > 1900 Then
'           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
'        End If
'    Else
'        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
'    End If
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Monto
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

vGrid.Row = Row
vGrid.col = 2
feG_Operacion.Text = vGrid.Text

vGrid.col = 3
feG_Factura.Text = vGrid.Text

vGrid.col = 12
gbDetalle.Caption = "Factura No. " & feG_Factura.Text & " [ " & vGrid.Text & " ]"

Dim i As Integer

i = tcDetalle.Selected.Index

'Carga Datos Base
tcDetalle.Item(0).Selected = True
Call sbFactura_Detalle

'Regresa al Tab Actual
tcDetalle.Item(i).Selected = True
Call sbFactura_Detalle

End Sub


Private Sub sbFactura_Detalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If Not IsNumeric(feG_Operacion.Text) Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "exec spCxC_Operacion_Factura_Detalle " & feG_Operacion.Text & ",'" & feG_Factura.Text & "'"

Select Case tcDetalle.SelectedItem
  Case 0 'Gestion
    strSQL = strSQL & ",'G'"
    Call OpenRecordSet(rs, strSQL)
    feG_Cliente.Text = rs!Cliente_Nombre
    feG_Cliente.Tag = rs!Cedula
    feG_Monto.Text = Format(rs!Monto, "Standard")
    feG_Pendiente.Text = Format(rs!Pendiente, "Standard")
    feG_Estado.Tag = rs!FACTURA_ESTADO
    feG_Estado.Text = rs!Factura_Estado_Desc
    
    btnTramita.Enabled = False
    btnCancela.Enabled = False
    btnSustituye.Enabled = False
    
   
    If rs!Factura_Proceso = "Registro" And btnTramita.Tag <> "0" Then
        btnTramita.Enabled = True
    End If
    
    If rs!Factura_Proceso = "Confirmación" Or rs!Factura_Proceso = "Registro" And btnSustituye.Tag <> "0" Then
        btnSustituye.Enabled = True
    End If
    
    If rs!Factura_Proceso = "Confirmación" And btnCancela.Tag <> "0" Then
        btnCancela.Enabled = True
    End If
    
    
  Case 1 'Historial
    strSQL = strSQL & ",'H'"
    
    With lswHistorial
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Fecha", 1800
            .ColumnHeaders.Add , , "Usuario", 2200
            .ColumnHeaders.Add , , "Estado", 2200
            .ColumnHeaders.Add , , "Notas", 4200
    
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!Registro_Fecha)
                itmX.SubItems(1) = rs!Registro_Usuario
                itmX.SubItems(2) = rs!Descripcion
                itmX.SubItems(3) = rs!notas
            rs.MoveNext
        Loop
    End With
    
    
  Case 2 'Desembolsos
    strSQL = strSQL & ",'D'"
  With lswDesembolsos
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Factura", 1800
            .ColumnHeaders.Add , , "Fac.Monto", 2200, vbRightJustify
            .ColumnHeaders.Add , , "Tesoreria [Id]", 2200
            .ColumnHeaders.Add , , "Estado", 1200, vbCenter
            .ColumnHeaders.Add , , "Fecha", 1800, vbCenter
            .ColumnHeaders.Add , , "Monto", 2200, vbRightJustify
            .ColumnHeaders.Add , , "Tipo", 1200, vbCenter
            .ColumnHeaders.Add , , "Banco", 3200
            .ColumnHeaders.Add , , "Beneficiario", 3200
    
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!cod_Factura)
                itmX.SubItems(1) = Format(rs!MONTO_MOV, "Standard")
                itmX.SubItems(2) = rs!NSolicitud
                itmX.SubItems(3) = rs!Estado
                itmX.SubItems(4) = rs!Fecha_Emision & ""
                itmX.SubItems(5) = Format(rs!Monto, "Standard")
                itmX.SubItems(6) = rs!Tipo
                itmX.SubItems(7) = rs!Banco
                itmX.SubItems(8) = rs!Beneficiario
                
            rs.MoveNext
        Loop
    End With
  
  
  
  Case 3 'Cancelacion
    strSQL = strSQL & ",'C'"
    With lswCancelacion
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Fecha", 2200, vbCenter
            .ColumnHeaders.Add , , "Usuario", 1800, vbCenter
            .ColumnHeaders.Add , , "Tipo Doc", 1800
            .ColumnHeaders.Add , , "Transacción", 2200
            .ColumnHeaders.Add , , "Forma de Pago", 2200
            .ColumnHeaders.Add , , "Num. Referencia", 2200
            .ColumnHeaders.Add , , "Monto Doc.", 2200, vbRightJustify
            .ColumnHeaders.Add , , "Monto Apl.", 2200, vbRightJustify
            .ColumnHeaders.Add , , "Cliente", 3200
            .ColumnHeaders.Add , , "Notas", 3200
    
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!Registro_Fecha)
                itmX.SubItems(1) = rs!Registro_Usuario
                itmX.SubItems(2) = rs!TIPO_DOCUMENTO
                itmX.SubItems(3) = rs!Cod_Transaccion
                itmX.SubItems(4) = rs!FormaPago
                itmX.SubItems(5) = rs!Num_Referencia
                itmX.SubItems(6) = Format(rs!Monto_Doc, "Standard")
                itmX.SubItems(7) = Format(rs!Monto, "Standard")
                itmX.SubItems(8) = rs!Cliente_Nombre
                itmX.SubItems(9) = rs!observaciones
            rs.MoveNext
        Loop
    End With
End Select


rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub vGrid_DblClick(ByVal col As Long, ByVal Row As Long)
Dim vOperacion As Long
Dim frm As Form

On Error GoTo vError

vOperacion = 0

vGrid.Row = Row
vGrid.col = col
If col = 2 Then
   vOperacion = vGrid.Text
   If vOperacion = 0 Then Exit Sub
  
    Call sbFormsCall("frmCxC_Cuentas")
    
    For Each frm In Forms
      If (UCase(frm.Name) = UCase("frmCxC_Cuentas")) Then
        Call frm.sbConsultaExterna(vOperacion)
        Exit For
      End If
    Next frm
     
End If

Exit Sub

vError:

End Sub
