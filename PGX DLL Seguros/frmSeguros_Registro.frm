VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSeguros_Registro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Pólizas"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
      _ExtentY        =   12303
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
      ItemCount       =   5
      Item(0).Caption =   "Recepción"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "txtCodigoAlterno"
      Item(0).Control(1)=   "txtNotas"
      Item(0).Control(2)=   "Label1(26)"
      Item(0).Control(3)=   "Label1(8)"
      Item(0).Control(4)=   "fraPoliza"
      Item(1).Caption =   "Coberturas"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lswCoberturas"
      Item(1).Control(1)=   "lswPolizasRelacionadas"
      Item(1).Control(2)=   "Label1(25)"
      Item(1).Control(3)=   "Label1(24)"
      Item(2).Caption =   "Resumen"
      Item(2).ControlCount=   28
      Item(2).Control(0)=   "txtBalance"
      Item(2).Control(1)=   "txtCobroPriDeduc"
      Item(2).Control(2)=   "txtCobroUltMov"
      Item(2).Control(3)=   "txtCobroTotal"
      Item(2).Control(4)=   "txtPagoPriMov"
      Item(2).Control(5)=   "txtPagoUltMov"
      Item(2).Control(6)=   "txtPagoTotal"
      Item(2).Control(7)=   "Label1(15)"
      Item(2).Control(8)=   "Label1(14)"
      Item(2).Control(9)=   "Label1(13)"
      Item(2).Control(10)=   "Label1(12)"
      Item(2).Control(11)=   "Label1(9)"
      Item(2).Control(12)=   "Label1(2)"
      Item(2).Control(13)=   "Label1(5)"
      Item(2).Control(14)=   "Label1(6)"
      Item(2).Control(15)=   "Label1(7)"
      Item(2).Control(16)=   "txtComisionVendedor"
      Item(2).Control(17)=   "txtComisionInterna"
      Item(2).Control(18)=   "txtComisionComercializa"
      Item(2).Control(19)=   "txtResultados"
      Item(2).Control(20)=   "txtComisionComercializaNeta"
      Item(2).Control(21)=   "chkComisionVendedorInformativa"
      Item(2).Control(22)=   "Label1(18)"
      Item(2).Control(23)=   "Label1(17)"
      Item(2).Control(24)=   "Label1(16)"
      Item(2).Control(25)=   "Label1(20)"
      Item(2).Control(26)=   "Label1(21)"
      Item(2).Control(27)=   "Label1(22)"
      Item(3).Caption =   "Pagos"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswPagos"
      Item(4).Caption =   "Cobros"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lswCobros"
      Begin XtremeSuiteControls.ListView lswCobros 
         Height          =   6495
         Left            =   -70000
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   11456
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
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswPagos 
         Height          =   6495
         Left            =   -70000
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   11456
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
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswCoberturas 
         Height          =   2655
         Left            =   -69520
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   8655
         _Version        =   1441793
         _ExtentX        =   15266
         _ExtentY        =   4683
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
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswPolizasRelacionadas 
         Height          =   2655
         Left            =   -69520
         TabIndex        =   62
         Top             =   4080
         Visible         =   0   'False
         Width           =   8655
         _Version        =   1441793
         _ExtentX        =   15266
         _ExtentY        =   4683
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
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox fraPoliza 
         Height          =   5295
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   9340
         _StockProps     =   79
         Appearance      =   16
         BorderStyle     =   2
         Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoSeguro 
            Height          =   255
            Left            =   8760
            TabIndex        =   37
            Top             =   2760
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoCuenta 
            Height          =   255
            Left            =   8760
            TabIndex        =   38
            Top             =   3240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin XtremeSuiteControls.ComboBox cboComercializadora 
            Height          =   330
            Left            =   1560
            TabIndex        =   54
            Top             =   1560
            Width           =   7815
            _Version        =   1441793
            _ExtentX        =   13785
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
         End
         Begin XtremeSuiteControls.ComboBox cboAseguradora 
            Height          =   330
            Left            =   1560
            TabIndex        =   55
            Top             =   2400
            Width           =   7815
            _Version        =   1441793
            _ExtentX        =   13785
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
         End
         Begin XtremeSuiteControls.ComboBox cboTipoPago 
            Height          =   315
            Left            =   4920
            TabIndex        =   56
            Top             =   4200
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
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
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInicia 
            Height          =   315
            Left            =   4920
            TabIndex        =   57
            Top             =   4560
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.DateTimePicker dtpRenueva 
            Height          =   315
            Left            =   4920
            TabIndex        =   58
            Top             =   4920
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.CheckBox chkPagador 
            Height          =   270
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Pagador?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkClienteCor 
            Height          =   270
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Corporativo?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   330
            Left            =   1560
            TabIndex        =   69
            Top             =   120
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
            Height          =   330
            Left            =   3240
            TabIndex        =   70
            Top             =   120
            Width           =   6135
            _Version        =   1441793
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPagador_Cedula 
            Height          =   330
            Left            =   1560
            TabIndex        =   71
            Top             =   600
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
         Begin XtremeSuiteControls.FlatEdit txtPagador_Nombre 
            Height          =   330
            Left            =   3240
            TabIndex        =   72
            Top             =   600
            Width           =   6135
            _Version        =   1441793
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtClienteCorId 
            Height          =   330
            Left            =   1560
            TabIndex        =   73
            Top             =   960
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
         Begin XtremeSuiteControls.FlatEdit txtClienteCorDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   74
            Top             =   960
            Width           =   6135
            _Version        =   1441793
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVendedorCod 
            Height          =   330
            Left            =   1560
            TabIndex        =   78
            Top             =   1920
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVendedorDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   79
            Top             =   1920
            Width           =   6135
            _Version        =   1441793
            _ExtentX        =   10821
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoSeguroCod 
            Height          =   330
            Left            =   1560
            TabIndex        =   80
            Top             =   2760
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoSeguroDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   81
            Top             =   2760
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoCuentaCod 
            Height          =   330
            Left            =   1560
            TabIndex        =   82
            Top             =   3240
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoCuentaDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   83
            Top             =   3240
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPago_Adelantado 
            Height          =   330
            Left            =   1560
            TabIndex        =   76
            Top             =   3840
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   330
            Left            =   1560
            TabIndex        =   84
            Top             =   4200
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   330
            Left            =   2520
            TabIndex        =   85
            Top             =   4560
            Width           =   735
            _Version        =   1441793
            _ExtentX        =   1296
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   330
            Left            =   1560
            TabIndex        =   86
            Top             =   4920
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstado 
            Height          =   330
            Left            =   7800
            TabIndex        =   89
            Top             =   4200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
         Begin XtremeSuiteControls.FlatEdit txtNumCuota 
            Height          =   330
            Left            =   7800
            TabIndex        =   90
            Top             =   4560
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
         Begin XtremeSuiteControls.FlatEdit txtOperacion 
            Height          =   330
            Left            =   7800
            TabIndex        =   91
            Top             =   4920
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
         Begin XtremeSuiteControls.ComboBox cboPrideduc 
            Height          =   330
            Left            =   4920
            TabIndex        =   92
            Top             =   3840
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
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
         Begin XtremeSuiteControls.ComboBox cboFrecuencia 
            Height          =   330
            Left            =   7800
            TabIndex        =   93
            Top             =   3840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin VB.Label Label1 
            Caption         =   "Frecuencia ..:"
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
            Index           =   19
            Left            =   6600
            TabIndex        =   94
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "P.Adelantado"
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
            Left            =   120
            TabIndex        =   77
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pri.Deduc."
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
            Height          =   375
            Index           =   4
            Left            =   3600
            TabIndex        =   75
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label lblPlazo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
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
            Height          =   252
            Left            =   120
            TabIndex        =   53
            Top             =   4560
            Width           =   732
         End
         Begin VB.Label lblPagador 
            Caption         =   "Tipo Cobro"
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
            TabIndex        =   52
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblContrato 
            Caption         =   "Tipo Seguro"
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
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   1092
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota"
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
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   4920
            Width           =   732
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   49
            Top             =   4200
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Identificación"
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
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   1212
         End
         Begin VB.Label Label1 
            Caption         =   "Vendedor"
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
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Estado ..:"
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
            Index           =   10
            Left            =   6600
            TabIndex        =   46
            Top             =   4200
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. Cuota.:"
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
            Index           =   11
            Left            =   6600
            TabIndex        =   45
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Aseguradora"
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
            Left            =   120
            TabIndex        =   44
            Top             =   2400
            Width           =   1092
         End
         Begin VB.Label Label3 
            Caption         =   "Comercializa"
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
            Left            =   120
            TabIndex        =   43
            Top             =   1560
            Width           =   1092
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Pago"
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
            Height          =   375
            Index           =   1
            Left            =   3600
            TabIndex        =   42
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Inicia"
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
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   41
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Renovación"
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
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   40
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Operación..:"
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
            Index           =   23
            Left            =   6600
            TabIndex        =   39
            Top             =   4920
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkComisionVendedorInformativa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Comisión del Vendedor es Informativa?"
         Enabled         =   0   'False
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
         Height          =   456
         Left            =   -68800
         TabIndex        =   29
         Top             =   5220
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtComisionComercializaNeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63640
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4380
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtResultados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63640
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3780
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtComisionComercializa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67600
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4380
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtComisionInterna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67600
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3780
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtComisionVendedor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67600
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4740
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPagoTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67720
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPagoUltMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPagoPriMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67720
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtCobroTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63520
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtCobroUltMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtCobroPriDeduc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -63520
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   570
         Left            =   1680
         TabIndex        =   87
         Top             =   5880
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   1005
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
      Begin XtremeSuiteControls.FlatEdit txtCodigoAlterno 
         Height          =   330
         Left            =   5520
         TabIndex        =   88
         Top             =   6600
         Width           =   3975
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Neta - (%Vendedor)..:"
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
         Index           =   22
         Left            =   -65560
         TabIndex        =   35
         Top             =   4380
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "Resultados.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   21
         Left            =   -65320
         TabIndex        =   34
         Top             =   3780
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Comercializadora..:"
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
         Index           =   20
         Left            =   -69400
         TabIndex        =   33
         Top             =   4380
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comisiones Registradas...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   16
         Left            =   -69880
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor.:"
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
         Index           =   17
         Left            =   -69400
         TabIndex        =   31
         Top             =   4740
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Administración..:"
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
         Index           =   18
         Left            =   -69400
         TabIndex        =   30
         Top             =   3780
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha 1er. Pago ..:"
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
         Index           =   7
         Left            =   -69280
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ult. Pago .:"
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
         Index           =   6
         Left            =   -69280
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Realizados ...:"
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
         Left            =   -69280
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Información de Pagos...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69640
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "Información de Cobros...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   -65560
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha 1er. Deduc..:"
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
         Index           =   12
         Left            =   -65200
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ult. Mov.:"
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
         Index           =   13
         Left            =   -65200
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Realizados ...:"
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
         Index           =   14
         Left            =   -65200
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Balance de Cobranza.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   15
         Left            =   -65200
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "Coberturas del Seguro...:"
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
         Index           =   24
         Left            =   -69640
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "Pólizas Vinculadas (HIjas)...:"
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
         Index           =   25
         Left            =   -69640
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   8
         Left            =   240
         TabIndex        =   5
         Top             =   5880
         Width           =   732
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Código Alterno"
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
         Index           =   26
         Left            =   3720
         TabIndex        =   4
         Top             =   6600
         Width           =   1452
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8325
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario Activa"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha Activa"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario - Cierra"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha Cierre"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0101
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0220
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0340
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0455
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0573
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":069D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":07C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":08DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":09DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Registro.frx":0AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   1920
      TabIndex        =   65
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
   Begin MSComctlLib.Toolbar tlbAux 
      Height          =   330
      Left            =   7680
      TabIndex        =   66
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Activar"
            Key             =   "Activar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   495
      Left            =   1800
      TabIndex        =   67
      Top             =   480
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Left            =   240
      TabIndex        =   68
      Top             =   480
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Póliza"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label lblNombre 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   4935
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   255
      Left            =   4800
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmSeguros_Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date

Function fxPersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from Socios where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxPersonaNombre = ""
Else
 fxPersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function



Private Sub ReporteBoleta()
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

'strRuta = SIFGlobal.fxPathReportes("CxC_BoletaActivacion.rpt")
'
'With frmContenedor.Crt
' .Reset
' .WindowShowRefreshBtn = True
' .WindowShowPrintSetupBtn = True
' .WindowState = crptMaximized
' .WindowShowSearchBtn = True
' .WindowTitle = "CxC...: Boleta de Activación"
' .ReportFileName = strRuta
'
' .Connect = glogon.ConectRPT
'
' .SelectionFormula = "{SEGUROS_REGISTRO.num_poliza}=" & txtPoliza.Text
' .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
' .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
' .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'
' .SubreportToChange = "sbAsiento"
' .StoredProcParam(0) = "CxC_FRM"
' .StoredProcParam(1) = txtPoliza.Text
' .StoredProcParam(2) = 0
' .PrintReport
'End With

Me.MousePointer = vbDefault

End Sub



Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo pError

fxValida = True
vMensaje = ""


If Len(txtPoliza.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- No se indicó el número de la póliza!"


If IsNumeric(txtPago_Adelantado.Text) Then
 If CCur(txtPago_Adelantado.Text) < 0 Then vMensaje = vMensaje & vbCrLf & "- El Pago Adelantado  NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Pago Adelantado No es Inválido"
End If


If IsNumeric(txtPlazo) Then
 If txtPlazo.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo es Inválido"
End If


If IsNumeric(txtMonto.Text) Then
 If CCur(txtMonto.Text) < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto No es Inválido"
End If

If chkClienteCor.Value = vbChecked Then
   If txtClienteCorId.Text = "" Then
        vMensaje = vMensaje & vbCrLf & "- No se indicó el Cliente Corporativo de referencia!"
   End If
End If


If chkPagador.Value = vbChecked Then
   If txtPagador_Cedula.Text = "" Then
        vMensaje = vMensaje & vbCrLf & "- No se indicó los datos del Pagador del Seguro!"
   End If
End If

strSQL = "select count(*) as Existe from SEGUROS_TIPOS_PRODUCTOS where COD_PRODUCTO ='" & txtTipoSeguroCod.Text & "' and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de seguro no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from SEGUROS_TIPOS_COBRO where TIPO_COBRO ='" & txtTipoCuentaCod.Text & "' and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Cobro no se encuentra Activo!"
rs.Close


strSQL = "select count(*) as Existe from SEGUROS_Vendedores where cod_vendedor = " & txtVendedorCod.Text & " and Activo = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El vendedor no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Socios where cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- La persona no existe en la base de datos!"
rs.Close


'Validar que si es Cargo Automático exista el registro de la tarjeta(Cambiar por SP todo el codigo de validacion)
If UCase(txtTipoCuentaCod.Text) = "CAU" Then
    strSQL = "select count(*) as Existe from AFI_PERSONA_TARJETAS where cedula = '" & txtCedula.Text & "'"
    Call OpenRecordSet(rs, strSQL)
     If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Se ha indicado Cargo Automático pero la Persona no tiene ninguna tarjeta registrada!"
    rs.Close
End If



'Verificar que la persona no tenga prestamos en Cobro Judicial Activos
strSQL = "select isnull(count(*),0) as Existe from Reg_Creditos" _
       & " where estado = 'A' and proceso = 'J' and cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene créditos en Cobro Judicial"
rs.Close


If Len(vMensaje) > 0 Then
 fxValida = False
 MsgBox vMensaje, vbExclamation
End If

Exit Function

pError:
    vMensaje = vMensaje & vbCrLf & Err.Description
    If Len(vMensaje) > 0 Then
        fxValida = False
        MsgBox vMensaje, vbExclamation
    End If

End Function


Private Sub cboComercializadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorCod.SetFocus

End Sub


Private Sub cboTipoPago_Click()

On Error GoTo vError

If vPaso Then Exit Sub

Select Case Mid(cboTipoPago.Text, 1, 1)
    Case "A" 'Anual
       txtPlazo.Text = 1
       dtpInicia.Tag = 12
    Case "T" 'Trimerstral
       txtPlazo.Text = 999
       dtpInicia.Tag = 3
    Case "S" 'Semestral
       txtPlazo.Text = 999
       dtpInicia.Tag = 6
    Case "M"
       txtPlazo.Text = 999
       dtpInicia.Tag = 1
    Case Else
       txtPlazo.Text = 999
       dtpInicia.Tag = 1
End Select

dtpRenueva.Value = DateAdd("m", dtpInicia.TabStop, dtpInicia.Value)

Exit Sub

vError:

End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicia.SetFocus
End Sub


Private Sub chkClienteCor_Click()

If chkClienteCor.Value = vbChecked Then
    txtClienteCorId.Visible = True
Else
    txtClienteCorId.Visible = False
End If

txtClienteCorDesc.Visible = txtClienteCorId.Visible

End Sub

Private Sub chkPagador_Click()
If chkPagador.Value = vbChecked Then
    txtPagador_Cedula.Visible = True
Else
    txtPagador_Cedula.Visible = False
End If

txtPagador_Nombre.Visible = txtPagador_Cedula.Visible
End Sub

Private Sub dtpInicia_Change()
On Error GoTo vError

dtpRenueva.Value = DateAdd("m", dtpInicia.TabStop, dtpInicia.Value)

vError:

End Sub

Private Sub dtpInicia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpRenueva.SetFocus
End Sub

Private Sub dtpRenueva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtPoliza.Text = "" Then txtPoliza.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 num_poliza from SEGUROS_REGISTRO"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where num_poliza > '" & txtPoliza & "' order by num_poliza asc"
Else
   strSQL = strSQL & " where num_poliza < '" & txtPoliza & "' order by num_poliza desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtPoliza.Text = rs!Num_Poliza
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub FlatScrollBarTipoSeguro_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoSeguro.Tag = "" Then FlatScrollBarTipoSeguro.Tag = 0

strSQL = "select Top 1 COD_PRODUCTO,Descripcion from SEGUROS_TIPOS_PRODUCTOS"

If FlatScrollBarTipoSeguro.Value > CLng(FlatScrollBarTipoSeguro.Tag) Then
   strSQL = strSQL & " where COD_ASEGURADORA = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and COD_PRODUCTO > '" & txtTipoSeguroCod.Text _
          & "' and Activo = 1  and dbo.fxSeguros_ProductosAcceso(COD_ASEGURADORA,COD_PRODUCTO,'" & glogon.Usuario & "') = 1" _
          & " order by COD_PRODUCTO asc"
Else
   strSQL = strSQL & " where COD_ASEGURADORA = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "' and COD_PRODUCTO < '" & txtTipoSeguroCod.Text _
          & "' and Activo = 1  and dbo.fxSeguros_ProductosAcceso(COD_ASEGURADORA,COD_PRODUCTO,'" & glogon.Usuario & "') = 1" _
          & " order by COD_PRODUCTO asc"
End If

FlatScrollBarTipoSeguro.Tag = FlatScrollBarTipoSeguro.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
        txtTipoSeguroCod.Text = rs!COD_PRODUCTO
        txtTipoSeguroDesc.Text = rs!Descripcion
Else
        txtTipoSeguroCod.Text = ""
        txtTipoSeguroDesc.Text = ""
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarTipoCuenta_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoCuenta.Tag = "" Then FlatScrollBarTipoCuenta.Tag = 0

strSQL = "select Top 1 TIPO_COBRO,Descripcion from SEGUROS_TIPOS_COBRO"

If FlatScrollBarTipoCuenta.Value > CLng(FlatScrollBarTipoCuenta.Tag) Then
   strSQL = strSQL & " where TIPO_COBRO > '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by TIPO_COBRO asc"
Else
   strSQL = strSQL & " where TIPO_COBRO < '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by TIPO_COBRO asc"
End If

FlatScrollBarTipoCuenta.Tag = FlatScrollBarTipoCuenta.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtTipoCuentaCod.Text = rs!TIPO_COBRO
  txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 17
 
vFecha = fxFechaServidor

'---------------------------------------------------------

cboFrecuencia.AddItem "Mensual"
cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
cboFrecuencia.Text = "Mensual"
        
'---------------------------------------------------------

With lswCoberturas.ColumnHeaders
  .Clear
  .Add , , "Cobertura", 4200
  .Add , , "Opcional", 1440, vbCenter
End With

With lswPolizasRelacionadas.ColumnHeaders
  .Clear
  .Add , , "No.Póliza", 2500
  .Add , , "Aseguradora", 1500, vbCenter
  .Add , , "Producto", 3500
  .Add , , "Monto", 1440, vbRightJustify
  .Add , , "Fecha", 2500
End With

With lswPagos.ColumnHeaders
  .Clear
  .Add , , "No.Cuota", 1000
  .Add , , "Fecha", 2500
  .Add , , "Remesa", 1240, vbCenter
  .Add , , "Trama", 1440, vbCenter
  .Add , , "Monto", 1440, vbRightJustify
  .Add , , "Comercializadora", 2000, vbCenter
  .Add , , "Vendedor", 1500, vbCenter
  .Add , , "Ganancia Interna", 2500, vbRightJustify
  .Add , , "Vendedor Info?", 1600, vbCenter
  
End With
 
With lswCobros.ColumnHeaders
  .Clear
  .Add , , "Fecha", 2500
  .Add , , "Monto", 1250, vbRightJustify
  .Add , , "Tipo Doc.", 2500
  .Add , , "Num. Doc.", 2500
  .Add , , "Usuario", 1600
  .Add , , "Concepto", 3400
  .Add , , "Operación", 1600
 
End With
 
 
 
 
strSQL = "select rtrim(cod_aseguradora) as 'IdX',  rtrim(nombre) as 'ItmX' from seguros_Aseguradoras"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)
 
strSQL = "select rtrim(cod_comercializadora) as 'IdX', rtrim(nombre) as 'ItmX' from seguros_comercializadoras where activo = 1"
Call sbCbo_Llena_New(cboComercializadora, strSQL, False, True)
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpiaPantalla


End Sub

Private Sub sbLimpiaPantalla()
 

Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
ImgAutorizacion.ToolTipText = "Pendiente: Consulta/Nuevo"
 
vPaso = True
 cboTipoPago.Clear
 cboTipoPago.AddItem "Mensual"
 cboTipoPago.AddItem "Anual"
 cboTipoPago.AddItem "Trimestral"
 cboTipoPago.AddItem "Semestral"
 cboTipoPago.Text = "Mensual"
vPaso = False
 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False

 txtCedula = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 
 chkClienteCor.Value = vbUnchecked
 txtClienteCorId.Text = ""
 txtClienteCorDesc.Text = ""
 
 chkPagador.Value = vbUnchecked
 txtPagador_Cedula.Text = ""
 txtPagador_Nombre.Text = ""
 
 
 txtVendedorCod.Text = ""
 txtVendedorDesc.Text = ""

 txtTipoSeguroCod.Text = ""
 txtTipoSeguroDesc.Text = ""

 txtTipoCuentaCod.Text = ""
 txtTipoCuentaDesc.Text = ""
   
 txtEstado.Text = "Pendiente"
   
 txtPago_Adelantado.Text = "0.00"
 txtMonto.Text = "0.00"
 txtPlazo.Text = "60"
 txtCuota.Text = "0.00"
  
 txtNumCuota.Text = 0
 txtOperacion.Text = 0
 
 txtBalance.Text = 0
 txtPagoPriMov.Text = ""
 txtPagoTotal.Text = 0
 txtPagoUltMov.Text = 0
 
 txtCobroPriDeduc.Text = ""
 txtCobroTotal.Text = 0
 txtCobroUltMov.Text = ""
 
 
 chkComisionVendedorInformativa.Value = vbUnchecked
 txtComisionComercializa.Text = 0
 txtComisionComercializaNeta.Text = 0
 txtComisionInterna.Text = 0
 txtComisionVendedor.Text = 0
 
 txtBalance.Text = 0
 txtResultados.Text = 0
 
 txtNotas = ""
 
 dtpInicia.Value = vFecha
 dtpInicia.Tag = 1
 
 
 tcMain.Item(0).Selected = True
 tcMain.Item(1).Enabled = False
 tcMain.Item(2).Enabled = False
 tcMain.Item(3).Enabled = False
 tcMain.Item(4).Enabled = False
 
 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 StatusBarX.Panels(4).Text = ""
 StatusBarX.Panels(5).Text = ""

 Call dtpInicia_Change
 Call chkClienteCor_Click
 Call chkPagador_Click
 
 
 '---------------------------------------------------------

Dim vProceso As Currency, i As Integer

vProceso = Format(vFecha, "YYYYMM")
cboPrideduc.AddItem vProceso
cboPrideduc.Text = vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboPrideduc.AddItem vProceso
Next i
       
'---------------------------------------------------------
 
End Sub



Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

vPaso = True

strSQL = "select Pol.*,Ts.Descripcion as 'TipoSeguroDesc', Per.Nombre, isnull(Pol.Estado,'P') as 'Estado'" _
       & ",Ven.Nombre as 'VendedorNombre',Tc.descripcion as 'TipoCuentaDesc',dbo.MyGetdate() as 'FechaServer'" _
       & ",rtrim(Ase.nombre) as 'AseguradoraDesc' " _
       & ",rtrim(com.Nombre) as 'ComercializadoraDesc' " _
       & ",isnull(Cc.Nombre,'-1') as 'ClienteCorporativo', isnull(Np.Nombre, '-1') as 'PagadorNombre'" _
       & " from SEGUROS_REGISTRO Pol inner join SEGUROS_TIPOS_PRODUCTOS Ts on Pol.COD_PRODUCTO = Ts.COD_PRODUCTO and Pol.Cod_Aseguradora = Ts.Cod_Aseguradora" _
       & " inner join SEGUROS_ASEGURADORAS Ase on Pol.Cod_Aseguradora = Ase.Cod_Aseguradora" _
       & " inner join SEGUROS_COMERCIALIZADORAS Com on Pol.cod_Comercializadora = com.cod_Comercializadora" _
       & " inner join Socios Per on Pol.cedula = Per.cedula" _
       & " left join SEGUROS_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
       & " left join SEGUROS_TIPOS_COBRO Tc on Pol.TIPO_COBRO = Tc.TIPO_COBRO" _
       & " left join SEGUROS_CLIENTE_CORPORATIVO Cc on Pol.Cod_Cliente_Corporativo  = Cc.Cod_Cliente_Corporativo" _
       & " left join SOCIOS Np on Pol.Cedula_Pagador = Np.Cedula" _
       & " where Pol.num_poliza = '" & txtPoliza.Text & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
 
 tcMain.Item(0).Selected = True
 
 txtCedula.SetFocus
 
 vFecha = rs!FechaServer
 
 txtCodigoAlterno.Text = Trim(rs!Codigo_Alterno & "")
 
 txtCedula.Text = rs!Cedula
 
 
 txtNombre.Text = rs!Nombre
 lblNombre.Caption = txtNombre.Text
 
 If rs!ClienteCorporativo <> "-1" Then
   chkClienteCor.Value = vbChecked
   txtClienteCorId.Text = rs!Cod_Cliente_Corporativo & ""
   txtClienteCorDesc.Text = rs!ClienteCorporativo
   Call chkClienteCor_Click
 End If
 
 If rs!PagadorNombre <> "-1" Then
   chkPagador.Value = vbChecked
   txtPagador_Cedula.Text = rs!Cedula_Pagador & ""
   txtPagador_Nombre.Text = rs!PagadorNombre
   Call chkPagador_Click
 End If
 
 
 txtVendedorCod.Text = rs!cod_vendedor
 txtVendedorDesc.Text = rs!VendedorNombre
 
 txtTipoSeguroCod.Text = rs!COD_PRODUCTO
 txtTipoSeguroDesc.Text = rs!TipoSeguroDesc
 txtTipoCuentaCod.Text = rs!TIPO_COBRO
 txtTipoCuentaDesc.Text = rs!TipoCuentaDesc
 
 
 txtMonto.Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
 txtPlazo.Text = rs!Plazo
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 txtNumCuota.Text = rs!Num_Cuota & ""
 
 txtCobroPriDeduc.Text = rs!Cobrado_Primer_Deduc & ""
 txtCobroTotal.Text = Format(rs!Cobrado_Total & "", "Standard")
 txtCobroUltMov.Text = rs!Cobrado_Fecha_Ult & ""
 
 txtPagoPriMov.Text = rs!Pagado_Primer_Pago & ""
 txtPagoTotal.Text = Format(rs!Pagado_Total, "Standard")
 txtPagoTotal.ToolTipText = "Neto ..:" & Format(rs!Pagado_Total_Neto, "Standard")
 
 txtPagoUltMov.Text = rs!Pagado_Fecha_Ult & ""

 txtBalance.Text = Format(rs!Cobrado_Total - rs!Pagado_Total, "Standard")

 txtNotas = IIf(IsNull(rs!notas), "", rs!notas)

 tcMain.Item(0).Selected = True
 tcMain.Item(1).Enabled = True
 tcMain.Item(2).Enabled = True
 tcMain.Item(3).Enabled = True
 tcMain.Item(4).Enabled = True

 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False
 
 lswCoberturas.Checkboxes = False
 lswPolizasRelacionadas.Checkboxes = False
 
 Select Case rs!Estado
   Case "P" 'Pendiente
      txtEstado.Text = "Pendiente"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
      ImgAutorizacion.ToolTipText = "Activación: Pendiente"
      
      tlbAux.Buttons.Item(1).Enabled = True
      tlbAux.Buttons.Item(3).Enabled = True
   
      lswCoberturas.Checkboxes = True
      lswPolizasRelacionadas.Checkboxes = True
   
   Case "A"
      txtEstado.Text = "Activa"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(5).Picture
      ImgAutorizacion.ToolTipText = "Póliza Activada!"
      tlbAux.Buttons.Item(3).Enabled = True
   Case "C"
      txtEstado.Text = "Cerrada"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(6).Picture
      ImgAutorizacion.ToolTipText = "Póliza Cerrada (Inactivada)"
  End Select

 txtOperacion.Text = rs!Operacion & ""

 txtComisionInterna.Text = Format(rs!Comision_Interna_Total, "Standard")
 txtComisionComercializa.Text = Format(rs!Comision_Comercializa_Total, "Standard")
 txtComisionVendedor.Text = Format(rs!Comision_Vendedor_Total, "Standard")
 
 If rs!Comision_Vendedor_Informativa = 1 Then
    txtComisionComercializaNeta.Text = txtComisionComercializa.Text
 Else
    txtComisionComercializaNeta.Text = Format(rs!Comision_Comercializa_Total - rs!Comision_Vendedor_Total, "Standard")
 End If
 txtResultados.Text = Format(rs!Comision_Interna_Total - rs!Comision_Comercializa_Total, "Standard")

 StatusBarX.Panels(1).Text = rs!Registro_Usuario
 StatusBarX.Panels(2).Text = rs!Registro_Fecha
 StatusBarX.Panels(3).Text = rs!Activa_Usuario & ""
 StatusBarX.Panels(4).Text = rs!ACTIVA_FECHA & ""
 StatusBarX.Panels(5).Text = rs!Cierra_usuario & ""
 StatusBarX.Panels(6).Text = rs!Cierra_fecha & ""
 
 txtBalance.Text = Format(rs!Cobrado_Total - rs!Pagado_Total, "Standard")
 
 Call sbCboAsignaDato(cboAseguradora, rs!AseguradoraDesc, True, rs!cod_Aseguradora)
 Call sbCboAsignaDato(cboComercializadora, rs!ComercializadoraDesc, True, rs!cod_Comercializadora)


 dtpInicia.Value = rs!FECHA_INICIO
 dtpRenueva.Value = rs!Fecha_Renovacion

 If rs!Tipo_Pago = "M" Then
   cboTipoPago.Text = "Mensual"
 Else
   cboTipoPago.Text = "Anual"
 End If


    If Not IsNull(rs!Pago_Adelantado) Then
        txtPago_Adelantado.Text = Format(rs!Pago_Adelantado, "Standard")
    End If
    
    If Not IsNull(rs!PriDeduc) Then
        Call sbCboAsignaDato(cboPrideduc, rs!PriDeduc, True, rs!PriDeduc)
    End If

Else
 
 If vEdita Then
    MsgBox "No existe la Póliza, verifique!", vbCritical
 End If

End If

rs.Close

 
vPaso = False

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra pendiente", vbExclamation
    Exit Sub
End If

strSQL = "delete SEGUROS_REGISTRO where num_poliza = '" & txtPoliza.Text & "'"
Call ConectionExecute(strSQL)

MsgBox "Poliza Eliminada!", vbInformation

Call sbLimpiaPantalla

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub








Private Sub lswCoberturas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String


If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
   strSQL = "exec spSeguros_Poliza_Coberturas_Add '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & Item.Tag _
          & "',1,'" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)

Else
   strSQL = "exec spSeguros_Poliza_Coberturas_Add '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & Item.Tag _
          & "',0,'" & glogon.Usuario & "'"
   
   If Item.SubItems(1) = "Sí" Then
      Call ConectionExecute(strSQL)
   Else
          Item.Checked = True
   End If
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswPolizasRelacionadas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spSeguros_Poliza_Relacionadas_Add '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & Item.SubItems(1) _
       & "','" & Item.Text & "'," & IIf(Item.Checked = True, 1, 0) & ",'" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 1 'Coberturas & Polizas Relacionadas
    Call sbCoberturas
    Call sbPolizas_Relacionadas
  Case 2 'Resumen
  Case 3 'Pagos
    Call sbHistorial_Pagos
  Case 4 'Cobros
    Call sbHistorial_Cobros
  Case Else
End Select
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "InsertAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtPoliza.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      If fraPoliza.Enabled Then
          txtCedula.SetFocus
      End If
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtPoliza.Text = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "nombre"
'       gBusquedas.Orden = "nombre"
'       gBusquedas.Consulta = "select cod_abogado,nombre from Cbr_Cj_Abogados"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = gBusquedas.Resultado
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Select Case Button.Key
  Case "Activar"
        i = MsgBox("Esta seguro que desea >> Activar << esta Póliza", vbYesNo)
        If i = vbYes Then
            
            strSQL = "exec spSeguros_PolizaActiva '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)
        
            'BITACORA
            Call Bitacora("Registra", "Activación de la Póliza: " & txtPoliza & "_" & cboAseguradora.ItemData(cboAseguradora.ListIndex))
            
            MsgBox "Póliza Activada Satisfactoriamente!", vbInformation
            
        End If
        
  Case "Cerrar"
        GLOBALES.gTag = txtPoliza.Text
        GLOBALES.gTag2 = cboAseguradora.ItemData(cboAseguradora.ListIndex)
        
        frmSeguros_PolizaCierre.Show vbModal, Me
  
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





Private Sub txtClienteCorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboComercializadora.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "cod_cliente_Corporativo"
   gBusquedas.Filtro = " and activo = 1"
   gBusquedas.Consulta = "Select cod_cliente_Corporativo,Nombre from SEGUROS_CLIENTE_CORPORATIVO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtClienteCorId.Text = gBusquedas.Resultado
      txtClienteCorDesc.Text = gBusquedas.Resultado2
   End If
End If

End Sub

Private Sub txtClienteCorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtClienteCorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_cliente_Corporativo"
   gBusquedas.Orden = "cod_cliente_Corporativo"
   gBusquedas.Filtro = " and activo = 1"
   gBusquedas.Consulta = "Select cod_cliente_Corporativo,Nombre from SEGUROS_CLIENTE_CORPORATIVO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtClienteCorId.Text = gBusquedas.Resultado
      txtClienteCorDesc.Text = gBusquedas.Resultado2
   End If
End If
End Sub

Private Sub txtCuota_GotFocus()
On Error GoTo vError

txtCuota.Text = CCur(txtCuota.Text)

vError:
End Sub

Private Sub txtCuota_LostFocus()
On Error GoTo vError

txtCuota.Text = Format(CCur(txtCuota.Text), "Standard")

vError:

End Sub


Private Sub txtPagador_Cedula_LostFocus()
On Error GoTo vError

Me.MousePointer = vbHourglass


txtPagador_Nombre.Text = fxPersonaNombre(txtPagador_Cedula.Text)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtPagador_Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPagador_Nombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtPagador_Cedula.Text = gBusquedas.Resultado
      txtPagador_Nombre.Text = gBusquedas.Resultado2
   End If
End If
End Sub

Private Sub txtPago_Adelantado_GotFocus()
On Error GoTo vError

txtPago_Adelantado.Text = CCur(txtPago_Adelantado.Text)

vError:
End Sub

Private Sub txtPago_Adelantado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtPago_Adelantado_LostFocus()
On Error GoTo vError

txtPago_Adelantado.Text = Format(CCur(txtPago_Adelantado.Text), "Standard")

vError:
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtPoliza_LostFocus()
  Call sbConsulta
End Sub

Private Sub txtTipoSeguroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "COD_PRODUCTO"
   gBusquedas.Orden = "COD_PRODUCTO"
   gBusquedas.Filtro = " and Activo = 1 and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
                     & "' and dbo.fxSeguros_ProductosAcceso(COD_ASEGURADORA,COD_PRODUCTO,'" & glogon.Usuario & "') = 1"
   gBusquedas.Consulta = "Select COD_PRODUCTO,Descripcion  from SEGUROS_TIPOS_PRODUCTOS"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoSeguroCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SEGUROS_TIPOS_PRODUCTOS where COD_PRODUCTO = '" & txtTipoSeguroCod.Text _
       & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTipoSeguroCod.Text = rs!COD_PRODUCTO
   txtTipoSeguroDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtTipoSeguroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1 and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
                     & "' and dbo.fxSeguros_ProductosAcceso(COD_ASEGURADORA,COD_PRODUCTO,'" & glogon.Usuario & "') = 1"
   gBusquedas.Consulta = "Select COD_PRODUCTO,Descripcion  from SEGUROS_TIPOS_PRODUCTOS"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If
End Sub


'--Tipo de Cuenta
Private Sub txtTipoCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "TIPO_COBRO"
   gBusquedas.Orden = "TIPO_COBRO"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select TIPO_COBRO,Descripcion  from SEGUROS_TIPOS_COBRO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoCuentaCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from SEGUROS_TIPOS_COBRO where TIPO_COBRO = '" & txtTipoCuentaCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTipoCuentaCod.Text = rs!TIPO_COBRO
   txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtTipoCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select TIPO_COBRO,Descripcion  from SEGUROS_TIPOS_COBRO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If
End Sub




Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub

Private Sub sbCoberturas()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

vPaso = True

strSQL = "exec spSeguros_Poliza_Coberturas_Consulta '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "'"
Call OpenRecordSet(rs, strSQL)
With lswCoberturas.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, rs!cod_Cobertura, rs!Descripcion)
          itmX.SubItems(1) = IIf((rs!Opcional = 1), "Sí", "No")
          itmX.Tag = rs!cod_Cobertura
          itmX.Checked = rs!Activada
      rs.MoveNext
    Loop
End With
rs.Close

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbPolizas_Relacionadas()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

strSQL = "exec spSeguros_Poliza_Relacionadas_Consulta '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "'"
Call OpenRecordSet(rs, strSQL)
With lswPolizasRelacionadas.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Num_Poliza)
          itmX.SubItems(1) = rs!cod_Aseguradora
          itmX.SubItems(2) = rs!Producto
          itmX.SubItems(3) = Format(rs!Monto, "Standard")
          itmX.SubItems(4) = rs!ACTIVA_FECHA & ""
          
          itmX.Checked = rs!RELACIONADA
      rs.MoveNext
    Loop
End With
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbHistorial_Pagos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

With lswPagos.ListItems
   .Clear
   strSQL = "exec spSeguros_Historal_Pagos '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "'"
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
      Set itmX = .Add(, , rs!Num_Cuota)
          itmX.SubItems(1) = rs!Registro_Fecha & ""
          itmX.SubItems(2) = rs!Remesa & ""
          itmX.SubItems(3) = rs!Trama_Id & ""
          itmX.SubItems(4) = Format(rs!Monto, "Standard")
          itmX.SubItems(5) = Format(rs!COMISION_COMERCIALIZA, "Standard")
          itmX.SubItems(6) = Format(rs!Comision_Vendedor, "Standard")
          itmX.SubItems(7) = Format(rs!Comision_Interna, "Standard")
          itmX.SubItems(8) = IIf(rs!COMISION_VENDEDOR_INFO = 1, "Sí", "No")
          
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


Private Sub sbHistorial_Cobros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

With lswCobros.ListItems
   .Clear
   strSQL = "exec spSeguros_Historal_Cobros '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "'"
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
      Set itmX = .Add(, , rs!fecha)
          itmX.SubItems(1) = Format(rs!Total_Mov, "Standard")
          itmX.SubItems(2) = rs!Tipo & ""
          itmX.SubItems(3) = rs!nCon & ""
          itmX.SubItems(4) = rs!Usuario & ""
          itmX.SubItems(5) = rs!Concepto & ""
          itmX.SubItems(6) = rs!Id_Solicitud & ""
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



Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer, pClienteCor As String, pCodigoAlterno As String

On Error GoTo vError

If Trim(txtCodigoAlterno.Text) = "" Then
   pCodigoAlterno = "Null"
Else
   pCodigoAlterno = "'" & Trim(txtCodigoAlterno.Text) & "'"
End If
       
       
If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra pendiente (Solo se actualiza las notas y codigo alterno)", vbExclamation
    
    strSQL = "update SEGUROS_REGISTRO set notas = '" & txtNotas.Text & "', CODIGO_ALTERNO = " & pCodigoAlterno _
           & " where num_poliza = '" & txtPoliza.Text & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
    Call ConectionExecute(strSQL)
    
    Exit Sub
End If
              
              
       
If Not vEdita Then
   strSQL = "Insert SEGUROS_REGISTRO(cod_aseguradora,num_poliza,CEDULA,cod_vendedor,COD_PRODUCTO,TIPO_COBRO,NOTAS,MONTO,CUOTA,PLAZO" _
          & ", ESTADO, FECHA_INICIO, FECHA_RENOVACION, TIPO_PAGO, COD_COMERCIALIZADORA, COD_CLIENTE_CORPORATIVO, CEDULA_PAGADOR" _
          & ", CODIGO_ALTERNO, PRIDEDUC, PAGO_ADELANTADO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & " VALUES('" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtPoliza.Text & "','" & txtCedula.Text _
          & "' ," & txtVendedorCod.Text & ",'" & txtTipoSeguroCod.Text & "','" & txtTipoCuentaCod.Text _
          & "' ,'" & txtNotas.Text & "'," & CCur(txtMonto.Text) & "," & CCur(txtMonto.Text) & "," _
          & txtPlazo.Text & ",'P','" & Format(dtpInicia.Value, "yyyy/mm/dd") & "','" & Format(dtpRenueva.Value, "yyyy/mm/dd") _
          & "', '" & Mid(cboTipoPago.Text, 1, 1) & "','" & cboComercializadora.ItemData(cboComercializadora.ListIndex) _
          & "',  " & IIf((chkClienteCor.Value = vbChecked), ("'" & txtClienteCorId.Text & "'"), "Null") _
          & " , " & IIf((chkPagador.Value = vbChecked), ("'" & txtPagador_Cedula.Text & "'"), "Null") _
          & " , " & pCodigoAlterno & ", " & cboPrideduc.Text & ", " & CCur(txtPago_Adelantado.Text) & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
          
   'Registra Coberturas
   strSQL = "exec spSeguros_Poliza_Coberturas_Inicial '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" _
            & txtPoliza.Text & "','" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)
          
Else
   strSQL = "update SEGUROS_REGISTRO set cod_vendedor = " & txtVendedorCod.Text & ",COD_PRODUCTO = '" & txtTipoSeguroCod.Text & "',TIPO_COBRO = '" _
          & txtTipoCuentaCod.Text & "',notas = '" & txtNotas.Text & "',Monto = " & CCur(txtMonto.Text) & ", Cuota =  " & CCur(txtCuota.Text) _
          & ", Plazo = " & txtPlazo.Text & ", cedula = '" & txtCedula.Text & "', FECHA_INICIO = '" & Format(dtpInicia.Value, "yyyy/mm/dd") _
          & "', Fecha_Renovacion = '" & Format(dtpRenueva.Value, "yyyy/mm/dd") & "', Tipo_Pago = '" & Mid(cboTipoPago.Text, 1, 1) _
          & "', COD_COMERCIALIZADORA = '" & cboComercializadora.ItemData(cboComercializadora.ListIndex) _
          & "', COD_CLIENTE_CORPORATIVO = " & IIf((chkClienteCor.Value = vbChecked), ("'" & txtClienteCorId.Text & "'"), "Null") _
          & " , CEDULA_PAGADOR = " & IIf((chkPagador.Value = vbChecked), ("'" & txtPagador_Cedula.Text & "'"), "Null") _
          & ", CODIGO_ALTERNO = " & pCodigoAlterno & ", PRIDEDUC = " & cboPrideduc.Text & ", PAGO_ADELANTADO = " & CCur(txtPago_Adelantado.Text) _
          & " where num_poliza = '" & txtPoliza.Text & "' and cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
   Call ConectionExecute(strSQL)


End If

MsgBox "Seguro Registrado / Actualizado satisactoriamente!", vbInformation


Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub


Private Sub txtPagador_Cedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPagador_Nombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtPagador_Cedula.Text = gBusquedas.Resultado
      txtPagador_Nombre.Text = gBusquedas.Resultado2
   End If
End If

End Sub


Private Sub txtCedula_LostFocus()
'Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


txtNombre.Text = fxPersonaNombre(txtCedula)
lblNombre.Caption = txtNombre.Text


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboComercializadora.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If
End Sub

'--Vendedor
Private Sub txtVendedorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Orden = "Cod_Vendedor"
   gBusquedas.Filtro = " and cod_comercializadora = '" & cboComercializadora.ItemData(cboComercializadora.ListIndex) & "'"
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If


End Sub

Private Sub txtVendedorCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores where cod_Vendedor = " & txtVendedorCod.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtVendedorDesc.Text = rs!Nombre
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtVendedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from SEGUROS_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If
End Sub


Public Sub sbConsultaExterna(pPoliza As String, Optional pAseguradora As String = "")
 txtPoliza.Text = pPoliza
 Call sbConsulta
End Sub


Private Sub txtPoliza_Change()
 Call sbLimpiaPantalla

End Sub

Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And fraPoliza.Enabled Then txtCedula.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Num_Poliza"
   gBusquedas.Orden = "Num_Poliza"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Num_Poliza,Codigo_Alterno,Cedula from SEGUROS_REGISTRO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtPoliza.Text = gBusquedas.Resultado
      Call txtPoliza_LostFocus
   End If
End If
End Sub


