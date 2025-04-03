VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_FacturaEletronica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturación Electrónica"
   ClientHeight    =   8175
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8040
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12855
      _Version        =   1441793
      _ExtentX        =   22669
      _ExtentY        =   12298
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
      Item(0).Caption =   "Cortes"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbCortes"
      Item(0).Control(1)=   "GroupBox1"
      Item(1).Caption =   "Facturas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbFacturas_Filtros"
      Item(1).Control(1)=   "tcFiltros"
      Item(2).Caption =   "Clientes"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gbClientes_Filtros"
      Item(2).Control(1)=   "lswClientes"
      Item(3).Caption =   "Configuración"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "tcConfiguracion"
      Begin XtremeSuiteControls.ListView lswClientes 
         Height          =   5052
         Left            =   -69760
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   12372
         _Version        =   1441793
         _ExtentX        =   21823
         _ExtentY        =   8911
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.TabControl tcConfiguracion 
         Height          =   6612
         Left            =   -69880
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   12732
         _Version        =   1441793
         _ExtentX        =   22458
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
         ItemCount       =   3
         Item(0).Caption =   "Configuración"
         Item(0).ControlCount=   29
         Item(0).Control(0)=   "btnConfig_Main"
         Item(0).Control(1)=   "txtClienteIdProv"
         Item(0).Control(2)=   "chkActiva"
         Item(0).Control(3)=   "chkEmailDefault"
         Item(0).Control(4)=   "chkNotificaCliente"
         Item(0).Control(5)=   "txtTipoId"
         Item(0).Control(6)=   "txtIdentificacion"
         Item(0).Control(7)=   "txtRazonSocial"
         Item(0).Control(8)=   "txtEmail"
         Item(0).Control(9)=   "dtpFechaInicio"
         Item(0).Control(10)=   "btnConfig_Consec"
         Item(0).Control(11)=   "btnConfig_Portal"
         Item(0).Control(12)=   "cboMetodo"
         Item(0).Control(13)=   "chkIncluyePolizas"
         Item(0).Control(14)=   "chkIncluyePrincipal"
         Item(0).Control(15)=   "btnConfig_Clientes_Sinc"
         Item(0).Control(16)=   "Label1(28)"
         Item(0).Control(17)=   "Label1(5)"
         Item(0).Control(18)=   "Label1(4)"
         Item(0).Control(19)=   "Label1(3)"
         Item(0).Control(20)=   "Label1(2)"
         Item(0).Control(21)=   "Label1(1)"
         Item(0).Control(22)=   "Label1(0)"
         Item(0).Control(23)=   "chkMontoMaximo"
         Item(0).Control(24)=   "txtMontoMaximo"
         Item(0).Control(25)=   "gbConsecutivos"
         Item(0).Control(26)=   "GroupBox2"
         Item(0).Control(27)=   "txtCabys"
         Item(0).Control(28)=   "Label1(30)"
         Item(1).Caption =   "Exclusiones"
         Item(1).ControlCount=   5
         Item(1).Control(0)=   "bntConfig_Exclusiones(0)"
         Item(1).Control(1)=   "bntConfig_Exclusiones(1)"
         Item(1).Control(2)=   "bntConfig_Exclusiones(2)"
         Item(1).Control(3)=   "bntConfig_Exclusiones(3)"
         Item(1).Control(4)=   "lswExclusiones"
         Item(2).Caption =   "Reactivación"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "GroupBox3"
         Begin XtremeSuiteControls.ListView lswExclusiones 
            Height          =   6012
            Left            =   -67960
            TabIndex        =   98
            Top             =   480
            Visible         =   0   'False
            Width           =   10692
            _Version        =   1441793
            _ExtentX        =   18860
            _ExtentY        =   10604
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnConfig_Main 
            Height          =   612
            Left            =   11040
            TabIndex        =   50
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Guarda Configuración"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkActiva 
            Height          =   252
            Left            =   3480
            TabIndex        =   52
            Top             =   2160
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Facturación Activa?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkEmailDefault 
            Height          =   372
            Left            =   3480
            TabIndex        =   53
            Top             =   3240
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4466
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Utiliza Email Default ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkNotificaCliente 
            Height          =   252
            Left            =   3480
            TabIndex        =   54
            Top             =   3000
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notifica al Cliente ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtRazonSocial 
            Height          =   312
            Left            =   3480
            TabIndex        =   57
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   3480
            TabIndex        =   58
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInicio 
            Height          =   312
            Left            =   8280
            TabIndex        =   59
            Top             =   2160
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.PushButton btnConfig_Consec 
            Height          =   492
            Left            =   11040
            TabIndex        =   60
            Top             =   1080
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Consecutivos"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnConfig_Portal 
            Height          =   492
            Left            =   11040
            TabIndex        =   61
            Top             =   1680
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Accesos"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ComboBox cboMetodo 
            Height          =   312
            Left            =   8280
            TabIndex        =   62
            Top             =   2520
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2566
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.CheckBox chkIncluyePolizas 
            Height          =   252
            Left            =   6360
            TabIndex        =   63
            Top             =   3000
            Width           =   3372
            _Version        =   1441793
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Incluye Movimientos a Pólizas?   "
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkIncluyePrincipal 
            Height          =   252
            Left            =   6360
            TabIndex        =   64
            Top             =   3360
            Width           =   3372
            _Version        =   1441793
            _ExtentX        =   5948
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Incluye Movimientos a Principal?   "
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnConfig_Clientes_Sinc 
            Height          =   492
            Left            =   11040
            TabIndex        =   65
            Top             =   3600
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Sincronizar Clientes"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkMontoMaximo 
            Height          =   252
            Left            =   3000
            TabIndex        =   73
            Top             =   4200
            Width           =   4692
            _Version        =   1441793
            _ExtentX        =   8276
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Utiliza Monto Maximo de Facturación en Lote?"
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtMontoMaximo 
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
            Left            =   7800
            TabIndex        =   74
            Top             =   4200
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox gbConsecutivos 
            Height          =   1812
            Left            =   600
            TabIndex        =   75
            Top             =   4680
            Width           =   5172
            _Version        =   1441793
            _ExtentX        =   9123
            _ExtentY        =   3196
            _StockProps     =   79
            Caption         =   "Consecutivos"
            ForeColor       =   8421504
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.FlatEdit txtConsecFE 
               Height          =   312
               Left            =   1680
               TabIndex        =   76
               Top             =   360
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtConsecNC 
               Height          =   312
               Left            =   1680
               TabIndex        =   77
               Top             =   720
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtConsecND 
               Height          =   312
               Left            =   1680
               TabIndex        =   78
               Top             =   1080
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtConsecTE 
               Height          =   312
               Left            =   1680
               TabIndex        =   79
               Top             =   1440
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtSucursal 
               Height          =   312
               Left            =   3720
               TabIndex        =   101
               Top             =   600
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTerminal 
               Height          =   312
               Left            =   3720
               TabIndex        =   102
               Top             =   1200
               Width           =   1212
               _Version        =   1441793
               _ExtentX        =   2138
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   32
               Left            =   3360
               TabIndex        =   104
               Top             =   960
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Terminal"
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
               Alignment       =   4
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   31
               Left            =   3360
               TabIndex        =   103
               Top             =   360
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Sucursal"
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
               Alignment       =   4
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   6
               Left            =   -960
               TabIndex        =   83
               Top             =   360
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Consecutivo FE"
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   7
               Left            =   -960
               TabIndex        =   82
               Top             =   720
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Consecutivo NC"
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   8
               Left            =   -960
               TabIndex        =   81
               Top             =   1080
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Consecutivo ND"
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   9
               Left            =   -960
               TabIndex        =   80
               Top             =   1440
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Consecutivo TE"
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
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1812
            Left            =   6000
            TabIndex        =   84
            Top             =   4680
            Width           =   6252
            _Version        =   1441793
            _ExtentX        =   11028
            _ExtentY        =   3196
            _StockProps     =   79
            Caption         =   "Credenciales de Acceso [PORTAL PROVEEDOR]"
            ForeColor       =   8421504
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.FlatEdit txtPortal_Server 
               Height          =   312
               Left            =   2400
               TabIndex        =   85
               Top             =   360
               Width           =   2412
               _Version        =   1441793
               _ExtentX        =   4254
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
               PasswordChar    =   "/"
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPortal_DB 
               Height          =   312
               Left            =   2400
               TabIndex        =   86
               Top             =   720
               Width           =   2412
               _Version        =   1441793
               _ExtentX        =   4254
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
               PasswordChar    =   "/"
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPortal_User 
               Height          =   312
               Left            =   2400
               TabIndex        =   87
               Top             =   1080
               Width           =   2412
               _Version        =   1441793
               _ExtentX        =   4254
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
               PasswordChar    =   "/"
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPortal_Key 
               Height          =   312
               Left            =   2400
               TabIndex        =   88
               Top             =   1440
               Width           =   2412
               _Version        =   1441793
               _ExtentX        =   4254
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
               PasswordChar    =   "*"
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboPortal 
               Height          =   312
               Left            =   4800
               TabIndex        =   89
               Top             =   360
               Width           =   1452
               _Version        =   1441793
               _ExtentX        =   2566
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   10
               Left            =   -240
               TabIndex        =   93
               Top             =   360
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Server: "
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   11
               Left            =   -240
               TabIndex        =   92
               Top             =   720
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Base de datos: "
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   12
               Left            =   -240
               TabIndex        =   91
               Top             =   1080
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Usuario: "
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   13
               Left            =   -240
               TabIndex        =   90
               Top             =   1440
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Clave: "
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
            End
         End
         Begin XtremeSuiteControls.PushButton bntConfig_Exclusiones 
            Height          =   612
            Index           =   0
            Left            =   -69880
            TabIndex        =   94
            Top             =   480
            Visible         =   0   'False
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Estado de la Persona"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton bntConfig_Exclusiones 
            Height          =   612
            Index           =   1
            Left            =   -69880
            TabIndex        =   95
            Top             =   1200
            Visible         =   0   'False
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Líneas de Crédito"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton bntConfig_Exclusiones 
            Height          =   612
            Index           =   2
            Left            =   -69880
            TabIndex        =   96
            Top             =   1920
            Visible         =   0   'False
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Conceptos de CxC"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton bntConfig_Exclusiones 
            Height          =   612
            Index           =   3
            Left            =   -69880
            TabIndex        =   97
            Top             =   2640
            Visible         =   0   'False
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Empresas e Instituciones"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtCabys 
            Height          =   312
            Left            =   7800
            TabIndex        =   99
            Top             =   1440
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtClienteIdProv 
            Height          =   312
            Left            =   3480
            TabIndex        =   51
            Top             =   720
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoId 
            Height          =   312
            Left            =   3480
            TabIndex        =   55
            Top             =   1080
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
            Height          =   312
            Left            =   3480
            TabIndex        =   56
            Top             =   1440
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   2892
            Left            =   -69760
            TabIndex        =   105
            Top             =   1080
            Visible         =   0   'False
            Width           =   12372
            _Version        =   1441793
            _ExtentX        =   21823
            _ExtentY        =   5101
            _StockProps     =   79
            Caption         =   "Reactivación de Movimientos Bloqueados por Politica de Exclusión de Montos Maximos"
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton btnReactivacion 
               Height          =   612
               Left            =   8280
               TabIndex        =   106
               Top             =   480
               Width           =   2172
               _Version        =   1441793
               _ExtentX        =   3831
               _ExtentY        =   1080
               _StockProps     =   79
               Caption         =   "Reactivar Movimientos Excluídos"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.DateTimePicker dtpR_Inicio 
               Height          =   312
               Left            =   5880
               TabIndex        =   107
               Top             =   480
               Width           =   1452
               _Version        =   1441793
               _ExtentX        =   2561
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
            Begin XtremeSuiteControls.DateTimePicker dtpR_Corte 
               Height          =   312
               Left            =   5880
               TabIndex        =   108
               Top             =   840
               Width           =   1452
               _Version        =   1441793
               _ExtentX        =   2561
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   34
               Left            =   3360
               TabIndex        =   110
               Top             =   480
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Mov. Fecha Inicio"
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
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   252
               Index           =   33
               Left            =   3360
               TabIndex        =   109
               Top             =   840
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Mov. Fecha Corte"
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
            End
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   30
            Left            =   5160
            TabIndex        =   100
            Top             =   1440
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cabys Intereses:"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   72
            Top             =   720
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "ID Cliente [Proveedor]"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   71
            Top             =   1080
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo Id"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   2
            Left            =   840
            TabIndex        =   70
            Top             =   1440
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Identificación"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   3
            Left            =   840
            TabIndex        =   69
            Top             =   1800
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Razón Social"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   4
            Left            =   5760
            TabIndex        =   68
            Top             =   2160
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha de Inicio"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   5
            Left            =   840
            TabIndex        =   67
            Top             =   3720
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email Default"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   28
            Left            =   5760
            TabIndex        =   66
            Top             =   2520
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Método de Facturación"
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
         End
      End
      Begin XtremeSuiteControls.TabControl tcFiltros 
         Height          =   5172
         Left            =   -69880
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   12492
         _Version        =   1441793
         _ExtentX        =   22034
         _ExtentY        =   9123
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
         Color           =   128
         ItemCount       =   2
         Item(0).Caption =   "Detalle"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "gbFacturas"
         Item(0).Control(1)=   "gbFactura_Detalle"
         Item(0).Control(2)=   "btnExportar"
         Item(1).Caption =   "Resumen"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "lswConceptos"
         Item(1).Control(1)=   "txtRes_Facturas"
         Item(1).Control(2)=   "Label1(24)"
         Item(1).Control(3)=   "txtRes_Inicio"
         Item(1).Control(4)=   "Label1(25)"
         Item(1).Control(5)=   "txtRes_Corte"
         Item(1).Control(6)=   "Label1(26)"
         Item(1).Control(7)=   "txtRes_Facturado"
         Item(1).Control(8)=   "Label1(27)"
         Item(1).Control(9)=   "btnExportarRsm"
         Begin XtremeSuiteControls.ListView lswConceptos 
            Height          =   4692
            Left            =   -65200
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   7692
            _Version        =   1441793
            _ExtentX        =   13568
            _ExtentY        =   8276
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.GroupBox gbFacturas 
            Height          =   2772
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   12372
            _Version        =   1441793
            _ExtentX        =   21823
            _ExtentY        =   4890
            _StockProps     =   79
            Caption         =   "Facturas"
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.ListView lswFacturas 
               Height          =   2292
               Left            =   360
               TabIndex        =   33
               Top             =   360
               Width           =   12012
               _Version        =   1441793
               _ExtentX        =   21188
               _ExtentY        =   4043
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
               Appearance      =   16
            End
         End
         Begin XtremeSuiteControls.GroupBox gbFactura_Detalle 
            Height          =   1812
            Left            =   120
            TabIndex        =   34
            Top             =   3240
            Width           =   12372
            _Version        =   1441793
            _ExtentX        =   21823
            _ExtentY        =   3196
            _StockProps     =   79
            Caption         =   "Detalle: "
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.ListView lswFacturaDetalle 
               Height          =   1452
               Left            =   360
               TabIndex        =   35
               Top             =   360
               Width           =   12012
               _Version        =   1441793
               _ExtentX        =   21188
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
               FullRowSelect   =   -1  'True
               Appearance      =   16
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtRes_Facturas 
            Height          =   312
            Left            =   -67960
            TabIndex        =   37
            Top             =   600
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRes_Inicio 
            Height          =   312
            Left            =   -67960
            TabIndex        =   39
            Top             =   960
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRes_Corte 
            Height          =   312
            Left            =   -67960
            TabIndex        =   41
            Top             =   1320
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRes_Facturado 
            Height          =   312
            Left            =   -67960
            TabIndex        =   43
            Top             =   1800
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   312
            Left            =   11280
            TabIndex        =   45
            Top             =   0
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Exportar"
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
            Transparent     =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnExportarRsm 
            Height          =   312
            Left            =   -67120
            TabIndex        =   46
            Top             =   2280
            Visible         =   0   'False
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Exportar"
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
            Transparent     =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   27
            Left            =   -69880
            TabIndex        =   44
            Top             =   1800
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monto Facturado"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   26
            Left            =   -69880
            TabIndex        =   42
            Top             =   1320
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Corte"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   25
            Left            =   -69880
            TabIndex        =   40
            Top             =   960
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Inicio"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   24
            Left            =   -69880
            TabIndex        =   38
            Top             =   600
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. Facturas"
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
            Alignment       =   4
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCortes 
         Height          =   3852
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   12372
         _Version        =   1441793
         _ExtentX        =   21823
         _ExtentY        =   6794
         _StockProps     =   79
         Caption         =   "Cortes realizados"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswCortes 
            Height          =   3372
            Left            =   1320
            TabIndex        =   3
            Top             =   360
            Width           =   11052
            _Version        =   1441793
            _ExtentX        =   19494
            _ExtentY        =   5948
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
            Appearance      =   16
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2892
         Left            =   240
         TabIndex        =   2
         Top             =   4560
         Width           =   12372
         _Version        =   1441793
         _ExtentX        =   21823
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Cortes realizados"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnCorte 
            Height          =   492
            Left            =   10800
            TabIndex        =   7
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Registrar Corte"
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
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   312
            Left            =   2760
            TabIndex        =   9
            Top             =   480
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.DateTimePicker dtpFacturacion 
            Height          =   312
            Left            =   2760
            TabIndex        =   47
            Top             =   840
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   29
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha de Facturación"
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
         End
         Begin XtremeSuiteControls.Label lblCorteEstado 
            Height          =   492
            Left            =   5640
            TabIndex        =   30
            Top             =   360
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "[Estado del Corte]"
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
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   14
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha de Corte"
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
         End
      End
      Begin XtremeSuiteControls.GroupBox gbFacturas_Filtros 
         Height          =   1212
         Left            =   -69880
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   12612
         _Version        =   1441793
         _ExtentX        =   22246
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Filtros: "
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtFiltro_RazonSocial 
            Height          =   312
            Left            =   2160
            TabIndex        =   15
            Top             =   600
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFiltro_Factura 
            Height          =   312
            Left            =   6120
            TabIndex        =   18
            Top             =   600
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFiltro_Inicio 
            Height          =   312
            Left            =   7680
            TabIndex        =   20
            Top             =   600
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
         Begin XtremeSuiteControls.DateTimePicker dtpFiltro_Corte 
            Height          =   312
            Left            =   9000
            TabIndex        =   21
            Top             =   600
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
         Begin XtremeSuiteControls.PushButton btnFacturas 
            Height          =   312
            Left            =   11640
            TabIndex        =   26
            Top             =   600
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ComboBox cboFiltro_Estado 
            Height          =   312
            Left            =   10320
            TabIndex        =   25
            Top             =   600
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2143
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtFiltro_Id 
            Height          =   312
            Left            =   600
            TabIndex        =   14
            Top             =   600
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   22
            Left            =   10320
            TabIndex        =   24
            Top             =   360
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Estado"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   21
            Left            =   9000
            TabIndex        =   23
            Top             =   360
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Corte"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   20
            Left            =   7680
            TabIndex        =   22
            Top             =   360
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Inicio"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   19
            Left            =   6120
            TabIndex        =   19
            Top             =   360
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Factura"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   18
            Left            =   2160
            TabIndex        =   17
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Nombre, Razón Social"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   17
            Left            =   600
            TabIndex        =   16
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Identificación"
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
            Alignment       =   4
         End
      End
      Begin XtremeSuiteControls.GroupBox gbClientes_Filtros 
         Height          =   1092
         Left            =   -69760
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   12372
         _Version        =   1441793
         _ExtentX        =   21823
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Filtros: "
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtFiltro_CL_Nombre 
            Height          =   312
            Left            =   2880
            TabIndex        =   11
            Top             =   480
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCliente 
            Height          =   312
            Left            =   9120
            TabIndex        =   27
            Top             =   480
            Width           =   612
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtFiltro_CL_Id 
            Height          =   312
            Left            =   960
            TabIndex        =   10
            Top             =   480
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   16
            Left            =   2880
            TabIndex        =   13
            Top             =   240
            Width           =   2412
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Nombre, Razón Social"
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
            Alignment       =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   15
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Identificación"
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
            Alignment       =   4
         End
      End
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   312
      Left            =   2280
      TabIndex        =   28
      Top             =   360
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   23
      Left            =   2280
      TabIndex        =   29
      Top             =   120
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cliente:"
      ForeColor       =   16777215
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
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13452
   End
End
Attribute VB_Name = "frmSYS_FacturaEletronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim vPaso As Boolean, vTipo As String

Private Sub bntConfig_Exclusiones_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

Select Case Index
    Case 0 'Estado de la Persona
        vTipo = "ESTP"
    Case 1 'Lineas de Credito
        vTipo = "CRD"
    Case 2 'Conceptos de Cuentas por Cobrar
        vTipo = "CxC"
    Case 3 'Instituciones
        vTipo = "INST"
End Select

strSQL = "exec spSYS_FE_PARAMETROS_Exclusion_Consulta '" & cboCliente.ItemData(cboCliente.ListIndex) & "','" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)

lswExclusiones.ListItems.Clear

vPaso = True

Do While Not rs.EOF
    Set itmX = lswExclusiones.ListItems.Add(, , rs!Codigo)
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = rs!registro_Fecha & ""
        itmX.SubItems(3) = rs!Registro_Usuario & ""
        
        itmX.Checked = IIf(rs!Asignado = 1, True, False)
    rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConfig_Clientes_Sinc_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vClienteId As Long


On Error GoTo vError

Me.MousePointer = vbHourglass

vClienteId = cboCliente.ItemData(cboCliente.ListIndex)

strSQL = "delete SYS_FE_CLIENTES where cod_cliente = '" & vClienteId & "'"
Call ConectionExecute(strSQL)

strSQL = " select [ID] AS 'CLIENTE_ID', TIPO_CLIENTE, CODIGO, CEDULA, RAZON_SOCIAL AS 'NOMBRE', EMAIL1, EMAIL2 " _
       & " From IW_CLIENTE" _
       & " Where ID_CLIENTE_ORIGEN = " & vClienteId _
       & " ORDER BY [ID]"
rs.Open strSQL, db, adOpenStatic

strSQL = ""
Do While Not rs.EOF
  strSQL = strSQL & Space(10) & "INSERT SYS_FE_CLIENTES (COD_CLIENTE, CEDULA, NOMBRE, CLIENTE_ID, CLIENTE_ID_FE, REGISTRO_FECHA, REGISTRO_USUARIO)" _
        & " VALUES('" & vClienteId & "','" & rs!Cedula & "','" & rs!Nombre & "'," & rs!Codigo & "," & rs!Cliente_ID _
        & ", getdate(), '" & glogon.Usuario & "')"
  If Len(strSQL) > 20000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If
  rs.MoveNext
Loop
rs.Close

'Ultimo Lote
If Len(strSQL) > 0 Then
  Call ConectionExecute(strSQL)
  strSQL = ""
End If

Me.MousePointer = vbDefault

MsgBox "Sincronización de Clientes versus Proveedor de Facturación Electrónica realizado satisfactoriamente!", vbInformation

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConfig_Main_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSYS_FE_PARAMETROS_Registra '" & txtClienteIdProv.Text & "','" & txtTipoId.Text & "','" & txtIdentificacion.Text & "','" & txtRazonSocial.Text _
      & "'," & chkActiva.Value & ",'" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") & " 23:59:59','" & txtEmail.Text & "'," & chkEmailDefault.Value & "," & chkNotificaCliente.Value _
      & "," & txtConsecFE.Text & "," & txtConsecNC.Text & "," & txtConsecND.Text & "," & txtConsecTE.Text _
      & ",'" & cboPortal.Text & "','" & txtPortal_Server.Text & "','" & txtPortal_DB.Text & "','" & txtPortal_User.Text & "','" & txtPortal_Key.Text _
      & "','A','" & glogon.Usuario & "','" & Mid(cboMetodo.Text, 1, 1) & "'," & chkIncluyePolizas.Value & "," & chkIncluyePrincipal.Value _
      & "," & chkMontoMaximo.Value & "," & CCur(txtMontoMaximo.Text) & ",'" & Trim(txtCabys.Text) _
      & "','" & txtSucursal.Text & "','" & txtTerminal.Text & "'"

Call ConectionExecute(strSQL)

tcMain.Item(0).Selected = True

vPaso = True
    strSQL = "select rtrim(COD_CLIENTE) as 'Idx', rtrim(RAZON_SOCIAL) as 'itmX' from SYS_FE_PARAMETROS"
    Call sbCbo_Llena_New(cboCliente, strSQL, False, True)
    
    If cboCliente.ListCount = 0 Then
        tcMain.Item(3).Selected = True
    End If
vPaso = False

Call cboCliente_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCortes_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

strSQL = "select * from SYS_FE_CLIENTE_CORTES where cod_cliente = '" & cboCliente.ItemData(cboCliente.ListIndex) _
       & "' order by Corte desc"
Call OpenRecordSet(rs, strSQL)

lswCortes.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswCortes.ListItems.Add(, , rs!Corte_ID)
      itmX.SubItems(1) = Format(rs!Corte, "dd/mm/yyyy")
      itmX.SubItems(2) = Format(rs!Facturacion, "dd/mm/yyyy")
          
    If rs!Metodo_Base = "D" Then
          itmX.SubItems(3) = "Devengado"
    Else
          itmX.SubItems(3) = "Efectivo"
    End If
      
      itmX.SubItems(4) = rs!Registro_Usuario & ""
      itmX.SubItems(5) = rs!registro_Fecha & ""

  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCorte_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rsRes As New ADODB.Recordset
Dim sResult As String, i As Long, iTotal As Long

Dim pFactura As String, pIPolizas As Boolean, pIPrincipal As Boolean, pCabys As String

Dim pActividad As String, pMoneda As String

Dim CodPais As String, FechaTransac As Date, idEmpresa As String, pCedula As String _
    , codSucursal As String, TerminalPOS As String, ComprobanteInterno As String _
    , SituacionComprobante As String, TipoComprobante As String
Dim pClave50 As String, pClave20 As String, pFecha As Date, pLinea As Integer

Dim pTotalGravado As Currency, pTotalExento As Currency, pImpuesto As Currency, pDescuento As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass

pFecha = fxFechaServidor

CodPais = "CRC"
codSucursal = "2"
TerminalPOS = "00001"
SituacionComprobante = "1"



lblCorteEstado.Caption = "Cargando parámetros..."

strSQL = "select INCLUYE_POLIZAS, INCLUYE_PRINCIPAL, CABYS, ACTIVIDAD_ECONOMICA, MONEDA" _
       & ", SUCURSAL, TERMINAL " _
       & " from SYS_FE_PARAMETROS"
Call OpenRecordSet(rs, strSQL)
    pIPolizas = IIf((rs!Incluye_Polizas = 1), True, False)
    pIPrincipal = IIf((rs!Incluye_Principal = 1), True, False)
    pCabys = Trim(rs!Cabys & "")
    pActividad = Trim(rs!Actividad_Economica)
    pMoneda = Trim(rs!Moneda)
    
    If Not IsNull(rs!Sucursal) Then
        codSucursal = Trim(rs!Sucursal)
    End If
    
    If Not IsNull(rs!Terminal) Then
        TerminalPOS = Trim(rs!Terminal)
    End If
rs.Close



''Valida Actividad
'If Len(pActividad) < 5 Then
'   Me.MousePointer = vbDefault
'   lblCorteEstado.Caption = ""
'   MsgBox "Definir la Actividad Economica!", vbCritical
'   Exit Sub
'End If

lblCorteEstado.Caption = "Paso No.1 : Creando BD para Facturar + Clientes"


'strSQL = "exec spCrd_Facturacion_Corte_Reenvio '" & cboCliente.ItemData(cboCliente.ListIndex) _
'        & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" & glogon.Usuario _
'        & "','" & Format(dtpFacturacion.Value, "yyyy/mm/dd") & "'"


strSQL = "exec spCrd_Facturacion_Corte '" & cboCliente.ItemData(cboCliente.ListIndex) _
        & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" & glogon.Usuario _
        & "','" & Format(dtpFacturacion.Value, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)


lblCorteEstado.Caption = "Paso No.2 : Registrando Clientes Nuevos"

sResult = ""

strSQL = "exec spCrd_Facturacion_Notifica_Clientes '" & cboCliente.ItemData(cboCliente.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    i = 0
    iTotal = rs.RecordCount
End If

Do While Not rs.EOF
  
 strSQL = "exec sp_IW_CLIENTEInsert_ProGrX '" & rs!COD_CLIENTE & "', '" & rs!Tipo_Id & "', '" & rs!Cliente_ID & "','" & Trim(rs!Cedula) _
        & "','" & rs!Nombre & "', '" & rs!Email & "', '', '', ''" _
        & ",'" & rs!DIRECCION & "',1,1,1,1, '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "', 30, 1"
 rsRes.Open strSQL, db, adOpenStatic
    sResult = sResult & Space(10) & "exec spCrd_Facturacion_Notifica_Clientes_Result '" & rs!COD_CLIENTE & "','" & Trim(rs!Cedula) & "'," & rsRes!CODIGO_INTERNO
 rsRes.Close
  i = i + 1
  lblCorteEstado.Caption = "Paso No.2 : Registrando Clientes Nuevos [" & i & ", " & iTotal & "]"
  DoEvents
 If Len(sResult) > 20000 Then
   Call ConectionExecute(sResult)
   sResult = ""
 End If
  
 rs.MoveNext
Loop
rs.Close

'Lote Final
If Len(sResult) > 0 Then
  Call ConectionExecute(sResult)
  sResult = ""
End If


lblCorteEstado.Caption = "Paso No.3 : Creando Facturas"

sResult = ""

strSQL = "exec spCrd_Facturacion_Notifica '" & cboCliente.ItemData(cboCliente.ListIndex) & "'"


Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    i = 0
    iTotal = rs.RecordCount
End If



TipoComprobante = "01" 'Factura Electronica

idEmpresa = cboCliente.ItemData(cboCliente.ListIndex)

strSQL = "select cedula from F_Cliente where [ID] = '" & idEmpresa & "'"
With glogon.Recordset
  .Open strSQL, db, adOpenStatic
  pCedula = Trim(!Cedula)
  .Close
End With

Do While Not rs.EOF
'CEDULA, NOMBRE, EMAIL, CORTE, FAC_NUMERO, INTCOR, INTMOR, CARGOS AS 'CARGOS', CLIENTE_ID, CLIENTE_ID_FE
'        , @Email as 'EMAIL_DEFAULT', @EmailApl as 'EMAIL_DEFAULT_APL', @EmailCliente as 'EMAIL_CLIENTE_NO'
  
  ComprobanteInterno = rs!FAC_NUMERO
  'Temporal los 3 dia de retraso
  'FechaTransac = DateAdd("d", -3, rs!Corte)
  
'  FechaTransac = rs!Corte
  FechaTransac = dtpFacturacion.Value
  
  pClave50 = fxHacienda_Clave50("506", FechaTransac, pCedula, codSucursal, TerminalPOS, ComprobanteInterno, SituacionComprobante, TipoComprobante)
  pClave20 = fxHacienda_Clave20(codSucursal, TerminalPOS, ComprobanteInterno, TipoComprobante)
  
  
 'Servicios
 pTotalGravado = 0
 pTotalExento = rs!IntCor + rs!IntMor + rs!Cargos ' + rs!Poliza + rs!Principal
 
 If pIPolizas Then
  pTotalExento = pTotalExento + rs!Poliza
 End If
 
 If pIPrincipal Then
  pTotalExento = pTotalExento + rs!Principal
 End If
 
 
 pImpuesto = 0
 pDescuento = 0
 pLinea = 0
 
 TipoComprobante = Format(TipoComprobante, "00")
 
 'Moneda= 'CRC' , Condicion de Venta = '02'
 
'(@ID_CLIENTE_ORIGEN INT, @CLAVE VARCHAR(50), @NUMERO_CONSECUTIVO VARCHAR(20),
'@COD_CLIENTE VARCHAR(5), @COD_MONEDA VARCHAR(4), @COD_SUCURSAL VARCHAR(3),
'@COD_CONDICION_VENTA VARCHAR(2), @CONSECUTIVO_INTERNO VARCHAR(10), @DIAS_CREDITO SMALLINT,
'@EMAIL_CLIENTE VARCHAR(150), @ENVIAR_CLIENTE BIT, @FECHA_EMISION DATETIME,
'@EXTRANJERO BIT, @MEDIOPAGO1 VARCHAR(2),
'@MEDIOPAGO2 VARCHAR(2), @MEDIOPAGO3 VARCHAR(2), @MEDIOPAGO4 VARCHAR(2), @REF_CODIGO VARCHAR(2),
'@REF_FECHAEMISION DATETIME, @REF_NUMERO_CLAVE VARCHAR(50),  @REF_RAZON VARCHAR(180), @REF_TIPO_DOCUMENTO VARCHAR(2),
'@SITUACION VARCHAR(1), @TERMINAL VARCHAR(5), @TIPO_CAMBIO DECIMAL(8, 4), @TIPO_DOCUMENTO VARCHAR(2),@TIPO_IDENTIFICACION VARCHAR(2),

'@TOTAL_SERVICIOS_GRAVADOS DECIMAL(18, 5), @TOTAL_SERVICIOS_EXENTOS DECIMAL(18, 5), @TOTAL_SERVICIOS_EXONERADOS DECIMAL(18, 5),--
'@TOTAL_MERCADERIA_GRAVADA DECIMAL(18, 5), @TOTAL_MERCANCIA_EXENTA DECIMAL(18, 5), @TOTAL_MERCANCIA_EXONERADA DECIMAL(18, 5),--
'@TOTAL_GRAVADO DECIMAL(18, 5), @TOTAL_EXENTO DECIMAL(18, 5), @TOTAL_EXONERADO DECIMAL(18, 5),--
'@TOTAL_VENTA DECIMAL(18, 5), @TOTAL_DESCUENTOS DECIMAL(18, 5), @TOTAL_VENTA_NETA DECIMAL(18, 5),
'@TOTAL_IMPUESTOS DECIMAL(18, 5), @TOTAL_IVA_DEVUELTO DECIMAL(18, 5),--
'@TOTAL_OTROS_CARGOS DECIMAL(18, 5),-- @TOTAL_COMPROBANTE DECIMAL(18, 5),
'@ESTADO_HACIENDA VARCHAR(10), @US_INGRESA VARCHAR(30), @FECHA__INGRESA DATETIME,
'@NUM_RESOLUCION VARCHAR(50), @FECHA_SOLUCION VARCHAR(20), @COMENTARIOS VARCHAR(800)='', @COD_ACTIVIDAD VARCHAR(6))

 strSQL = "exec sp_IW_ENC_FACTURAInsert " & idEmpresa & ",'" & pClave50 & "','" & pClave20 & "','" & rs!Cliente_ID _
         & "','" & Trim(rs!Moneda) & "','" & codSucursal & "','02','" & ComprobanteInterno & "',30,'"
 If rs!EMAIL_DEFAULT_APL = 1 Then
        strSQL = strSQL & rs!EMAIL_DEFAULT & "',"
 Else
        strSQL = strSQL & rs!Email & "',"
 End If
 
 strSQL = strSQL & rs!EMAIL_CLIENTE_NO & ",'" & Format(FechaTransac, "yyyy/mm/dd hh:mm:ss") & "',0,'05','','',''" _
        & ",'', Null, '', '', ''" _
        & ",'1','" & TerminalPOS & "', " & rs!Tipo_Cambio & ", '" & TipoComprobante & "','" & Format(rs!Tipo_Id, "00") _
        & "'," & pTotalGravado & ", " & pTotalExento & ", 0" _
        & ", " & pTotalGravado & ", " & 0 & ", 0" _
        & ", " & pTotalGravado & ", " & pTotalExento & ", 0" _
        & ", " & pTotalGravado + pTotalExento & ", " & pDescuento _
        & ", " & pTotalGravado + pTotalExento - pDescuento & "," & pImpuesto & ", 0, 0" _
        & ", " & (pTotalGravado + pTotalExento - pDescuento + pImpuesto) _
        & ", 1,'" & glogon.Usuario & "','" & Format(pFecha, "yyyy/mm/dd hh:mm:ss") _
        & "','','','FACT. PROGRX','" & pActividad & "'"

 rsRes.Open strSQL, db, adOpenStatic
   pFactura = rsRes!id_Factura
 rsRes.Close



  '--Factura (Detalle)

 If rs!IntCor > 0 Then
    pLinea = pLinea + 1
    strSQL = "exec sp_IW_DET_FACTURAInsert " & idEmpresa & "," & pFactura & "," & pLinea _
           & ",'CRD001',1, 'I','INTERES CORRIENTE DEL MES'," & rs!IntCor & "," & rs!IntCor _
           & ",0, 'DESCUENTO CLIENTES'," & rs!IntCor & ", '01', 0, 0" _
           & ", Null, Null, Null, Null, Null, Null, Null,'01','" & pClave50 & "'"
    db.Execute strSQL
  End If

 If rs!IntMor > 0 Then
    pLinea = pLinea + 1
    strSQL = "exec sp_IW_DET_FACTURAInsert " & idEmpresa & "," & pFactura & "," & pLinea _
           & ",'CRD002',1, 'I','INTERES ATRASADOS'," & rs!IntMor & "," & rs!IntMor _
           & ",0, 'DESCUENTO CLIENTES'," & rs!IntMor & ", '01', 0, 0" _
           & ", Null, Null, Null, Null, Null, Null, Null,'01','" & pClave50 & "'"
    db.Execute strSQL
  End If

 If rs!Cargos > 0 Then
    pLinea = pLinea + 1
    strSQL = "exec sp_IW_DET_FACTURAInsert " & idEmpresa & "," & pFactura & "," & pLinea _
           & ",'CRD003',1, 'I','CARGOS ADM Y DE FORMALIZACION'," & rs!Cargos & "," & rs!Cargos _
           & ",0, 'DESCUENTO CLIENTES'," & rs!Cargos & ", '01', 0, 0" _
           & ", Null, Null, Null, Null, Null, Null, Null,'01','" & pClave50 & "'"
    db.Execute strSQL
  End If

 If rs!Poliza > 0 And pIPolizas Then
    pLinea = pLinea + 1
    strSQL = "exec sp_IW_DET_FACTURAInsert " & idEmpresa & "," & pFactura & "," & pLinea _
           & ",'CRD004',1, 'Unid','POLIZAS DEL CREDITO'," & rs!Poliza & "," & rs!Poliza _
           & ",0, 'DESCUENTO CLIENTES'," & rs!Poliza & ", '01', 0, 0" _
           & ", Null, Null, Null, Null, Null, Null, Null,'01','" & pClave50 & "'"
    db.Execute strSQL
  End If

 If rs!Principal > 0 And pIPrincipal Then
    pLinea = pLinea + 1
    strSQL = "exec sp_IW_DET_FACTURAInsert " & idEmpresa & "," & pFactura & "," & pLinea _
           & ",'CRD005',1, 'Unid','ABONO AL CREDITO'," & rs!Principal & "," & rs!Principal _
           & ",0, 'DESCUENTO CLIENTES'," & rs!Principal & ", '01', 0, 0" _
           & ", Null, Null, Null, Null, Null, Null, Null,'01','" & pClave50 & "'"
    db.Execute strSQL
  End If

  '--Factura Procesada: ProGrX
 
  sResult = sResult & Space(10) & "exec spCrd_Facturacion_Notifica_Result '" & cboCliente.ItemData(cboCliente.ListIndex) _
         & "','" & rs!FAC_NUMERO & "','" & pFactura & "','" & glogon.Usuario & "'"
  i = i + 1
  lblCorteEstado.Caption = "Paso No.3 : Registrando Facturas [" & i & ", " & iTotal & "]"
  DoEvents
  
 If Len(sResult) > 20000 Then
   Call ConectionExecute(sResult)
   sResult = ""
 End If
  
 rs.MoveNext
Loop
rs.Close

'Lote Final
If Len(sResult) > 0 Then
  Call ConectionExecute(sResult)
  sResult = ""
End If


lblCorteEstado.Caption = ""


Me.MousePointer = vbDefault

MsgBox "Proceso de Corte + Facturación realizado satisfactoriamente!", vbInformation

Call sbCortes_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub







Public Sub FE_Portal_Access()
Dim strSQL As String, vPaso As Boolean

On Error GoTo pConPortalError

Screen.MousePointer = vbHourglass


strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & RTrim(txtPortal_Server.Text) _
       & ";Database=" & RTrim(txtPortal_DB.Text) & ";APP=PGX_Facturacion;tcp:" & RTrim(txtPortal_Server.Text) _
       & "," & SIFGlobal.PuertosDisponibles & ";"


db.Close

vPaso = False
 
  
Conexion_Portal_Inicial:
  
With db
  
  vPaso = True
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open strSQL, RTrim(txtPortal_User.Text), RTrim(txtPortal_Key.Text)
  .CommandTimeout = 360
End With

Screen.MousePointer = vbDefault
Exit Sub

pConPortalError:
  If Not vPaso Then GoTo Conexion_Portal_Inicial
  
  Screen.MousePointer = vbDefault
  MsgBox "No se tiene Conexión con el Servidor de Facturación!", vbCritical, "Contacte a su Administrador"

End Sub


Private Sub btnExportar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lswFacturas)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnExportarRsm_Click()
 
On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lswConceptos)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnFacturas_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

lswFacturas.ListItems.Clear
lswFacturaDetalle.ListItems.Clear

lswConceptos.ListItems.Clear

'Detalle

tcFiltros.Item(0).Selected = True

strSQL = "exec spProGrX_Facturas_Consulta '" & cboCliente.ItemData(cboCliente.ListIndex) & "','" & txtFiltro_Factura.Text _
        & "','" & txtFiltro_Id.Text & "','" & txtFiltro_RazonSocial.Text & "','" _
        & Format(dtpFiltro_Inicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpFiltro_Corte.Value, "yyyy/mm/dd") _
        & " 23:59:59','" & Mid(cboFiltro_Estado.Text, 1, 1) & "'"
rs.Open strSQL, db, adOpenStatic

Do While Not rs.EOF
  If rs!Tipo_Documento = "01" Then
      Set itmX = lswFacturas.ListItems.Add(, , "FE")
  Else
      Set itmX = lswFacturas.ListItems.Add(, , "NC")
  End If
      
      itmX.SubItems(1) = "_" & rs!Numero_Consecutivo
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Razon_Social
      itmX.SubItems(4) = rs!Fecha_Emision
      itmX.SubItems(5) = Format(rs!Total_Venta, "Standard")
      itmX.SubItems(6) = Format(rs!Total_Exento, "Standard")
      itmX.SubItems(7) = Format(rs!Total_Gravado, "Standard")
      itmX.SubItems(8) = Format(rs!Total_Impuestos, "Standard")
      itmX.SubItems(9) = Format(rs!Total_Descuentos, "Standard")
      itmX.SubItems(10) = Format(rs!Total_Comprobante, "Standard")
      itmX.SubItems(11) = "_" & rs!Clave
      itmX.SubItems(12) = rs!XML_Respuesta & ""
      itmX.SubItems(13) = rs!Observaciones & ""
      
      itmX.Tag = rs!id_Factura
  
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

vPaso = False


'Resumen
strSQL = "exec spProGrX_Facturas_Consulta_Rsm '" & cboCliente.ItemData(cboCliente.ListIndex) & "','" & txtFiltro_Factura.Text _
        & "','" & txtFiltro_Id.Text & "','" & txtFiltro_RazonSocial.Text & "','" _
        & Format(dtpFiltro_Inicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpFiltro_Corte.Value, "yyyy/mm/dd") _
        & " 23:59:59','" & Mid(cboFiltro_Estado.Text, 1, 1) & "','R'"
rs.Open strSQL, db, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   txtRes_Inicio.Text = rs!Inicio
   txtRes_Corte.Text = rs!Corte
   txtRes_Facturado.Text = Format(rs!Total_Venta, "Standard")
   txtRes_Facturas.Text = Format(rs!Facturas, "###,###,##0")

Else

   txtRes_Inicio.Text = ""
   txtRes_Corte.Text = ""
   txtRes_Facturado.Text = "0.00"
   txtRes_Facturas.Text = "0"
   
End If
rs.Close

'Resumen Conceptos
strSQL = "exec spProGrX_Facturas_Consulta_Rsm '" & cboCliente.ItemData(cboCliente.ListIndex) & "','" & txtFiltro_Factura.Text _
        & "','" & txtFiltro_Id.Text & "','" & txtFiltro_RazonSocial.Text & "','" _
        & Format(dtpFiltro_Inicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpFiltro_Corte.Value, "yyyy/mm/dd") _
        & " 23:59:59','" & Mid(cboFiltro_Estado.Text, 1, 1) & "','D'"
rs.Open strSQL, db, adOpenStatic

Do While Not rs.EOF
  If rs!Tipo_Documento = "01" Then
      Set itmX = lswConceptos.ListItems.Add(, , "FE")
  Else
      Set itmX = lswConceptos.ListItems.Add(, , "NC")
  End If
      itmX.SubItems(1) = rs!Lineas
      itmX.SubItems(2) = rs!Detalle
      itmX.SubItems(3) = Format(rs!Total_Venta, "Standard")
      itmX.SubItems(4) = rs!Inicio
      itmX.SubItems(5) = rs!Corte
      itmX.SubItems(6) = rs!XML_Respuesta & ""
      
  rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnReactivacion_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Operacion_Factura_Reactivar '" & Format(dtpR_Inicio.Value, "yyyy-mm-dd") _
       & "','" & Format(dtpR_Corte.Value, "yyyy-mm-dd") & " 23:59:59'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Reactivación de Casos Excluidos por Montos, realizado satisfactoriamente!", vbInformation

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboCliente_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


tcMain.Item(0).Selected = True


If vPaso Or cboCliente.ListCount = 0 Then Exit Sub

strSQL = "select P.*, isnull(C.descripcion,'') as 'Cabys_Desc' " _
       & " from sys_FE_Parametros P left join vINV_Cabys C on P.Cabys = C.COD_BYS" _
       & " where P.cod_Cliente = '" & cboCliente.ItemData(cboCliente.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

txtClienteIdProv.Text = rs!COD_CLIENTE
txtTipoId.Text = rs!Tipo_Id
txtIdentificacion.Text = rs!Cedula

txtRazonSocial.Text = rs!Razon_Social
chkActiva.Value = rs!Activo

dtpFechaInicio.Value = rs!FECHA_INICIO
txtEmail.Text = rs!NOTIFICA_EMAIL

chkEmailDefault.Value = rs!NOTIFICA_EMAIL_ACTIVO
chkNotificaCliente.Value = rs!NOTIFICA_CLIENTE

txtConsecFE.Text = rs!CONSECUTIVO_FE
txtConsecNC.Text = rs!CONSECUTIVO_NC
txtConsecND.Text = rs!CONSECUTIVO_ND
txtConsecTE.Text = rs!CONSECUTIVO_TE

cboPortal.Text = RTrim(rs!ACC_CODIGO)

txtPortal_Server.Text = rs!ACC_SERVER
txtPortal_DB.Text = rs!ACC_DB
txtPortal_User.Text = rs!ACC_USR
txtPortal_Key.Text = rs!ACC_KEY

If rs!Metodo_Base = "D" Then
    cboMetodo.Text = "Devengado"
Else
    cboMetodo.Text = "Efectivo"

End If

chkIncluyePolizas.Value = rs!Incluye_Polizas
chkIncluyePrincipal.Value = rs!Incluye_Principal

chkMontoMaximo.Value = rs!Max_Monto_Apl
txtMontoMaximo.Text = Format(rs!Max_Monto, "Standard")

txtCabys.Text = rs!Cabys & ""
txtCabys.ToolTipText = rs!Cabys_Desc & ""

txtSucursal.Text = rs!Sucursal
txtTerminal.Text = rs!Terminal

rs.Close

Call sbCortes_Load

'Estableciendo Conexion
Call FE_Portal_Access


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 10

imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

dtpCorte.Value = fxFechaServidor

dtpR_Inicio.Value = dtpCorte.Value
dtpR_Corte.Value = dtpCorte.Value

cboMetodo.Clear
cboMetodo.AddItem "Devengado"
cboMetodo.AddItem "Efectivo"
cboMetodo.Text = "Devengado"

dtpFechaInicio.Value = dtpCorte.Value
dtpFiltro_Inicio.Value = dtpCorte.Value
dtpFiltro_Corte.Value = dtpCorte.Value

dtpFacturacion.Value = dtpCorte.Value
dtpFacturacion.MaxDate = dtpCorte.Value
dtpFacturacion.MinDate = DateAdd("d", -2, dtpCorte.Value)

cboPortal.AddItem "AVS"
cboPortal.AddItem "BO"
cboPortal.AddItem "GTI"
cboPortal.Text = "AVS"

cboFiltro_Estado.AddItem "TODAS"
cboFiltro_Estado.AddItem "Aceptada"
cboFiltro_Estado.AddItem "Rechazada"
cboFiltro_Estado.Text = "TODAS"

With lswCortes.ColumnHeaders
    .Clear
    .Add , , "Corte Id", 1000
    .Add , , "Corte", 1800, vbCenter
    .Add , , "Facturación", 1800, vbCenter
    .Add , , "Método", 1500, vbCenter
    .Add , , "Reg. Usuario", 2500, vbCenter
    .Add , , "Reg. Fecha", 2500
End With

With lswConceptos.ColumnHeaders
    .Clear
    .Add , , "Tipo", 1000
    .Add , , "Lineas", 1000
    .Add , , "Detalle", 3500
    .Add , , "Facturado", 2000, vbRightJustify
    .Add , , "Fec.Inicio", 2000
    .Add , , "Fec.Corte", 2000
    .Add , , "Estado", 1500, vbCenter
End With



With lswFacturas.ColumnHeaders
    .Clear
    .Add , , "Tipo", 1000
    .Add , , "Comprobante", 2500
    .Add , , "Identificación", 1500
    .Add , , "Razón Social", 3500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Total", 1600, vbRightJustify
    .Add , , "Total Exento", 1600, vbRightJustify
    .Add , , "Total Gravado", 1600, vbRightJustify
    .Add , , "Total Impuesto", 1600, vbRightJustify
    .Add , , "Total Descuento", 1600, vbRightJustify
    .Add , , "Total Comprobante", 1600, vbRightJustify
    .Add , , "No. Referencia", 3500
    .Add , , "Estado", 1500, vbCenter
    .Add , , "Observaciones", 3500
    
End With


With lswFacturaDetalle.ColumnHeaders
    .Clear
    .Add , , "Línea", 900
    .Add , , "Código", 1000, vbCenter
    .Add , , "Producto", 3500, vbCenter
    .Add , , "Precio Ud", 1600, vbRightJustify
    .Add , , "Qty", 1500, vbCenter
    .Add , , "Unidad", 1500, vbCenter
    .Add , , "Total", 1600, vbRightJustify
    .Add , , "Descuento", 1600, vbRightJustify
    .Add , , "Impuesto", 1600, vbRightJustify
    .Add , , "Cabys", 1600, vbCenter
End With

With lswClientes.ColumnHeaders
    .Clear
    .Add , , "Id Prov", 1200
    .Add , , "Tipo Id", 1000, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Razón Social/Nombre", 3500
    .Add , , "Email No.1", 3500
    .Add , , "Email No.2", 3500
    .Add , , "Telefono No.1", 1200
    .Add , , "Telefono No.2", 1200
    .Add , , "Id. Provincia", 1200, vbCenter
    .Add , , "Id. Cantón", 1200, vbCenter
    .Add , , "Id. Distrito", 1200, vbCenter
    .Add , , "Id. Barrio", 1200, vbCenter
    .Add , , "Dirección", 4200
    
End With

With lswExclusiones.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Descripción", 4000
    .Add , , "Fecha", 2800, vbCenter
    .Add , , "Usuario", 2000, vbCenter
End With

lswExclusiones.Checkboxes = True

tcMain.Item(0).Selected = True

vPaso = True
    strSQL = "select rtrim(COD_CLIENTE) as 'Idx', rtrim(RAZON_SOCIAL) as 'itmX' from SYS_FE_PARAMETROS"
    Call sbCbo_Llena_New(cboCliente, strSQL, False, True)
    
    If cboCliente.ListCount = 0 Then
        tcMain.Item(3).Selected = True
    End If
vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)


If btnConfig_Consec.Tag = "0" Then
    txtConsecFE.Locked = True
    txtConsecNC.Locked = True
    txtConsecND.Locked = True
    txtConsecTE.Locked = True
Else
    txtConsecFE.Locked = False
    txtConsecNC.Locked = False
    txtConsecND.Locked = False
    txtConsecTE.Locked = False
End If

If btnConfig_Portal.Tag = "0" Then
    txtPortal_Server.Locked = True
    txtPortal_DB.Locked = True
    txtPortal_User.Locked = True
    txtPortal_Key.Locked = True
    
    cboPortal.Locked = True
    
    txtPortal_Server.PasswordChar = "/"
    txtPortal_DB.PasswordChar = "/"
    txtPortal_User.PasswordChar = "/"
    
    
Else
    txtPortal_Server.Locked = False
    txtPortal_DB.Locked = False
    txtPortal_User.Locked = False
    txtPortal_Key.Locked = False

    cboPortal.Locked = False

    txtPortal_Server.PasswordChar = ""
    txtPortal_DB.PasswordChar = ""
    txtPortal_User.PasswordChar = ""

End If


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lswClientes_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswClientes.SortKey = ColumnHeader.Index - 1
  If lswClientes.SortOrder = 0 Then lswClientes.SortOrder = 1 Else lswClientes.SortOrder = 0
  lswClientes.Sorted = True
End Sub

Private Sub lswCortes_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCortes.SortKey = ColumnHeader.Index - 1
  If lswCortes.SortOrder = 0 Then lswCortes.SortOrder = 1 Else lswCortes.SortOrder = 0
  lswCortes.Sorted = True
End Sub

Private Sub lswExclusiones_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswExclusiones.SortKey = ColumnHeader.Index - 1
  If lswExclusiones.SortOrder = 0 Then lswExclusiones.SortOrder = 1 Else lswExclusiones.SortOrder = 0
  lswExclusiones.Sorted = True
End Sub

Private Sub lswExclusiones_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, pMovimiento As String

On Error GoTo vError

If vPaso Then Exit Sub
                    
If Item.Checked Then
   pMovimiento = "A"
Else
   pMovimiento = "E"
End If

strSQL = "exec spSYS_FE_PARAMETROS_Exclusion '" & cboCliente.ItemData(cboCliente.ListIndex) _
       & "','" & Item.Text & "','" & pMovimiento & "','" & vTipo & "','" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswFacturaDetalle_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswFacturaDetalle.SortKey = ColumnHeader.Index - 1
  If lswFacturaDetalle.SortOrder = 0 Then lswFacturaDetalle.SortOrder = 1 Else lswFacturaDetalle.SortOrder = 0
  lswFacturaDetalle.Sorted = True
End Sub

Private Sub lswFacturas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswFacturas.SortKey = ColumnHeader.Index - 1
  If lswFacturas.SortOrder = 0 Then lswFacturas.SortOrder = 1 Else lswFacturas.SortOrder = 0
  lswFacturas.Sorted = True
End Sub

Private Sub lswFacturas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswFacturaDetalle.ListItems.Clear


strSQL = "exec spProGrX_Factura_Detalle '" & cboCliente.ItemData(cboCliente.ListIndex) & "','" & Item.Tag & "'"
rs.Open strSQL, db, adOpenStatic

Do While Not rs.EOF
  Set itmX = lswFacturaDetalle.ListItems.Add(, , rs!Num_Linea)
      itmX.SubItems(1) = rs!Tipo_Producto
      itmX.SubItems(2) = rs!Detalle
      itmX.SubItems(3) = Format(rs!Precio_Unitario, "Standard")
      itmX.SubItems(4) = rs!Cantidad
      itmX.SubItems(5) = rs!Unidad_Medida
      itmX.SubItems(6) = Format(rs!Monto_Total, "Standard")
      itmX.SubItems(7) = Format(rs!Monto_Descuento, "Standard")
      itmX.SubItems(8) = Format(rs!Monto_Impuesto, "Standard")
      itmX.SubItems(9) = rs!Cabys_Desc & ""
  
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tcConfiguracion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    Call bntConfig_Exclusiones_Click(0)
End If
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 3 Then
    tcConfiguracion.Item(0).Selected = True
End If
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

If cboCliente.ListCount > 0 Then
  Call cboCliente_Click
End If
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
    txtCabys.ToolTipText = gBusquedas.Resultado2
End If
End Sub

Private Sub txtFiltro_CL_Id_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtFiltro_CL_Id.Text)
   
   txtFiltro_CL_Id.Text = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select cedula, nombre from SYS_FE_CLIENTES"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtFiltro_CL_Id.Text = Trim(gBusquedas.Resultado)
   txtFiltro_CL_Nombre.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtFiltro_Id_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtFiltro_Id.Text)
   
   txtFiltro_Id.Text = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select cedula, nombre from SYS_FE_CLIENTES"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtFiltro_Id.Text = Trim(gBusquedas.Resultado)
   txtFiltro_RazonSocial.Text = gBusquedas.Resultado2
   Call btnFacturas_Click
End If

End Sub
