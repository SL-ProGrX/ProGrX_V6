VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Gestor_Inversiones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Gestión de Inversiones"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   15600
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   2400
      Top             =   0
   End
   Begin XtremeSuiteControls.CheckBox chkFechaCompra 
      Height          =   216
      Left            =   4320
      TabIndex        =   35
      Top             =   7200
      Width           =   216
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   9372
      Left            =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   16531
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
      SelectedItem    =   3
      Item(0).Caption =   "Inversiones"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "tcCaso"
      Item(0).Control(2)=   "scCaso"
      Item(0).Control(3)=   "btnExport"
      Item(1).Caption =   "Vencimientos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "btnCupon"
      Item(1).Control(1)=   "dtpFechaCupon(0)"
      Item(1).Control(2)=   "dtpFechaCupon(1)"
      Item(1).Control(3)=   "Label1(18)"
      Item(1).Control(4)=   "cboCuponEstado"
      Item(1).Control(5)=   "Label1(19)"
      Item(1).Control(6)=   "gVencimientos"
      Item(2).Caption =   "Valorización"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Asientos"
      Item(3).ControlCount=   15
      Item(3).Control(0)=   "dtpFechaAsiento(0)"
      Item(3).Control(1)=   "dtpFechaAsiento(1)"
      Item(3).Control(2)=   "Label1(15)"
      Item(3).Control(3)=   "lswAsiento"
      Item(3).Control(4)=   "scAsientos"
      Item(3).Control(5)=   "lswAsientoMain"
      Item(3).Control(6)=   "btnAsientos"
      Item(3).Control(7)=   "txtNumAsiento"
      Item(3).Control(8)=   "Label1(16)"
      Item(3).Control(9)=   "gbAsiento"
      Item(3).Control(10)=   "txtAsientoDetalle"
      Item(3).Control(11)=   "Label1(17)"
      Item(3).Control(12)=   "Label1(20)"
      Item(3).Control(13)=   "txtCuenta"
      Item(3).Control(14)=   "txtCuentaDesc"
      Begin XtremeSuiteControls.ListView lswAsientoMain 
         Height          =   3495
         Left            =   0
         TabIndex        =   65
         Top             =   1320
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswAsiento 
         Height          =   3372
         Left            =   0
         TabIndex        =   57
         Top             =   5280
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnAsientos 
         Height          =   315
         Left            =   10320
         TabIndex        =   66
         Top             =   480
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3012
         Left            =   -70000
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   12252
         _Version        =   524288
         _ExtentX        =   21611
         _ExtentY        =   5313
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   44
         SpreadDesigner  =   "frmIVR_Gestor_Inversiones.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.TabControl tcCaso 
         Height          =   3612
         Left            =   -70000
         TabIndex        =   2
         Top             =   5040
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1441793
         _ExtentX        =   18436
         _ExtentY        =   6371
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
         Item(0).Caption =   "Intereses"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lswFlujos_Ingresos"
         Item(1).Caption =   "Primas y Descuentos"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswFlujos_PyD"
         Item(2).Caption =   "Movimientos"
         Item(2).ControlCount=   3
         Item(2).Control(0)=   "lswFi_Mov"
         Item(2).Control(1)=   "cboFi_Tipo"
         Item(2).Control(2)=   "cboFi_Estado"
         Item(3).Caption =   "Cupones"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "lswCupones"
         Item(3).Control(1)=   "cboCupones"
         Item(4).Caption =   "Cierres"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "lswCierres"
         Begin XtremeSuiteControls.ListView lswFlujos_Ingresos 
            Height          =   3252
            Left            =   0
            TabIndex        =   44
            Top             =   360
            Width           =   10332
            _Version        =   1441793
            _ExtentX        =   18224
            _ExtentY        =   5736
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
         End
         Begin XtremeSuiteControls.ListView lswFlujos_PyD 
            Height          =   3252
            Left            =   -70000
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   10332
            _Version        =   1441793
            _ExtentX        =   18224
            _ExtentY        =   5736
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
         End
         Begin XtremeSuiteControls.ListView lswFi_Mov 
            Height          =   2892
            Left            =   -70000
            TabIndex        =   46
            Top             =   720
            Visible         =   0   'False
            Width           =   10332
            _Version        =   1441793
            _ExtentX        =   18224
            _ExtentY        =   5101
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
         End
         Begin XtremeSuiteControls.ListView lswCupones 
            Height          =   2892
            Left            =   -70000
            TabIndex        =   47
            Top             =   684
            Visible         =   0   'False
            Width           =   10332
            _Version        =   1441793
            _ExtentX        =   18224
            _ExtentY        =   5101
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
         End
         Begin XtremeSuiteControls.ListView lswCierres 
            Height          =   3252
            Left            =   -70000
            TabIndex        =   51
            Top             =   360
            Visible         =   0   'False
            Width           =   10452
            _Version        =   1441793
            _ExtentX        =   18436
            _ExtentY        =   5736
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
         End
         Begin XtremeSuiteControls.ComboBox cboCupones 
            Height          =   312
            Left            =   -70000
            TabIndex        =   48
            Top             =   360
            Visible         =   0   'False
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3201
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboFi_Tipo 
            Height          =   312
            Left            =   -70000
            TabIndex        =   49
            Top             =   360
            Visible         =   0   'False
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3201
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboFi_Estado 
            Height          =   312
            Left            =   -68200
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2990
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaAsiento 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   54
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaAsiento 
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   55
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.GroupBox gbAsiento 
         Height          =   612
         Left            =   0
         TabIndex        =   59
         Top             =   8640
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   1080
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtA_Debito 
            Height          =   312
            Left            =   2760
            TabIndex        =   60
            Top             =   120
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtA_Credito 
            Height          =   312
            Left            =   4800
            TabIndex        =   61
            Top             =   120
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtA_Diferencia 
            Height          =   312
            Left            =   8160
            TabIndex        =   62
            Top             =   120
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   0
            Left            =   960
            TabIndex        =   64
            Top             =   120
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Débito / Crédito"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   1
            Left            =   6480
            TabIndex        =   63
            Top             =   120
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Diferencia"
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
      Begin XtremeSuiteControls.FlatEdit txtNumAsiento 
         Height          =   315
         Left            =   5640
         TabIndex        =   67
         Top             =   480
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
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
      Begin XtremeSuiteControls.FlatEdit txtAsientoDetalle 
         Height          =   315
         Left            =   8400
         TabIndex        =   69
         Top             =   480
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
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
      Begin XtremeSuiteControls.PushButton btnCupon 
         Height          =   312
         Left            =   -62800
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCupon 
         Height          =   312
         Index           =   0
         Left            =   -68320
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCupon 
         Height          =   312
         Index           =   1
         Left            =   -67120
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.ComboBox cboCuponEstado 
         Height          =   312
         Left            =   -64720
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3201
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread gVencimientos 
         Height          =   3012
         Left            =   -70000
         TabIndex        =   77
         Top             =   960
         Visible         =   0   'False
         Width           =   12252
         _Version        =   524288
         _ExtentX        =   21611
         _ExtentY        =   5313
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         SpreadDesigner  =   "frmIVR_Gestor_Inversiones.frx":1449
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Left            =   -70000
         TabIndex        =   78
         ToolTipText     =   "Exportar a Excel"
         Top             =   4580
         Visible         =   0   'False
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Gestor_Inversiones.frx":1D94
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   80
         Top             =   840
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   315
         Left            =   3840
         TabIndex        =   81
         Top             =   840
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11245
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   19
         Left            =   -65560
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   18
         Left            =   -69880
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha Vence"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   17
         Left            =   7560
         TabIndex        =   70
         Top             =   480
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detalle"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   16
         Left            =   4560
         TabIndex        =   68
         Top             =   480
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "No. Asiento"
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
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scAsientos 
         Height          =   372
         Left            =   0
         TabIndex        =   58
         Top             =   4920
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Asiento"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha Asientos"
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
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scCaso 
         Height          =   372
         Left            =   -70000
         TabIndex        =   3
         Top             =   4560
         Visible         =   0   'False
         Width           =   12012
         _Version        =   1441793
         _ExtentX        =   21188
         _ExtentY        =   656
         _StockProps     =   14
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
   End
   Begin XtremeSuiteControls.ComboBox cboInstrumento 
      Height          =   312
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboEmisor 
      Height          =   312
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboAdministrador 
      Height          =   312
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboClasificacion 
      Height          =   312
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPortafolio 
      Height          =   312
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboRecurso 
      Height          =   312
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboReserva 
      Height          =   312
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   312
      Left            =   1800
      TabIndex        =   18
      Top             =   4440
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtSerie 
      Height          =   312
      Left            =   1800
      TabIndex        =   19
      Top             =   4800
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtISIN 
      Height          =   312
      Left            =   1800
      TabIndex        =   22
      Top             =   5160
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtCupIp 
      Height          =   312
      Left            =   1800
      TabIndex        =   23
      Top             =   5520
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   312
      Left            =   1800
      TabIndex        =   26
      Top             =   6000
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodicidad 
      Height          =   312
      Left            =   1800
      TabIndex        =   27
      Top             =   6360
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboBaseCalculo 
      Height          =   312
      Left            =   1800
      TabIndex        =   28
      Top             =   6720
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFechaCompra 
      Height          =   312
      Index           =   0
      Left            =   1800
      TabIndex        =   29
      Top             =   7200
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.DateTimePicker dtpFechaCompra 
      Height          =   312
      Index           =   1
      Left            =   3000
      TabIndex        =   34
      Top             =   7200
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.DateTimePicker dtpFechaVence 
      Height          =   312
      Index           =   0
      Left            =   1800
      TabIndex        =   36
      Top             =   7560
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.DateTimePicker dtpFechaVence 
      Height          =   312
      Index           =   1
      Left            =   3000
      TabIndex        =   38
      Top             =   7560
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.CheckBox chkFechaVence 
      Height          =   216
      Left            =   4320
      TabIndex        =   39
      Top             =   7560
      Width           =   216
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   1800
      TabIndex        =   40
      Top             =   8880
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   868
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
      Picture         =   "frmIVR_Gestor_Inversiones.frx":2665
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   492
      Left            =   3000
      TabIndex        =   41
      Top             =   8880
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmIVR_Gestor_Inversiones.frx":3083
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   1800
      TabIndex        =   42
      Top             =   8400
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCierre 
      Height          =   312
      Left            =   1800
      TabIndex        =   52
      Top             =   8040
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   14
      Left            =   240
      TabIndex        =   53
      Top             =   8040
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cierre"
      ForeColor       =   16777215
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
      Index           =   6
      Left            =   240
      TabIndex        =   43
      Top             =   8400
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado"
      ForeColor       =   16777215
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
      Index           =   0
      Left            =   240
      TabIndex        =   37
      Top             =   7560
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha Vence"
      ForeColor       =   16777215
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
      Index           =   13
      Left            =   240
      TabIndex        =   33
      Top             =   7200
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha Compra"
      ForeColor       =   16777215
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
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   6720
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Base Cálculo"
      ForeColor       =   16777215
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
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   6360
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Periodicidad"
      ForeColor       =   16777215
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
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   6000
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Divisa"
      ForeColor       =   16777215
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
      Index           =   49
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "ISIN"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   51
      Left            =   240
      TabIndex        =   24
      Top             =   5520
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "CupIp"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   11
      Left            =   240
      TabIndex        =   21
      Top             =   4800
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Serie"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   10
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Operación"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Reserva"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   47
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Recurso"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Portafolio"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Clasificación"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Administrador"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Emisor"
      ForeColor       =   16777215
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Instrumento"
      ForeColor       =   16777215
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
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmIVR_Gestor_Inversiones.frx":3888
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4728
   End
End
Attribute VB_Name = "frmIVR_Gestor_Inversiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbBuscar()

Dim pEstado As String, pTabla As String

On Error GoTo vError

scCaso.Tag = "0"
scCaso.Caption = ""
tcCaso.Item(0).Selected = True
lswFlujos_Ingresos.ListItems.Clear

If cboCierre.Text = "Actual" Then
    pTabla = "vIVR_Inversiones"
Else
    pTabla = "vIVR_Inversiones_Cierre"
End If

Select Case Mid(cboEstado.Text, 1, 1)
    Case "T"
        pEstado = " in('A','L','S')"
    Case Else
        pEstado = " = '" & Mid(cboEstado.Text, 1, 1) & "'"
End Select

strSQL = "select 0, 0, Titulo_Id, substring(Instrumento_IdX , 1,1) as Tipo, Operacion, Serie, Isin, Administrador_Desc, Instrumento_Desc, Emisor_Desc" _
    & ", Categoria_Desc, Valor_Libros, cod_Divisa, Tipo_Cambio, Fecha_Compra, Fecha_Vencimiento, Tasa_Inicial, Precio" _
    & ", Valor_Facial, Valor_Transado, IVA, Comision_Puesto, Comision_BNV, Interes_Acum_Dias, Interes_Acum_Monto" _
    & ", Costo_Neto,PyD_Tipo, PyD_Base, Pyd_Dias_Acumulados, PyD_Monto_Acumulado, isnull(PyD_Saldo, abs(PyD_Base) ) as 'PyD_Saldo'" _
    & ", Dias_Acumulados, Dias_Acum_Monto" _
    & ", Participacion_Numero, Participacion_Valor, Base_Intereses, Periodicidad_Desc " _
    & ", CTA_INVERSION_MASK,CTA_INTERESES_ACUM_COBRAR_MASK, case when PyD_Tipo = 'P' then CTA_PRIMA_MASK else CTA_DESCUENTOS_MASK end  " _
    & ", CTA_INGRESOS_INTERESES_MASK" _
    & ", CONVERT(varchar,Corte,23)  as 'CORTE', Recurso_Desc, Monto_Inversion" _
    & " from " & pTabla _
    & " Where Estado " & pEstado

If cboCierre.Text <> "Actual" Then
    strSQL = strSQL & " AND CORTE = '" & cboCierre.Text & " 23:59:59'"
End If

If Len(txtOperacion.Text) > 0 Then
    strSQL = strSQL & " AND OPERACION Like '%" & txtOperacion.Text & "%'"
End If

If Len(txtSerie.Text) > 0 Then
    strSQL = strSQL & " AND SERIE Like '%" & txtSerie.Text & "%'"
End If

If Len(txtISIN.Text) > 0 Then
    strSQL = strSQL & " AND ISIN Like '%" & txtISIN.Text & "%'"
End If

If Len(txtCupIp.Text) > 0 Then
    strSQL = strSQL & " AND CUPIP Like '%" & txtCupIp.Text & "%'"
End If

If cboInstrumento.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_INSTRUMENTO = '" & cboInstrumento.ItemData(cboInstrumento.ListIndex) & "'"
End If

If cboEmisor.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_EMISIOR = '" & cboEmisor.ItemData(cboEmisor.ListIndex) & "'"
End If

If cboAdministrador.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_ADMINISTRADOR = '" & cboAdministrador.ItemData(cboAdministrador.ListIndex) & "'"
End If

If cboPortafolio.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_PORTAFOLIO = '" & cboPortafolio.ItemData(cboPortafolio.ListIndex) & "'"
End If

If cboRecurso.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_RECURSO = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
End If

If cboReserva.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_RESERVA = '" & cboReserva.ItemData(cboReserva.ListIndex) & "'"
End If


If cboClasificacion.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_CATEGORIA = '" & cboClasificacion.ItemData(cboClasificacion.ListIndex) & "'"
End If




If cboDivisa.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_DIVISA = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
End If


If cboPeriodicidad.Text <> "TODOS" Then
    strSQL = strSQL & " AND COD_PERIODICIDAD = '" & cboPeriodicidad.ItemData(cboPeriodicidad.ListIndex) & "'"
End If


If cboBaseCalculo.Text <> "Todas" Then
    strSQL = strSQL & " AND BASE_INTERESES = '" & cboBaseCalculo.ItemData(cboBaseCalculo.ListIndex) & "'"
End If

If chkFechaCompra.Value = xtpUnchecked Then
    strSQL = strSQL & " AND FECHA_COMPRA between '" & Format(dtpFechaCompra(0).Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpFechaCompra(1).Value, "yyyy/mm/dd") & " 23:59:59'"
End If


If chkFechaVence.Value = xtpUnchecked Then
    strSQL = strSQL & " AND FECHA_VENCIMIENTO between '" & Format(dtpFechaVence(0).Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpFechaVence(1).Value, "yyyy/mm/dd") & " 23:59:59'"
End If


Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:


End Sub


Private Sub btnAsientos_Click()
        
Dim pInicio As String, pCorte As String, pFiltro As String, pDetalle As String, pCuenta As String

pInicio = Format(dtpFechaAsiento(0).Value, "yyyy/mm/dd") & " 00:00:00"
pCorte = Format(dtpFechaAsiento(1).Value, "yyyy/mm/dd") & " 23:59:59"
        
pFiltro = Trim(txtNumAsiento.Text)
pDetalle = Trim(txtAsientoDetalle.Text)
'pCuenta = fxgCntCuentaFormato(False, Trim(txtCuenta.Text), 0)
pCuenta = Replace(txtCuenta, "-", "")
vPaso = True
Call sbIVR_Asientos_Main(lswAsientoMain, pInicio, pCorte, pFiltro, pDetalle, pCuenta)
        
vPaso = False
        
scAsientos.Caption = "Seleccione un asiento"
lswAsiento.ListItems.Clear
        
End Sub

Private Sub btnBuscar_Click()
 Call sbBuscar
End Sub

Private Sub btnCupon_Click()

On Error GoTo vError

strSQL = "select 0, V.TITULO_ID, case when V.MNT_PRINCIPAL = 0 then 'Cupón' else 'Inversión' end as 'Tipo'" _
       & ", V.FECHA_CORTE, V.SEQ_ID, I.COD_DIVISA, V.MNT_INTERES + V.MNT_PRINCIPAL as 'Total'" _
       & ", V.MNT_INTERES, V.MNT_PRINCIPAL" _
       & ", I.OPERACION, I.SERIE, I.ISIN, I.Administrador_Desc, I.Instrumento_Desc" _
       & " from  vIVR_INVERSIONES I inner join IVR_TITULO_CUPONES_PROY V on I.TITULO_ID = V.TITULO_ID" _
       & " where I.ESTADO = 'A' AND V.ESTADO = '" & Mid(cboCuponEstado.Text, 1, 1) & "'" _
       & " and V.FECHA_CORTE between '" & Format(dtpFechaCupon(0).Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpFechaCupon(1).Value, "yyyy/mm/dd") & " 23:59:59'"

vPaso = True


Call sbCargaGrid(gVencimientos, gVencimientos.MaxCols, strSQL)

gVencimientos.MaxRows = gVencimientos.MaxRows - 1

vPaso = False

Exit Sub

vError:

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case tcCaso.SelectedItem
    Case 0 'Flujo Intereses
        Call Excel_Exportar_Lsw(lswFlujos_Ingresos)
    
    Case 1 'Primas/Descuentos
        Call Excel_Exportar_Lsw(lswFlujos_PyD)
    
    Case 2 'Fondos Mov
        Call Excel_Exportar_Lsw(lswFi_Mov)
    
    Case 3 'Cupones
        Call Excel_Exportar_Lsw(lswCupones)
    
    Case 2 'Cierres
        Call Excel_Exportar_Lsw(lswCierres)

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()


 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "..."
    
    vHeaders.Headers(3) = "Titulo Id"
    vHeaders.Headers(4) = "Tipo"
    vHeaders.Headers(5) = "Operacion"
    vHeaders.Headers(6) = "Serie"
    vHeaders.Headers(7) = "ISIN"
    vHeaders.Headers(8) = "Administrador"
    vHeaders.Headers(9) = "Instrumento"
    vHeaders.Headers(10) = "Emisor"
    vHeaders.Headers(11) = "Categoria"
    vHeaders.Headers(12) = "Valor Libros"

    vHeaders.Headers(13) = "Divisa"
    vHeaders.Headers(14) = "Tipo Cambio"
    vHeaders.Headers(15) = "Fecha Compra"
    vHeaders.Headers(16) = "Fecha Vencimiento"
    vHeaders.Headers(17) = "Tasa Inicial"
    vHeaders.Headers(18) = "Precio"
    vHeaders.Headers(19) = "Valor Facial"
    vHeaders.Headers(20) = "Valor Transado"
    vHeaders.Headers(21) = "I.V.A."
    vHeaders.Headers(22) = "Comision ADM"

    vHeaders.Headers(23) = "Comision BNV"
    vHeaders.Headers(24) = "Cpr. Int.Acum. Días"
    vHeaders.Headers(25) = "Cpr. Int.Acum. Monto"

    vHeaders.Headers(26) = "Costo Neto"
    vHeaders.Headers(27) = "PyD Tipo"
    vHeaders.Headers(28) = "PyD Base"
    vHeaders.Headers(29) = "PyD Dias Acum."
    vHeaders.Headers(30) = "PyD Monto Acum."
    vHeaders.Headers(31) = "PyD Saldo"
    
    vHeaders.Headers(32) = "Int. Acum. Dias"
    vHeaders.Headers(33) = "Int. Acum. Monto"
    
    vHeaders.Headers(34) = "Particip. Número"
    vHeaders.Headers(35) = "Particip. Valor"
    
    vHeaders.Headers(36) = "Base Cálculo"
    vHeaders.Headers(37) = "Periodicidad"
    
    vHeaders.Headers(38) = "Cta. Inversión"
    vHeaders.Headers(39) = "Cta. Int.Acum."
    vHeaders.Headers(40) = "Cta. Prima/Desc."
    vHeaders.Headers(41) = "Cta. Ingresos"
    
    vHeaders.Headers(42) = "Corte"
    vHeaders.Headers(43) = "Recurso"
    vHeaders.Headers(44) = "Monto Inversión"


 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_SCGI_" & cboCierre.Text)
       


End Sub

Private Sub cboCupones_Click()
If vPaso Then Exit Sub

Call sbIVR_Cupones_Load(lswCupones, scCaso.Tag, Mid(cboCupones.Text, 1, 1))

End Sub

Private Sub cboFi_Estado_Click()
If vPaso Then Exit Sub

Call sbIVR_Fondos_Load(lswFi_Mov, scCaso.Tag, cboFi_Estado.Text, cboFi_Tipo.Text)

End Sub

Private Sub cboFi_Tipo_Click()
If vPaso Then Exit Sub

Call sbIVR_Fondos_Load(lswFi_Mov, scCaso.Tag, cboFi_Estado.Text, cboFi_Tipo.Text)

End Sub


Private Sub chkFechaCompra_Click()
If chkFechaCompra.Value = xtpChecked Then
    dtpFechaCompra(0).Enabled = False
Else
    dtpFechaCompra(0).Enabled = True
End If

dtpFechaCompra(1).Enabled = dtpFechaCompra(0).Enabled
End Sub

Private Sub chkFechaVence_Click()
If chkFechaVence.Value = xtpChecked Then
    dtpFechaVence(0).Enabled = False
Else
    dtpFechaVence(0).Enabled = True
End If

dtpFechaVence(1).Enabled = dtpFechaVence(0).Enabled

End Sub

Private Sub Form_Load()


vPaso = True

cboCupones.Clear
cboCupones.AddItem "Registrados"
cboCupones.AddItem "Proyectados"
cboCupones.Text = "Registrados"


cboFi_Estado.Clear
cboFi_Estado.AddItem "Todos"
cboFi_Estado.AddItem "Activos"
cboFi_Estado.AddItem "Anulados"
cboFi_Estado.Text = "Todos"

cboFi_Tipo.Clear
cboFi_Tipo.AddItem "Todos"
cboFi_Tipo.AddItem "Aportaciones"
cboFi_Tipo.AddItem "Retiros"
cboFi_Tipo.AddItem "Fondos Cierre"
cboFi_Tipo.AddItem "Cambio Partic"
cboFi_Tipo.Text = "Todos"

vPaso = False

End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Height = Me.Height

tcMain.Width = Me.Width - (tcMain.Left + 250)
tcMain.Height = Me.Height - 450

vGrid.Width = tcMain.Width
vGrid.Height = tcMain.Height - (vGrid.Top + tcCaso.Height + scCaso.Height + 250)

scCaso.Top = vGrid.Top + vGrid.Height + 100
scCaso.Width = vGrid.Width

tcCaso.Top = scCaso.Top + scCaso.Height + 100
tcCaso.Width = vGrid.Width

btnExport.Top = scCaso.Top + 60

lswCierres.Width = tcCaso.Width
lswCupones.Width = tcCaso.Width
lswFi_Mov.Width = tcCaso.Width
lswFlujos_Ingresos.Width = tcCaso.Width
lswFlujos_PyD.Width = tcCaso.Width

'Vencimientos

gVencimientos.Width = tcMain.Width
gVencimientos.Height = tcMain.Height - (gVencimientos.Top + 250)


'Asientos
lswAsientoMain.Width = vGrid.Width
lswAsiento.Width = vGrid.Width
gbAsiento.Width = vGrid.Width
scAsientos.Width = vGrid.Width

lswAsientoMain.Height = tcMain.Height - (scAsientos.Height + lswAsiento.Height + gbAsiento.Height + 1350)

scAsientos.Top = lswAsientoMain.Top + lswAsientoMain.Height + 50
lswAsiento.Top = scAsientos.Top + scAsientos.Height + 50
gbAsiento.Top = lswAsiento.Top + lswAsiento.Height + 50




End Sub




Private Sub gVencimientos_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

Dim pTituloId As Long

With gVencimientos
    .Row = Row
    .Col = 2
    pTituloId = .Text
End With

If Col = 1 Then
    Dim frm As Form
    
    Call sbFormsCall("frmIVR_Rec_Titulos", 0, 0, 0, False, Me, False)
    Call sbFormActivo("frmIVR_Rec_Titulos", frm)
    
    Call frm.sbConsulta_Externa(pTituloId)
    
End If

End Sub

Private Sub lswAsientoMain_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub


scAsientos.Caption = Item.SubItems(1)

Call sbIVR_Asiento_Detalle(lswAsiento, Item.SubItems(1), Item.Text)

Dim pDebito As Currency, pCredito As Currency, i As Long

pDebito = 0
pCredito = 0

'Totales
With lswAsiento.ListItems

For i = 1 To .Count
    pDebito = pDebito + CCur(.Item(i).SubItems(3))
    pCredito = pCredito + CCur(.Item(i).SubItems(4))
Next i

End With

txtA_Debito.Text = Format(pDebito, "Standard")
txtA_Credito.Text = Format(pCredito, "Standard")
txtA_Diferencia.Text = Format(pDebito - pCredito, "Standard")

End Sub

Private Sub tcCaso_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If CLng(scCaso.Tag) = 0 Then
      tcCaso.Item(0).Selected = True
      Exit Sub
End If

Dim i As Integer, pMonto As Currency
Dim pDebito As Currency, pCredito As Currency
 

Select Case Item.Index
    Case 0 'Intereses
        Call sbIVR_Flujos_Load(lswFlujos_Ingresos, scCaso.Tag, "INT")
    
    Case 1 'Primas y Descuentos
        Call sbIVR_Flujos_Load(lswFlujos_PyD, scCaso.Tag, "PYD")
    
    Case 2 'Fondos Mov
        Call sbIVR_Fondos_Load(lswFi_Mov, scCaso.Tag, cboFi_Estado.Text, cboFi_Tipo.Text)
    
    Case 3 'Cupones
        Call sbIVR_Cupones_Load(lswCupones, scCaso.Tag, Mid(cboCupones.Text, 1, 1))
    
    Case 4 'Cierres
        Call sbIVR_Cierres_Load(lswCierres, scCaso.Tag)
    
'    Case 1 'Transac
'        Call sbAdquisicion_Load
'    Case 2 'Comisiones
'        Call sbIVR_Comisiones_Load(lswCom, scCaso.Tag)
'
'        'Totales
'        With lswCom.ListItems
'
'        pMonto = 0
'        For i = 1 To .Count
'            pMonto = pMonto + CCur(.Item(i).SubItems(2))
'        Next i
'        End With
'
'        txtComisionTotal.Text = Format(pMonto, "Standard")
'
'    Case 7 'Asiento
'        Call sbIVR_Asiento_Load(lswAsiento, scCaso.Tag)
'
'        pDebito = 0
'        pCredito = 0
'
'        'Totales
'        With lswAsiento.ListItems
'
'        For i = 1 To .Count
'            pDebito = pDebito + CCur(.Item(i).SubItems(3))
'            pCredito = pCredito + CCur(.Item(i).SubItems(4))
'        Next i
'
'        End With
'
'        txtA_Debito.Text = Format(pDebito, "Standard")
'        txtA_Credito.Text = Format(pCredito, "Standard")
'        txtA_Diferencia.Text = Format(pDebito - pCredito, "Standard")

End Select




End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Call Form_Resize

Select Case Item.Index
    Case 1 'Vencimientos
        Call btnCupon_Click
        
    Case 3 'Asientos
        Call btnAsientos_Click
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

Dim vFecha As Date

vPaso = True

tcMain.Item(0).Selected = True

cboBaseCalculo.Clear
cboBaseCalculo.AddItem "Base Actual"
cboBaseCalculo.ItemData(cboBaseCalculo.ListCount - 1) = CStr(365)
cboBaseCalculo.AddItem "Base 360"
cboBaseCalculo.ItemData(cboBaseCalculo.ListCount - 1) = CStr(360)
cboBaseCalculo.AddItem "Todas"
cboBaseCalculo.ItemData(cboBaseCalculo.ListCount - 1) = CStr(0)

cboBaseCalculo.Text = "Todas"

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.AddItem "Activo"
cboEstado.AddItem "Liquidado"
cboEstado.AddItem "Solicitado"
cboEstado.Text = "Activo"

strSQL = "select  isnull(max(CORTE), dbo.mygetdate())  as 'CORTE'" _
       & "  From IVR_CIERRES"
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!Corte
rs.Close

dtpFechaCompra(0).Value = DateAdd("y", -5, vFecha)
dtpFechaCompra(1).Value = DateAdd("m", 1, vFecha)

dtpFechaVence(0).Value = DateAdd("d", 1, vFecha)
dtpFechaVence(1).Value = DateAdd("m", 1, vFecha)


dtpFechaAsiento(0).Value = DateAdd("d", 1, vFecha)
dtpFechaAsiento(1).Value = DateAdd("m", 1, vFecha)

dtpFechaCupon(0).Value = DateAdd("d", 1, vFecha)
dtpFechaCupon(1).Value = DateAdd("m", 1, vFecha)


cboCuponEstado.Clear
cboCuponEstado.AddItem "Pendiente"
cboCuponEstado.AddItem "Cancelado"
cboCuponEstado.Text = "Pendiente"

strSQL = " select Corte as 'IdX',  CONVERT(varchar,Corte,23)   as 'ItmX'" _
       & " From IVR_CIERRES   order by corte desc"
Call sbCbo_Llena_New(cboCierre, strSQL, False, True)

 cboCierre.AddItem "Actual"
 cboCierre.ItemData(cboCierre.ListCount - 1) = "Actual"
 cboCierre.Text = "Actual"

strSQL = "select rtrim(COD_INSTRUMENTO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & "  From IVR_INSTRUMENTOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboInstrumento, strSQL, True, True)

strSQL = "select  rtrim(COD_RECURSO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_FUENTE_RECURSOS" _
       & " Where ACTIVA = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)

strSQL = "select  rtrim(COD_EMISOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_EMISORES" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboEmisor, strSQL, True, True)


strSQL = "select  rtrim(COD_ADMINISTRADOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_ADMINISTRADOR" _
       & " Where ESTADO = 'A'" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboAdministrador, strSQL, True, True)

strSQL = "select  rtrim(COD_PORTAFOLIO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PORTAFOLIOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboPortafolio, strSQL, True, True)


strSQL = "select  rtrim(COD_CATEGORIA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_CATEGORIA_TIPOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboClasificacion, strSQL, True, True)


strSQL = "select  rtrim(COD_RESERVA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_RESERVAS" _
       & " Where ACTIVA = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboReserva, strSQL, True, True)


strSQL = "select  rtrim(COD_PERIODICIDAD) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PERIODICIDAD" _
       & " Where ACTIVA = 1" _
       & " order by dias"
Call sbCbo_Llena_New(cboPeriodicidad, strSQL, True, True)


strSQL = "select rtrim(COD_DIVISA) AS 'Idx', rtrim(DESCRIPCION) as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

vPaso = False


Call chkFechaCompra_Click
Call chkFechaVence_Click

Call sbBuscar

'Call sbInicializa

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCuenta_Consulta(pCuenta As Object, pDesc As Object)
   frmCntX_ConsultaCuentas.Show vbModal
   pCuenta.Text = gCuenta
   pDesc.Text = fxgCntCuentaDesc(gCuenta)
   pCuenta.Text = fxgCntCuentaFormato(True, pCuenta.Text, 0)
End Sub

Private Sub sbCuenta_LostFocus(pCuenta As Object, pDesc As Object)
   
   pDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, pCuenta.Text, 0))
   pCuenta.Text = fxgCntCuentaFormato(True, pCuenta.Text, 0)

End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCuenta, txtCuentaDesc)
End If
End Sub

Private Sub txtCuenta_LostFocus()
    Call sbCuenta_LostFocus(txtCuenta, txtCuentaDesc)
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pTituloId As Long, pDescripcion As String
With vGrid
    .Row = Row
    .Col = 3
    pTituloId = .Text
    .Col = 5
    pDescripcion = "Operación: " & .Text
    .Col = 6
    pDescripcion = pDescripcion & " ¦ Serie: " & .Text
    .Col = 7
    pDescripcion = pDescripcion & " ¦ Isin: " & .Text
    .Col = 8
    pDescripcion = pDescripcion & " ¦ Administrador: " & .Text
    .Col = 9
    pDescripcion = pDescripcion & " ¦ Instrumento: " & .Text
End With

If Col = 1 Then
    scCaso.Tag = pTituloId
    scCaso.Caption = "[" & scCaso.Tag & "]  " & pDescripcion
    
    tcCaso.Item(0).Selected = True
    Call sbIVR_Flujos_Load(lswFlujos_Ingresos, pTituloId, "INT")
End If

If Col = 2 Then
    Dim frm As Form
    
    Call sbFormsCall("frmIVR_Rec_Titulos", 0, 0, 0, False, Me, False)
    Call sbFormActivo("frmIVR_Rec_Titulos", frm)
    
    Call frm.sbConsulta_Externa(pTituloId)
    
End If


End Sub
