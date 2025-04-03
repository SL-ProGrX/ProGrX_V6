VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_Convenios_Rebajos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Convenios: Registro de Rebajos por Cobrar "
   ClientHeight    =   7065
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5772
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   8892
      _Version        =   1310722
      _ExtentX        =   15684
      _ExtentY        =   10181
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
      Item(0).Caption =   "Asignación"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "chkActivo"
      Item(0).Control(1)=   "txtNTransaccion"
      Item(0).Control(2)=   "txtNPagos"
      Item(0).Control(3)=   "UpDown_NPagos"
      Item(0).Control(4)=   "txtCargoCod"
      Item(0).Control(5)=   "txtCargoDesc"
      Item(0).Control(6)=   "txtDocumento"
      Item(0).Control(7)=   "txtMonto"
      Item(0).Control(8)=   "txtDetalle"
      Item(0).Control(9)=   "dtpCobroInicia"
      Item(0).Control(10)=   "Label3(5)"
      Item(0).Control(11)=   "Label3(3)"
      Item(0).Control(12)=   "Label3(0)"
      Item(0).Control(13)=   "Label3(1)"
      Item(0).Control(14)=   "Label3(2)"
      Item(0).Control(15)=   "Label3(4)"
      Item(0).Control(16)=   "Label3(6)"
      Item(0).Control(17)=   "Label4(0)"
      Item(0).Control(18)=   "Label5"
      Item(0).Control(19)=   "lblSaldo"
      Item(0).Control(20)=   "lblRecaudado"
      Item(0).Control(21)=   "Label3(9)"
      Item(0).Control(22)=   "FlatScrollBarTransac"
      Item(0).Control(23)=   "txtNotas"
      Item(1).Caption =   "Cobros Realizados"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "lswPagos"
      Item(1).Control(2)=   "rbCobros(0)"
      Item(1).Control(3)=   "rbCobros(1)"
      Item(1).Control(4)=   "rbCobros(2)"
      Item(2).Caption =   "Informes"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "GroupBox1"
      Begin XtremeSuiteControls.ListView lswPagos 
         Height          =   2052
         Left            =   -70000
         TabIndex        =   46
         Top             =   3600
         Visible         =   0   'False
         Width           =   8892
         _Version        =   1310722
         _ExtentX        =   15684
         _ExtentY        =   3619
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2772
         Left            =   -70000
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   8892
         _Version        =   1310722
         _ExtentX        =   15684
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.RadioButton rbCobros 
         Height          =   252
         Index           =   0
         Left            =   -66160
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   252
         Left            =   6720
         TabIndex        =   44
         Top             =   360
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activo ?"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
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
      Begin VB.TextBox txtNPagos 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "1"
         Top             =   4440
         Width           =   384
      End
      Begin XtremeSuiteControls.UpDown UpDown_NPagos 
         Height          =   252
         Left            =   3480
         TabIndex        =   6
         Top             =   4440
         Width           =   252
         _Version        =   1310722
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   64
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Min             =   1
         Max             =   36
         BuddyControl    =   ""
         BuddyProperty   =   ""
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarTransac 
         Height          =   252
         Left            =   8160
         TabIndex        =   19
         Top             =   720
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   4572
         Left            =   -69760
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   8172
         _Version        =   1310722
         _ExtentX        =   14414
         _ExtentY        =   8064
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Pagos de Cargos Inmediatos"
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
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   29
            Top             =   480
            Width           =   3135
         End
         Begin VB.CommandButton cmdReporte 
            Caption         =   "&Reporte"
            Height          =   735
            Left            =   6480
            Picture         =   "frmCR_Convenios_Rebajos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkTodas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Todas las Fechas"
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
            Height          =   375
            Left            =   6000
            TabIndex        =   27
            Top             =   2640
            Width           =   1575
         End
         Begin VB.ComboBox cboEstado 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            ItemData        =   "frmCR_Convenios_Rebajos.frx":0173
            Left            =   1080
            List            =   "frmCR_Convenios_Rebajos.frx":0175
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2640
            Width           =   4455
         End
         Begin VB.ComboBox cboCargo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2280
            Width           =   4455
         End
         Begin VB.ComboBox cboProveedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1920
            Width           =   4455
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Pagos de Cargos Periódicos"
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   3135
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Saldos de Cargos Registrados"
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   3135
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "Cargos Registrados"
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Value           =   -1  'True
            Width           =   3135
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInicio 
            Height          =   315
            Left            =   6360
            TabIndex        =   51
            Top             =   1920
            Width           =   1335
            _Version        =   1310722
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
            Height          =   315
            Left            =   6360
            TabIndex        =   52
            Top             =   2280
            Width           =   1335
            _Version        =   1310722
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
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   2640
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Corte"
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
            Left            =   5640
            TabIndex        =   33
            Top             =   2280
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio"
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
            Left            =   5640
            TabIndex        =   32
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label2 
            Caption         =   "Cargo"
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
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   2280
            Width           =   972
         End
         Begin VB.Label Label2 
            Caption         =   "Proveedor"
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
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   972
         Left            =   1200
         TabIndex        =   35
         Top             =   1800
         Width           =   6852
         _Version        =   1310722
         _ExtentX        =   12086
         _ExtentY        =   1714
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   972
         Left            =   1200
         TabIndex        =   36
         Top             =   2880
         Width           =   6852
         _Version        =   1310722
         _ExtentX        =   12086
         _ExtentY        =   1714
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   6240
         TabIndex        =   38
         Top             =   1440
         Width           =   1812
         _Version        =   1310722
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargoCod 
         Height          =   312
         Left            =   1200
         TabIndex        =   41
         Top             =   1080
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
         Height          =   312
         Left            =   2880
         TabIndex        =   42
         Top             =   1080
         Width           =   5172
         _Version        =   1310722
         _ExtentX        =   9123
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
      Begin XtremeSuiteControls.FlatEdit txtNTransaccion 
         Height          =   312
         Left            =   6240
         TabIndex        =   43
         Top             =   720
         Width           =   1812
         _Version        =   1310722
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbCobros 
         Height          =   252
         Index           =   1
         Left            =   -64720
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
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
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbCobros 
         Height          =   252
         Index           =   2
         Left            =   -62800
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cancelados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCobroInicia 
         Height          =   312
         Left            =   6720
         TabIndex        =   50
         Top             =   4320
         Width           =   1332
         _Version        =   1310722
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
      Begin VB.Label Label3 
         Caption         =   "Fecha de Inicio del Cobro de este cargo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   9
         Left            =   4080
         TabIndex        =   18
         Top             =   4320
         Width           =   2532
      End
      Begin VB.Label lblRecaudado 
         Alignment       =   1  'Right Justify
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
         Height          =   312
         Left            =   5880
         TabIndex        =   17
         Top             =   5280
         Width           =   2172
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
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
         Height          =   312
         Left            =   5880
         TabIndex        =   16
         Top             =   4920
         Width           =   2172
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recaudado"
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
         Height          =   312
         Left            =   4800
         TabIndex        =   15
         Top             =   5280
         Width           =   1092
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
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
         Height          =   312
         Index           =   0
         Left            =   4800
         TabIndex        =   14
         Top             =   4920
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   972
      End
      Begin VB.Label Label3 
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
         Left            =   4800
         TabIndex        =   12
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Documento"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Cargo"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   4680
         TabIndex        =   9
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Realizar en el Cobro en -N- pagos:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Top             =   4320
         Width           =   1692
      End
      Begin VB.Label Label3 
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
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   972
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   6816
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Divisa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Tipo de Cambio"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8280
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2880
      TabIndex        =   40
      Top             =   480
      Width           =   5292
      _Version        =   1310722
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   39
      Top             =   480
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmCR_Convenios_Rebajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean
Dim vDivisa As String, vTipoCambio As Currency, vPaso As Boolean


Private Sub FlatScrollBarTransac_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtNTransaccion.Text) Then
    txtNTransaccion.Text = 0
End If

If vScroll Then
    strSQL = "select Top 1 ID_TRANSAC from CRD_CONVENIOS_CARGOS_CXP_CONTROL"
           
    If FlatScrollBarTransac.Value = 1 Then
       strSQL = strSQL & " where COD_CONVENIO = '" & txtCodigo.Text & "' AND ID_TRANSAC > " & txtNTransaccion.Text & " order by ID_TRANSAC asc"
    Else
       strSQL = strSQL & " where COD_CONVENIO = '" & txtCodigo.Text & "' AND ID_TRANSAC < " & txtNTransaccion.Text & " order by ID_TRANSAC desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        txtNTransaccion.Text = rs!ID_TRANSAC
        If IsNumeric(txtNTransaccion.Text) Then
            Call sbConsulta(txtNTransaccion.Text)
        End If
    End If
    rs.Close
End If

vScroll = False
FlatScrollBarTransac.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub rbCobros_Click(Index As Integer)
Call sbCobrosRealizados
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNPagos.SetFocus
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curMonto As Currency

If lsw.ListItems.Count = 0 Then Exit Sub
If lsw.SelectedItem.Text = "" Then Exit Sub

curMonto = 0
lswPagos.ListItems.Clear

strSQL = " select O.COD_ORDEN, O.FECHA_CORTE, O.NOTAS , O.COD_FACTURA, P.ID_FRACCION, P.MONTO " _
       & " from CRD_CONVENIOS_ORDENES O inner join  CRD_CONVENIOS_DT_CARGOS_CXP P on O.COD_CONVENIO= P.COD_CONVENIO and O.COD_ORDEN = P.COD_ORDEN" _
       & " where P.COD_CONVENIO = '" & vCodigo & "' and P.ID_TRANSAC = " & Item.Text _
       & " and O.ESTADO = 'C'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswPagos.ListItems.Add(, , rs!ID_FRACCION)
     itmX.SubItems(1) = rs!cod_factura & ""
     itmX.SubItems(2) = Format(rs!FECHA_CORTE, "yyyy/mm/dd")
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = rs!NOTAS
 curMonto = curMonto + rs!Monto
 rs.MoveNext
Loop
rs.Close

Set itmX = lswPagos.ListItems.Add(, , "")
     itmX.SubItems(3) = "____________"
Set itmX = lswPagos.ListItems.Add(, , "TOTAL")
     itmX.SubItems(3) = Format(curMonto, "Standard")
     
     
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
  Case 1 'Pagos Realizados
    Call sbCobrosRealizados
  Case 2 'Reportes
    Call sbInicializaReportes
End Select
End Sub

Private Sub txtNPagos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCobroInicia.SetFocus
End Sub


Private Sub chkTodas_Click()
If chkTodas.Value = vbChecked Then
 dtpInicio.Enabled = False
 dtpCorte.Enabled = False
Else
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COD_CONVENIO,descripcion from CRD_CONVENIOS"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_CONVENIO > '" & txtCodigo.Text & "' order by COD_CONVENIO asc"
    Else
       strSQL = strSQL & " where COD_CONVENIO < '" & txtCodigo.Text & "' order by COD_CONVENIO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_CONVENIO
      txtNombre.Text = rs!Descripcion
      Call txtCodigo_LostFocus
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
vModulo = 16
End Sub

Private Sub Form_Load()

vModulo = 16

On Error GoTo vError

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
   .Clear
   .Add , , "No. Transac", 1200, vbCenter
   .Add , , "Cargo", 1200, vbCenter
   .Add , , "Descripción", 3200
   .Add , , "Documento", 1800
   .Add , , "Detalle", 2800
   .Add , , "Monto", 1400, vbRightJustify
   .Add , , "Saldo", 1400, vbRightJustify
   .Add , , "Recaudado", 1400, vbRightJustify
   .Add , , "Inicio Cobro", 1800
   .Add , , "Fracciones", 1200, vbCenter
   .Add , , "Activo", 1100, vbCenter
   .Add , , "Notas", 4200
End With

With lswPagos.ColumnHeaders
   .Clear
   .Add , , "No. Transac", 1200, vbCenter
   .Add , , "No. Factura", 2200
   .Add , , "Fecha Corte", 1800
   .Add , , "Monto", 1800, vbRightJustify
   .Add , , "Notas", 4200
End With


 vScroll = False
    FlatScrollBar.Value = 0
    FlatScrollBarTransac.Value = 0
 vScroll = True
 

dtpCobroInicia.Value = fxFechaServidor
dtpCobroInicia.MinDate = dtpCobroInicia.Value
dtpCobroInicia.MaxDate = DateAdd("d", 45, dtpCobroInicia.Value)
 
 
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

Private Sub sbLimpiaPantalla(Optional pNuevo As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset

If pNuevo Then
    txtCodigo.Enabled = False
    txtNTransaccion.Text = 0
Else
    vCodigo = ""
    txtCodigo.Text = ""
    txtNombre.Text = ""
    txtCodigo.Enabled = True
End If

txtCargoCod.Text = ""
txtCargoDesc.Text = ""

txtDocumento.Text = ""
txtDetalle.Text = ""
txtNotas.Text = ""
txtMonto = "0.00"

lblSaldo.Caption = "0.00"
lblRecaudado.Caption = "0.00"

chkActivo.Value = vbChecked
dtpCobroInicia.Value = fxFechaServidor

StatusBarX.Panels(1).Text = "Usr: "
StatusBarX.Panels(2).Text = "Reg: "
StatusBarX.Panels(3).Text = "Pagos: "


tcMain.Item(0).Selected = True

End Sub


Private Sub sbCobrosRealizados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lswPagos.ListItems.Clear

strSQL = "select C.*,D.descripcion as CargoDesc" _
       & " from CRD_CONVENIOS_CARGOS_CXP_CONTROL C inner join cxp_cargos D on C.cod_cargo = D.cod_cargo" _
       & " where C.COD_CONVENIO = '" & vCodigo & "'"
       
Select Case True
  Case rbCobros.Item(0).Value
  Case rbCobros.Item(1).Value
        strSQL = strSQL & " and Saldo > 0"
  Case rbCobros.Item(2).Value
        strSQL = strSQL & " and Saldo <= 0"
End Select
       
       
strSQL = strSQL & " order by C.ID_TRANSAC desc"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!ID_TRANSAC)
     itmX.SubItems(1) = rs!cod_cargo
     itmX.SubItems(2) = rs!CargoDesc
     itmX.SubItems(3) = rs!Documento
     itmX.SubItems(4) = rs!Detalle
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = Format(rs!SALDO, "Standard")
     itmX.SubItems(7) = Format(rs!Monto - rs!SALDO, "Standard")
     itmX.SubItems(8) = Format(rs!COBRO_INICIO_FECHA, "yyyy/mm/dd")
     itmX.SubItems(9) = rs!COBRO_FRACCIONES
     itmX.SubItems(10) = rs!Activo
     itmX.SubItems(11) = rs!NOTAS
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbInicializaReportes()
Dim strSQL As String, rs As New ADODB.Recordset

cboEstado.AddItem "01 - En Cobro"
cboEstado.AddItem "02 - Cancelados"
cboEstado.AddItem "03 - Todos"
cboEstado.Text = "03 - Todos"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

chkTodas.Value = vbUnchecked

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        vEdita = False
        Call sbLimpiaPantalla(True)
        
        txtCodigo.Enabled = False
        Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.Enabled = False
      txtCargoCod.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
      txtCodigo.Enabled = True
      Call sbToolBar(tlb, "activo")
        
    
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtNTransaccion.Text = "" Or txtNTransaccion.Text = "0" Then
        Call sbLimpiaPantalla(True)
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(txtNTransaccion.Text)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_CONVENIO,descripcion from CRD_CONVENIOS"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
       tcMain.Item(2).Selected = True
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(pTransac As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spConvenios_Rebajos_Programa_Consulta '" & txtCodigo.Text & "'," & txtNTransaccion.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!COD_CONVENIO
  txtCodigo = rs!COD_CONVENIO
  txtNombre = rs!ConvenioDesc
  
  txtCargoCod = rs!cod_cargo
  txtCargoDesc = rs!CargoDesc
  
  txtDocumento.Text = rs!Documento
  txtDetalle.Text = rs!Detalle
  txtNotas.Text = rs!NOTAS
  
  lblSaldo.Caption = Format(rs!SALDO, "Standard")
  lblRecaudado.Caption = Format((rs!Recaudado), "Standard")
  
  txtMonto = Format(rs!Monto, "Standard")
    
  txtNTransaccion.Text = rs!ID_TRANSAC
  chkActivo.Value = rs!Activo
  
  UpDown_NPagos.Value = rs!COBRO_FRACCIONES
  txtNPagos.Text = rs!COBRO_FRACCIONES
  
  StatusBarX.Panels(1).Text = "Usr: " & rs!registro_usuario & ""
  StatusBarX.Panels(2).Text = "Reg: " & rs!registro_Fecha & ""
  StatusBarX.Panels(3).Text = "Pagos: " & rs!Numero_Pagos & ""
    
  'Puede dar Error por Bloqueo de Fechas
  If rs!COBRO_INICIO_FECHA > dtpCobroInicia.MinDate And rs!COBRO_INICIO_FECHA < dtpCobroInicia.MaxDate Then
    dtpCobroInicia.Value = rs!COBRO_INICIO_FECHA
  End If
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxRecaudado(pCargoID As Long, pConvenio As Long) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

fxRecaudado = 0

strSQL = "select isnull(sum(C.monto),0) as Monto" _
       & " from cxp_pagoprov P inner join cxp_pagoprovcargos C" _
       & " on P.npago = C.npago and P.COD_CONVENIO = C.COD_CONVENIO" _
       & " and P.cod_factura = C.cod_factura and P.tesoreria is not null" _
       & " Where C.id = " & pCargoID & " And C.COD_CONVENIO = " & pConvenio
Call OpenRecordSet(rs, strSQL)
    fxRecaudado = rs!Monto
rs.Close

End Function

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Convenio no es válido ..."
If txtDocumento = "" Then vMensaje = vMensaje & vbCrLf & " - El documento no es válido ..."
If Not IsNumeric(txtMonto) Then
  vMensaje = vMensaje & vbCrLf & " - El monto no es válido ..."
Else
  If CCur(txtMonto) <= 0 Then vMensaje = vMensaje & vbCrLf & " - El valor (no marca ningun rango de referencia) no es válido ..."
End If

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long

On Error GoTo vError

If vEdita Then
                       
  strSQL = "exec spConvenios_Rebajos_Programa '" & vCodigo & "'," & txtNTransaccion.Text & "," & txtNPagos.Text & ",'" & txtCargoCod.Text _
          & "'," & CCur(txtMonto) & ",'" & Format(dtpCobroInicia.Value, "yyyy/mm/dd") & "','" & txtDocumento.Text _
          & "','" & txtDetalle.Text & "','" & txtNotas.Text & "','" & glogon.Usuario _
          & "'," & chkActivo.Value & ",'E'"
  Call ConectionExecute(strSQL, , i)

  If i > 0 Then
      Call Bitacora("Modifica", "Rebajo Programado:" & vCodigo & " ID: " & txtNTransaccion.Text)
  End If

Else

  strSQL = "exec spConvenios_Rebajos_Programa '" & vCodigo & "'," & txtNTransaccion.Text & "," & txtNPagos.Text & ",'" & txtCargoCod.Text _
          & "'," & CCur(txtMonto) & ",'" & Format(dtpCobroInicia.Value, "yyyy/mm/dd") & "','" & txtDocumento.Text _
          & "','" & txtDetalle.Text & "','" & txtNotas.Text & "','" & glogon.Usuario _
          & "'," & chkActivo.Value & ",'A'"
  Call OpenRecordSet(rs, strSQL)
   
  txtNTransaccion.Text = rs!TransaccionID & ""
  
  rs.Close
   
   If IsNumeric(txtNTransaccion.Text) Then
        Call Bitacora("Registra", "Rebajo Programado:" & vCodigo & " ID: " & txtNTransaccion.Text)
   End If
End If


'Activa el codigo
txtCodigo.Enabled = True

'Actualiza todos los datos
Call txtCodigo_LostFocus
Call RefrescaTags(Me)


 MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Long, strSQL As String

On Error GoTo vError

If txtNTransaccion.Text = "" Then Exit Sub

If lblRecaudado.Caption > 0 Then
    MsgBox "No se puede eliminar el Cargo porque ya tiene recaudación aplicada...!", vbExclamation
    Exit Sub
End If

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  
  strSQL = "exec spConvenios_Rebajos_Programa '" & vCodigo & "'," & txtNTransaccion & ",0,'" & txtCargoCod.Text _
          & "'," & CCur(txtMonto) & ",'" & Format(dtpCobroInicia.Value, "yyyy/mm/dd") & "','" & txtDocumento.Text _
          & "','" & txtDetalle.Text & "','" & txtNotas.Text & "','" & glogon.Usuario _
          & "'," & chkActivo.Value & ",'B'"
  Call ConectionExecute(strSQL, , i)
  
  If i > 0 Then
    Call Bitacora("Elimina", "Rebajo Programado:" & vCodigo & " ID: " & txtNTransaccion.Text & "..Mnt..:" & txtMonto.Text)
  End If
  Call txtCodigo_LostFocus
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Ca.COD_CARGO"
  gBusquedas.Orden = "Ca.COD_CARGO"
  gBusquedas.Consulta = "select Ca.COD_CARGO, Ca.DESCRIPCION" _
        & " from CRD_CONVENIOS_CARGOS_CXP Cc inner join CXP_CARGOS Ca on Cc.COD_CARGO = Ca.COD_CARGO"
  gBusquedas.Filtro = " AND Cc.COD_CONVENIO = '" & txtCodigo.Text & "' and Ca.ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCargoCod_LostFocus()
txtCargoDesc = fxSIFCCodigos("D", txtCargoCod, "CargosProv")
End Sub

Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Ca.DESCRIPCION"
  gBusquedas.Orden = "Ca.DESCRIPCION"
  gBusquedas.Consulta = "select Ca.COD_CARGO, Ca.DESCRIPCION" _
        & " from CRD_CONVENIOS_CARGOS_CXP Cc inner join CXP_CARGOS Ca on Cc.COD_CARGO = Ca.COD_CARGO"
  gBusquedas.Filtro = " AND Cc.COD_CONVENIO = '" & txtCodigo.Text & "' and Ca.ACTIVO = 1"
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_CONVENIO"
  gBusquedas.Orden = "COD_CONVENIO"
  gBusquedas.Consulta = "select COD_CONVENIO,descripcion from CRD_CONVENIOS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select COD_CONVENIO,descripcion" _
       & " from CRD_CONVENIOS where COD_CONVENIO = '" & txtCodigo.Text & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

  txtNombre.Text = rs!Descripcion
  vCodigo = rs!COD_CONVENIO

End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNTransaccion.SetFocus
On Error GoTo vError
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CONVENIO,descripcion from CRD_CONVENIOS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  vCodigo = txtCodigo
  txtNombre = gBusquedas.Resultado2
End If
vError:
End Sub

Private Sub txtNTransaccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoCod.SetFocus

On Error GoTo vError

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "ID_TRANSAC"
  gBusquedas.Orden = "ID_TRANSAC"
  gBusquedas.Consulta = "select ID_TRANSAC,DOCUMENTO,DETALLE,MONTO, REGISTRO_FECHA from CRD_CONVENIOS_CARGOS_CXP_CONTROL"
  gBusquedas.Filtro = " AND COD_CONVENIO = '" & txtCodigo.Text & "'"
  frmBusquedas.Show vbModal
  
  txtNTransaccion.Text = gBusquedas.Resultado
  
  If IsNumeric(txtNTransaccion.Text) Then
      Call sbConsulta(txtNTransaccion.Text)
  End If
  
End If

vError:

End Sub

Private Sub UpDown_NPagos_DownClick()
 txtNPagos.Text = UpDown_NPagos.Value
End Sub

Private Sub UpDown_NPagos_UpClick()
 txtNPagos.Text = UpDown_NPagos.Value
End Sub
