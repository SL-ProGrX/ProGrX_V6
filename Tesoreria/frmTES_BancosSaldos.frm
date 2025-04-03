VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_BancosSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos de Cuentas Bancarias "
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   Icon            =   "frmTES_BancosSaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12945
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   12240
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6852
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   12252
      _Version        =   1441793
      _ExtentX        =   21611
      _ExtentY        =   12086
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
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "Label1(0)"
      Item(0).Control(1)=   "cboBancoC"
      Item(0).Control(2)=   "lsw"
      Item(1).Caption =   "Histórico"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lswH"
      Item(1).Control(1)=   "Label1(1)"
      Item(1).Control(2)=   "cboBancoH"
      Item(1).Control(3)=   "Label1(3)"
      Item(1).Control(4)=   "cboH"
      Item(2).Caption =   "Cierres y Aperturas"
      Item(2).ControlCount=   21
      Item(2).Control(0)=   "txtSM"
      Item(2).Control(1)=   "txtAjuste"
      Item(2).Control(2)=   "txtSF"
      Item(2).Control(3)=   "txtTC"
      Item(2).Control(4)=   "txtTD"
      Item(2).Control(5)=   "txtSI"
      Item(2).Control(6)=   "cmdAplicar"
      Item(2).Control(7)=   "Label2(9)"
      Item(2).Control(8)=   "Label2(8)"
      Item(2).Control(9)=   "Label2(7)"
      Item(2).Control(10)=   "Label2(6)"
      Item(2).Control(11)=   "Label2(5)"
      Item(2).Control(12)=   "Label2(4)"
      Item(2).Control(13)=   "Label2(3)"
      Item(2).Control(14)=   "dtpInicio"
      Item(2).Control(15)=   "dtpCorte"
      Item(2).Control(16)=   "Label2(2)"
      Item(2).Control(17)=   "cboBanco"
      Item(2).Control(18)=   "Label1(2)"
      Item(2).Control(19)=   "Label1(13)"
      Item(2).Control(20)=   "cbo"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5532
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   7332
         _Version        =   1441793
         _ExtentX        =   12933
         _ExtentY        =   9758
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
      End
      Begin XtremeSuiteControls.ListView lswH 
         Height          =   5532
         Left            =   -70000
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   12252
         _Version        =   1441793
         _ExtentX        =   21611
         _ExtentY        =   9758
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
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   552
         Left            =   -61480
         TabIndex        =   5
         Top             =   6000
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   974
         _StockProps     =   79
         Caption         =   "&Aplicar"
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
         Picture         =   "frmTES_BancosSaldos.frx":6852
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -61720
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -61720
         TabIndex        =   14
         Top             =   1920
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   -66640
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1441793
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.FlatEdit txtSI 
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
         Left            =   -63160
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTD 
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
         Left            =   -63160
         TabIndex        =   22
         Top             =   2760
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTC 
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
         Left            =   -63160
         TabIndex        =   23
         Top             =   3120
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSF 
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
         Left            =   -63160
         TabIndex        =   24
         Top             =   3480
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAjuste 
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
         Left            =   -63160
         TabIndex        =   25
         Top             =   3840
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSM 
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
         Left            =   -63160
         TabIndex        =   26
         Top             =   4200
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   -66640
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1441793
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.ComboBox cboBancoH 
         Height          =   312
         Left            =   -66520
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1441793
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.ComboBox cboH 
         Height          =   312
         Left            =   -66520
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1441793
         _ExtentX        =   11245
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
      Begin XtremeSuiteControls.ComboBox cboBancoC 
         Height          =   312
         Left            =   2400
         TabIndex        =   17
         Top             =   960
         Width           =   7332
         _Version        =   1441793
         _ExtentX        =   12938
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -68440
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Bancarias"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -68440
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entidad"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   -68560
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas Bancarias"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -68560
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   2
         Left            =   -62560
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   3
         Left            =   -62560
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   4
         Left            =   -65440
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(-)  Total Débitos"
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
         Left            =   -65440
         TabIndex        =   10
         Top             =   2760
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Total Créditos"
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
         Left            =   -65440
         TabIndex        =   9
         Top             =   3120
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Final"
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
         Left            =   -65440
         TabIndex        =   8
         Top             =   3480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste"
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
         Index           =   8
         Left            =   -65440
         TabIndex        =   7
         Top             =   3840
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Minimo (Control)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   9
         Left            =   -65440
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   $"frmTES_BancosSaldos.frx":7030
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
         Height          =   492
         Index           =   0
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   7332
      End
   End
   Begin XtremeSuiteControls.PushButton cmdSeguridad 
      Height          =   252
      Left            =   11880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Accesos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldos en Cuentas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4692
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_BancosSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()

If vPaso Then Exit Sub

Call sbCargaCierre

End Sub

Private Sub cboBanco_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
     Call sbMonitoreoCbo(cbo, cboBanco.ItemData(cboBanco.ListIndex))
vPaso = False

Call cbo_Click

End Sub

Private Sub cboBancoC_Click()
If vPaso Then Exit Sub

Call sbCuentas_Load

End Sub

Private Sub cboBancoH_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
     Call sbMonitoreoCbo(cboH, cboBancoH.ItemData(cboBancoH.ListIndex))
vPaso = False

Call cboH_Click

End Sub

Private Sub cboH_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lswH.ListItems.Clear

If vPaso Then Exit Sub
If cboH.ListCount = 0 Then Exit Sub



strSQL = "select * from TES_BANCOS_CIERRES where id_banco = " & cboH.ItemData(cboH.ListIndex)
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswH.ListItems.Add(, , rs!IdX)
     itmX.SubItems(1) = Format(rs!Inicio, "dd/mm/yyyy")
     itmX.SubItems(2) = Format(rs!Corte, "dd/mm/yyyy")
     itmX.SubItems(3) = Format(rs!saldo_inicial, "Standard")
     itmX.SubItems(4) = Format(rs!total_debitos, "Standard")
     itmX.SubItems(5) = Format(rs!total_creditos, "Standard")
     itmX.SubItems(6) = Format(rs!saldo_final, "Standard")
     itmX.SubItems(7) = Format(rs!ajuste, "Standard")
     itmX.SubItems(8) = Format(rs!saldo_minimo, "Standard")
     itmX.SubItems(9) = Format(rs!fecha, "dd/mm/yyyy")
     itmX.SubItems(10) = rs!Usuario
 rs.MoveNext
Loop
rs.Close
End Sub

Private Sub sbCargaCierre()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then Exit Sub

strSQL = " select corte,saldo_final,saldo_minimo from TES_BANCOS_CIERRES" _
       & " where idX = (select max(idX) from TES_BANCOS_CIERRES" _
       & " where id_banco = " & cbo.ItemData(cbo.ListIndex) & ")"
Call OpenRecordSet(rs, strSQL)

txtSI.Text = Format(0, "Standard")
txtTD.Text = Format(0, "Standard")
txtTC.Text = Format(0, "Standard")
txtSF.Text = Format(0, "Standard")
txtSM.Text = Format(0, "Standard")
txtAjuste.Text = Format(0, "Standard")

If Not rs.EOF And Not rs.BOF Then
  dtpInicio.Enabled = False
  dtpInicio.Value = DateAdd("d", 1, rs!Corte)
  txtSI.Text = Format(rs!saldo_final, "Standard")
  txtSM.Text = Format(rs!saldo_minimo, "Standard")
Else
  dtpInicio.Enabled = True
  txtSI.Text = 0
  txtSM.Text = 0
End If
rs.Close

Call dtpCorte_Change

End Sub


Private Sub sbCargaMovimientos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String

On Error GoTo vError

Me.MousePointer = vbHourglass

txtTD.Text = "0"
txtTC.Text = "0"

strSQL = "select ctaConta from Tes_Bancos where id_Banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
    vCuenta = rs!ctaConta & ""
    vCuenta = Trim(vCuenta)
rs.Close


'Saca Debitos y Creditos de las Cuentas Bancarias

'Emisiones de Documentos
strSQL = "select D.debehaber as Movimiento,sum(D.monto) as Total" _
       & " from Tes_Transacciones C inner join Tes_Trans_Asiento D on C.nsolicitud = D.nsolicitud" _
       & " where C.fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
       & " 23:59:59' and C.estado in('I','T','A') and D.cuenta_contable = '" & vCuenta _
       & "' group by D.debehaber"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 If rs!Movimiento = "D" Then
    txtTC.Text = Format(rs!Total, "Standard")
 Else
    txtTD.Text = Format(rs!Total, "Standard")
 End If
 rs.MoveNext
Loop
rs.Close


'Anulaciones de Documentos
strSQL = "select D.debehaber as Movimiento,sum(D.monto) as Total" _
       & " from Tes_Transacciones C inner join Tes_Trans_Asiento D on C.nsolicitud = D.nsolicitud" _
       & " where C.fecha_anula between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
       & " 23:59:59' and C.estado in('A') and D.cuenta_contable = '" & vCuenta _
       & "' group by D.debehaber"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 If rs!Movimiento = "D" Then
    txtTD.Text = Format(CCur(txtTD) + rs!Total, "Standard")
 Else
    txtTC.Text = Format(CCur(txtTC) + rs!Total, "Standard")
 End If
 rs.MoveNext
Loop
rs.Close

txtTD.Text = Format(CCur(txtTD.Text), "Standard")
txtTC.Text = Format(CCur(txtTC.Text), "Standard")


txtSF.Tag = CCur(txtSI) - CCur(txtTD) + CCur(txtTC)
txtSF.Text = Format(CCur(txtSI) - CCur(txtTD) + CCur(txtTC), "Standard")
txtAjuste.Text = "0"

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String

On Error GoTo vError

If dtpInicio.Value > dtpCorte.Value Then
  MsgBox "La fecha de Corte no puede ser menor a la fecha de Inicio, verifique...", vbExclamation
  Exit Sub
End If

strSQL = "insert TES_BANCOS_CIERRES(id_banco,fecha,usuario,inicio,corte,saldo_inicial" _
       & ",total_debitos,total_creditos,saldo_final,ajuste,saldo_minimo) values(" & cbo.ItemData(cbo.ListIndex) _
       & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'," & CCur(txtSI) & "," & CCur(txtTD) _
       & "," & CCur(txtTC) & "," & CCur(txtSF) & "," & CCur(txtAjuste) & "," & CCur(txtSM) & ")"
Call ConectionExecute(strSQL)

Call cbo_Click
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpCorte_Change()
If vPaso Then Exit Sub

Call sbCargaMovimientos

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub


Private Sub Form_Load()

vModulo = 9

vPaso = False

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Descripción", 3800
    .Add , , "Cuenta", 2700, vbCenter
 End With

 With lswH.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Inicio", 1200, vbCenter
    .Add , , "Corte", 1200, vbCenter
    .Add , , "Saldo Inicial", 2100, vbRightJustify
    .Add , , "Total Débitos", 2100, vbRightJustify
    .Add , , "Total Créditos", 2100, vbRightJustify
    .Add , , "Saldo Final", 2100, vbRightJustify
    .Add , , "Ajustes", 2100, vbRightJustify
    .Add , , "Saldo Mínimo", 2100, vbRightJustify
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Usuario", 1800, vbCenter
 End With

Call Formularios(Me)
Call RefrescaTags(Me)


lsw.Enabled = cmdSeguridad.Enabled
cmdAplicar.Enabled = cmdSeguridad.Enabled

End Sub



Private Sub sbMonitoreoCbo(cboX As Object, pGrupo As String)
Dim strSQL As String

vPaso = True
    strSQL = "select id_banco as 'IdX',rtrim(descripcion) as 'ItmX'" _
           & " from Tes_Bancos where monitoreo = 1 and cod_Grupo = '" & pGrupo & "'"
    Call sbCbo_Llena_New(cboX, strSQL, False, True)
vPaso = False
End Sub


Private Sub sbCuentas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbDefault

vPaso = True

lsw.ListItems.Clear
strSQL = "select id_banco,descripcion,cta,isnull(monitoreo,0) as Monitoreo" _
       & " from Tes_Bancos" _
       & " where cod_grupo = '" & cboBancoC.ItemData(cboBancoC.ListIndex) _
       & "' order by monitoreo desc, descripcion"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_Banco)
     itmX.SubItems(1) = rs!DESCRIPCION
     itmX.SubItems(2) = rs!Cta & ""
     itmX.Checked = IIf((rs!Monitoreo = 1), vbChecked, vbUnchecked)
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

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

If Item.Checked Then
  strSQL = "update Tes_Bancos set monitoreo = 1 where id_banco = " & Item.Text
Else
  strSQL = "update Tes_Bancos set monitoreo = 0 where id_banco = " & Item.Text
End If

Call ConectionExecute(strSQL)
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Me.MousePointer = vbDefault

Select Case Item.Index
  Case 0 'Seleccion de Tes_Bancos
     Call sbCuentas_Load
  
  Case 1 'Historico
     Call cboBancoH_Click

  Case 2 'Cierres
     dtpCorte.Value = fxFechaServidor
     Call cboBanco_Click

End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

vPaso = True
 
strSQL = "select rtrim(cod_grupo) as  'IdX', rtrim(Descripcion) as 'ItmX' from TES_BANCOS_GRUPOS" _
      & " where Activo = 1"
      
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
Call sbCbo_Llena_New(cboBancoH, strSQL, False, True)
Call sbCbo_Llena_New(cboBancoC, strSQL, False, True)
 
vPaso = False
 
tcMain.Item(0).Selected = True

Call sbCuentas_Load

End Sub

Private Sub txtSF_Change()
On Error GoTo vError

txtAjuste = CCur(txtSF) - CCur(txtSF.Tag)
txtAjuste = Format(CCur(txtAjuste), "Standard")

If CCur(txtAjuste) >= 0 Then
  txtAjuste.ForeColor = vbBlack
Else
  txtAjuste.ForeColor = vbRed
End If

vError:
End Sub

Private Sub txtSF_GotFocus()
On Error GoTo vError
txtSF = CCur(txtSF)
vError:
End Sub

Private Sub txtSF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSM.SetFocus
End Sub

Private Sub txtSF_LostFocus()
On Error GoTo vError
txtSF = Format(CCur(txtSF), "Standard")
vError:
End Sub

Private Sub txtSM_GotFocus()
On Error GoTo vError
txtSM = CCur(txtSM)
vError:
End Sub

Private Sub txtSM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
End Sub

Private Sub txtSM_LostFocus()
On Error GoTo vError
txtSM = Format(CCur(txtSM), "Standard")
vError:
End Sub
