VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIVR_Rec_Titulos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Registro de Inversiones"
   ClientHeight    =   9645
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9000
      Top             =   120
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   10575
      _Version        =   1572864
      _ExtentX        =   18653
      _ExtentY        =   13573
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
      Item(0).Caption =   "Inversión"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbGeneral"
      Item(0).Control(1)=   "GroupBox1"
      Item(1).Caption =   "Adquisición"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gbPago"
      Item(2).Caption =   "Comisiones"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gbComisiones"
      Item(2).Control(1)=   "lswCom"
      Item(3).Caption =   "Flujos"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "lswFlujos_PyD"
      Item(3).Control(1)=   "lswFlujos_Ingresos"
      Item(3).Control(2)=   "scFlujosPyD"
      Item(3).Control(3)=   "scFlujosIngresos"
      Item(3).Control(4)=   "btnExport(0)"
      Item(3).Control(5)=   "btnExport(1)"
      Item(4).Caption =   "Cierres"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "scHistorial"
      Item(4).Control(1)=   "lswCierres"
      Item(4).Control(2)=   "btnExport(2)"
      Item(5).Caption =   "Fondos"
      Item(5).ControlCount=   5
      Item(5).Control(0)=   "lswFi_Mov"
      Item(5).Control(1)=   "ShortcutCaption2"
      Item(5).Control(2)=   "cboFi_Tipo"
      Item(5).Control(3)=   "cboFi_Estado"
      Item(5).Control(4)=   "btnExport(3)"
      Item(6).Caption =   "Cupones"
      Item(6).ControlCount=   4
      Item(6).Control(0)=   "lswCupones"
      Item(6).Control(1)=   "ShortcutCaption4(0)"
      Item(6).Control(2)=   "cboCupones"
      Item(6).Control(3)=   "btnExport(4)"
      Item(7).Caption =   "Asiento"
      Item(7).ControlCount=   7
      Item(7).Control(0)=   "lswAsiento"
      Item(7).Control(1)=   "GroupBox2"
      Item(7).Control(2)=   "btnExport(5)"
      Item(7).Control(3)=   "scAsientos"
      Item(7).Control(4)=   "lswAsientoMain"
      Item(7).Control(5)=   "btnExport(6)"
      Item(7).Control(6)=   "ShortcutCaption4(2)"
      Begin XtremeSuiteControls.ListView lswAsiento 
         Height          =   1932
         Left            =   -69880
         TabIndex        =   161
         Top             =   4320
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   3408
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
         Height          =   6252
         Left            =   -69880
         TabIndex        =   116
         Top             =   720
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   11028
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
         Height          =   6252
         Left            =   -69880
         TabIndex        =   114
         Top             =   720
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   11028
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
         Height          =   6252
         Left            =   -69880
         TabIndex        =   112
         Top             =   720
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   11028
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
      Begin XtremeSuiteControls.ListView lswFlujos_Ingresos 
         Height          =   3012
         Left            =   -69880
         TabIndex        =   68
         Top             =   3960
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   5313
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
         Height          =   2892
         Left            =   -69880
         TabIndex        =   67
         Top             =   720
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ListView lswCom 
         Height          =   5892
         Left            =   -70000
         TabIndex        =   119
         Top             =   1200
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1572864
         _ExtentX        =   18648
         _ExtentY        =   10393
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
      Begin XtremeSuiteControls.ListView lswAsientoMain 
         Height          =   3132
         Left            =   -69880
         TabIndex        =   175
         Top             =   720
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   5524
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   0
         Left            =   -59800
         TabIndex        =   169
         ToolTipText     =   "Exportar a Excel"
         Top             =   396
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   612
         Left            =   -69880
         TabIndex        =   163
         Top             =   6360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   1080
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtA_Debito 
            Height          =   312
            Left            =   2760
            TabIndex        =   164
            Top             =   120
            Width           =   1932
            _Version        =   1572864
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
            TabIndex        =   165
            Top             =   120
            Width           =   1932
            _Version        =   1572864
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
            TabIndex        =   166
            Top             =   120
            Width           =   1932
            _Version        =   1572864
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
            Index           =   1
            Left            =   6480
            TabIndex        =   168
            Top             =   120
            Width           =   1572
            _Version        =   1572864
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   0
            Left            =   960
            TabIndex        =   167
            Top             =   120
            Width           =   1572
            _Version        =   1572864
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
      End
      Begin XtremeSuiteControls.GroupBox gbComisiones 
         Height          =   852
         Left            =   -70000
         TabIndex        =   120
         Top             =   360
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1572864
         _ExtentX        =   18648
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Comisiones"
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
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnComision 
            Height          =   372
            Index           =   1
            Left            =   8400
            TabIndex        =   121
            Top             =   360
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Nuevo"
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
         Begin XtremeSuiteControls.PushButton btnComision 
            Height          =   372
            Index           =   3
            Left            =   9960
            TabIndex        =   122
            Top             =   360
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
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
            Picture         =   "frmIVR_Rec_Titulos.frx":08D1
         End
         Begin XtremeSuiteControls.FlatEdit txtComisionTotal 
            Height          =   312
            Left            =   6120
            TabIndex        =   123
            Top             =   360
            Width           =   2052
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   46
            Left            =   4080
            TabIndex        =   124
            Top             =   360
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Total Comisiones"
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
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPago 
         Height          =   7212
         Left            =   -70000
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1572864
         _ExtentX        =   18648
         _ExtentY        =   12721
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.ListView lswAd 
            Height          =   5172
            Left            =   0
            TabIndex        =   58
            Top             =   1320
            Width           =   10572
            _Version        =   1572864
            _ExtentX        =   18648
            _ExtentY        =   9123
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
         Begin XtremeSuiteControls.PushButton btnAdquisicion 
            Height          =   372
            Index           =   0
            Left            =   8520
            TabIndex        =   62
            Top             =   720
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Nuevo"
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
         Begin XtremeSuiteControls.FlatEdit txtAd_Requerido 
            Height          =   312
            Left            =   1440
            TabIndex        =   59
            Top             =   360
            Width           =   2052
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAd_Registrado 
            Height          =   312
            Left            =   1440
            TabIndex        =   60
            Top             =   720
            Width           =   2052
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAd_Pendiente 
            Height          =   312
            Left            =   5040
            TabIndex        =   61
            Top             =   720
            Width           =   2052
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnAdquisicion 
            Height          =   372
            Index           =   2
            Left            =   10080
            TabIndex        =   66
            Top             =   720
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
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
            Picture         =   "frmIVR_Rec_Titulos.frx":0E75
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   35
            Left            =   3840
            TabIndex        =   65
            Top             =   720
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Pendiente"
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
            Index           =   34
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Registrado"
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
            Index           =   33
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Inversión"
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
      End
      Begin XtremeSuiteControls.GroupBox gbGeneral 
         Height          =   2052
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   10692
         _Version        =   1572864
         _ExtentX        =   18860
         _ExtentY        =   3619
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtEjecutivo 
            Height          =   312
            Left            =   5280
            TabIndex        =   6
            Top             =   240
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboReserva 
            Height          =   312
            Left            =   240
            TabIndex        =   7
            Top             =   1680
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8705
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
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8705
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1032
            Left            =   5280
            TabIndex        =   9
            Top             =   960
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9123
            _ExtentY        =   1820
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
         Begin XtremeSuiteControls.ComboBox cboRecurso 
            Height          =   312
            Left            =   240
            TabIndex        =   127
            Top             =   240
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8705
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
            Index           =   47
            Left            =   240
            TabIndex        =   126
            Top             =   0
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Recurso"
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
            Index           =   5
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Portafolio"
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
            Left            =   5280
            TabIndex        =   12
            Top             =   0
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Ejecutivo"
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
            Index           =   7
            Left            =   240
            TabIndex        =   11
            Top             =   1440
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Reserva"
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
            Index           =   25
            Left            =   5280
            TabIndex        =   10
            Top             =   720
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas"
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   5535
         Left            =   0
         TabIndex        =   14
         Top             =   2400
         Width           =   10695
         _Version        =   1572864
         _ExtentX        =   18865
         _ExtentY        =   9763
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.GroupBox gbCDP 
            Height          =   4095
            Left            =   2400
            TabIndex        =   147
            Top             =   5040
            Visible         =   0   'False
            Width           =   10575
            _Version        =   1572864
            _ExtentX        =   18653
            _ExtentY        =   7223
            _StockProps     =   79
            Caption         =   "CDPS:"
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
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   252
               Left            =   8280
               TabIndex        =   148
               Top             =   840
               Width           =   2772
               _Version        =   1572864
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Capitaliza Intereses?"
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
            Begin XtremeSuiteControls.DateTimePicker dtpC_Compra 
               Height          =   312
               Left            =   6720
               TabIndex        =   149
               Top             =   360
               Width           =   1332
               _Version        =   1572864
               _ExtentX        =   2350
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
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   312
               Left            =   6000
               TabIndex        =   150
               Top             =   2040
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   312
               Left            =   6000
               TabIndex        =   151
               Top             =   1680
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   312
               Left            =   6000
               TabIndex        =   152
               Top             =   1200
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboC_BaseCalculo 
               Height          =   312
               Left            =   6000
               TabIndex        =   153
               Top             =   840
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3625
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
               Index           =   57
               Left            =   4680
               TabIndex        =   158
               Top             =   360
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Fecha Compra"
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
               Index           =   56
               Left            =   3960
               TabIndex        =   157
               Top             =   2040
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "No. de Participaciones"
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
               Index           =   55
               Left            =   3960
               TabIndex        =   156
               Top             =   1680
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Valor de la Participación"
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
               Index           =   43
               Left            =   3960
               TabIndex        =   155
               Top             =   1200
               Width           =   1932
               _Version        =   1572864
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Monto de la Inversión"
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
               Index           =   24
               Left            =   3960
               TabIndex        =   154
               Top             =   840
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Base Cálculo"
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
         End
         Begin XtremeSuiteControls.GroupBox gbFondos 
            Height          =   4095
            Left            =   0
            TabIndex        =   72
            Top             =   5040
            Visible         =   0   'False
            Width           =   10575
            _Version        =   1572864
            _ExtentX        =   18653
            _ExtentY        =   7223
            _StockProps     =   79
            Caption         =   "Fondos de Inversion:"
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
            Begin XtremeSuiteControls.CheckBox chkFi_CapiltalizaInt 
               Height          =   252
               Left            =   8280
               TabIndex        =   103
               Top             =   840
               Width           =   2772
               _Version        =   1572864
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Capitaliza Intereses?"
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
            Begin XtremeSuiteControls.DateTimePicker dtpFi_FechaCompra 
               Height          =   312
               Left            =   6720
               TabIndex        =   83
               Top             =   360
               Width           =   1332
               _Version        =   1572864
               _ExtentX        =   2350
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
            Begin XtremeSuiteControls.FlatEdit txtFi_ParticipacionNo 
               Height          =   312
               Left            =   6000
               TabIndex        =   85
               Top             =   2040
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtFi_ParticipacionValor 
               Height          =   312
               Left            =   6000
               TabIndex        =   86
               Top             =   1680
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtFi_Inversion 
               Height          =   312
               Left            =   6000
               TabIndex        =   89
               Top             =   1200
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboFi_BaseCalculo 
               Height          =   312
               Left            =   6000
               TabIndex        =   101
               Top             =   840
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3625
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
               Index           =   45
               Left            =   3960
               TabIndex        =   102
               Top             =   840
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Base Cálculo"
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
               Index           =   44
               Left            =   3960
               TabIndex        =   90
               Top             =   1200
               Width           =   1932
               _Version        =   1572864
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Monto de la Inversión"
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
               Index           =   42
               Left            =   3960
               TabIndex        =   88
               Top             =   1680
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Valor de la Participación"
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
               Index           =   41
               Left            =   3960
               TabIndex        =   87
               Top             =   2040
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "No. de Participaciones"
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
               Index           =   40
               Left            =   4680
               TabIndex        =   84
               Top             =   360
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Fecha Compra"
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
         End
         Begin XtremeSuiteControls.GroupBox gbAcciones 
            Height          =   4215
            Left            =   10440
            TabIndex        =   71
            Top             =   840
            Visible         =   0   'False
            Width           =   10455
            _Version        =   1572864
            _ExtentX        =   18436
            _ExtentY        =   7429
            _StockProps     =   79
            Caption         =   "Acciones:"
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
            Begin XtremeSuiteControls.DateTimePicker dtpAc_FechaCompra 
               Height          =   312
               Left            =   6000
               TabIndex        =   73
               Top             =   360
               Width           =   1332
               _Version        =   1572864
               _ExtentX        =   2350
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
            Begin XtremeSuiteControls.FlatEdit txtAc_NoAcciones 
               Height          =   312
               Left            =   6000
               TabIndex        =   75
               Top             =   1200
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAc_ValorAccion 
               Height          =   312
               Left            =   6000
               TabIndex        =   76
               Top             =   1560
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAc_ValorTransado 
               Height          =   312
               Left            =   6000
               TabIndex        =   79
               Top             =   2160
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAc_ValorActual 
               Height          =   312
               Left            =   6000
               TabIndex        =   80
               Top             =   2520
               Width           =   2052
               _Version        =   1572864
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboAc_BaseCalculo 
               Height          =   312
               Left            =   6000
               TabIndex        =   130
               Top             =   840
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3625
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
               Index           =   50
               Left            =   4200
               TabIndex        =   131
               Top             =   840
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Base Cálculo"
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
               Index           =   39
               Left            =   4200
               TabIndex        =   82
               Top             =   2520
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Valor Real"
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
               Index           =   38
               Left            =   4200
               TabIndex        =   81
               Top             =   2160
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Monto Transado"
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
               Index           =   37
               Left            =   4200
               TabIndex        =   78
               Top             =   1560
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Valor por Acción"
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
               Index           =   36
               Left            =   4200
               TabIndex        =   77
               Top             =   1200
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "No. de Acciones"
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
               Index           =   32
               Left            =   4200
               TabIndex        =   74
               Top             =   360
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Fecha Compra"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtPyD 
            Height          =   312
            Left            =   6120
            TabIndex        =   15
            Top             =   2400
            Width           =   1932
            _Version        =   1572864
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaVence 
            Height          =   312
            Left            =   1920
            TabIndex        =   16
            Top             =   2760
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   312
            Left            =   1920
            TabIndex        =   33
            Top             =   960
            Width           =   1812
            _Version        =   1572864
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
            Left            =   1920
            TabIndex        =   17
            Top             =   1320
            Width           =   1812
            _Version        =   1572864
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
            Left            =   1920
            TabIndex        =   18
            Top             =   1680
            Width           =   1812
            _Version        =   1572864
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
            Left            =   1920
            TabIndex        =   19
            Top             =   2400
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
            Height          =   312
            Left            =   6120
            TabIndex        =   35
            Top             =   960
            Width           =   1932
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtValorFacial 
            Height          =   312
            Left            =   6120
            TabIndex        =   37
            Top             =   1320
            Width           =   1932
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtValorTransado 
            Height          =   315
            Left            =   8280
            TabIndex        =   20
            Top             =   2040
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtComisionBNV 
            Height          =   315
            Left            =   6120
            TabIndex        =   21
            Top             =   3480
            Width           =   1935
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtComisionAdm 
            Height          =   315
            Left            =   6120
            TabIndex        =   22
            Top             =   3840
            Width           =   1935
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtDiasAcumulados 
            Height          =   312
            Left            =   1920
            TabIndex        =   25
            Top             =   3840
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDiasVencimiento 
            Height          =   312
            Left            =   1920
            TabIndex        =   26
            Top             =   4200
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFechaUltPago 
            Height          =   312
            Left            =   1920
            TabIndex        =   27
            Top             =   3480
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIntAcumulado 
            Height          =   312
            Left            =   6120
            TabIndex        =   28
            Top             =   2760
            Width           =   1932
            _Version        =   1572864
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtOperacion 
            Height          =   312
            Left            =   0
            TabIndex        =   29
            Top             =   480
            Width           =   2652
            _Version        =   1572864
            _ExtentX        =   4678
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
            Left            =   2640
            TabIndex        =   30
            Top             =   480
            Width           =   2652
            _Version        =   1572864
            _ExtentX        =   4678
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
         Begin XtremeSuiteControls.FlatEdit txtPrecio 
            Height          =   312
            Left            =   6120
            TabIndex        =   41
            Top             =   2040
            Width           =   1932
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtTasaNominal 
            Height          =   312
            Left            =   6120
            TabIndex        =   39
            Top             =   1680
            Width           =   1932
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtCostoNeto 
            Height          =   315
            Left            =   8280
            TabIndex        =   128
            ToolTipText     =   "Costo Neto"
            Top             =   1320
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtISIN 
            Height          =   312
            Left            =   5280
            TabIndex        =   31
            Top             =   480
            Width           =   2652
            _Version        =   1572864
            _ExtentX        =   4678
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
            Left            =   7920
            TabIndex        =   32
            Top             =   480
            Width           =   2652
            _Version        =   1572864
            _ExtentX        =   4678
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
         Begin XtremeSuiteControls.FlatEdit txtIVA 
            Height          =   315
            Left            =   6120
            TabIndex        =   141
            ToolTipText     =   "Costo Neto"
            Top             =   4200
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtLiquidacion 
            Height          =   315
            Left            =   8280
            TabIndex        =   143
            ToolTipText     =   "Costo Neto"
            Top             =   2760
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRendNominal 
            Height          =   315
            Left            =   9360
            TabIndex        =   179
            Top             =   3480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRendNeto 
            Height          =   315
            Left            =   8280
            TabIndex        =   180
            Top             =   3480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTEA 
            Height          =   315
            Left            =   9360
            TabIndex        =   23
            Top             =   4200
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTIR 
            Height          =   315
            Left            =   8280
            TabIndex        =   24
            Top             =   4200
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnTablaPyD 
            Height          =   255
            Left            =   9360
            TabIndex        =   182
            Top             =   3960
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "TEA:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   0
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtISRMonto 
            Height          =   315
            Left            =   6120
            TabIndex        =   183
            ToolTipText     =   "Costo Neto"
            Top             =   4560
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   23
            Left            =   4320
            TabIndex        =   184
            Top             =   4560
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ISR Monto:"
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
            Index           =   60
            Left            =   8280
            TabIndex        =   181
            Top             =   3240
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Rend. Neto / Nominal:"
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
            Index           =   58
            Left            =   8280
            TabIndex        =   178
            Top             =   3960
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "TIR:"
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
            Index           =   54
            Left            =   8280
            TabIndex        =   144
            Top             =   2520
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monto a Liquidar:"
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
            Index           =   53
            Left            =   4320
            TabIndex        =   142
            Top             =   4200
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "IVA Monto:"
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
            Index           =   52
            Left            =   8280
            TabIndex        =   140
            Top             =   1080
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Costo Neto:"
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
            Index           =   51
            Left            =   7920
            TabIndex        =   139
            Top             =   240
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "CupIp"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   49
            Left            =   5400
            TabIndex        =   138
            Top             =   240
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "ISIN"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   56
            Top             =   960
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Divisa"
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
            TabIndex        =   55
            Top             =   1320
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Periodicidad"
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
            Index           =   10
            Left            =   0
            TabIndex        =   54
            Top             =   240
            Width           =   2652
            _Version        =   1572864
            _ExtentX        =   4678
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. Operación"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   11
            Left            =   2760
            TabIndex        =   53
            Top             =   240
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "No. Serie"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   12
            Left            =   240
            TabIndex        =   52
            Top             =   1680
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Base Cálculo"
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
            TabIndex        =   51
            Top             =   2400
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha Compra"
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
            Index           =   14
            Left            =   240
            TabIndex        =   50
            Top             =   2760
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha Vencimiento"
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
            Index           =   16
            Left            =   4320
            TabIndex        =   49
            Top             =   960
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo de Cambio"
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
            Index           =   17
            Left            =   4320
            TabIndex        =   48
            Top             =   1320
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Valor Facial"
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
            Index           =   18
            Left            =   8280
            TabIndex        =   47
            Top             =   1800
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Valor Transado:"
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
            Index           =   19
            Left            =   4320
            TabIndex        =   46
            Top             =   3480
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Comisión BNV"
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
            Index           =   20
            Left            =   4320
            TabIndex        =   45
            Top             =   3840
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Comisión Adm"
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
            Index           =   22
            Left            =   4320
            TabIndex        =   44
            Top             =   1680
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tasa Facial"
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
            Index           =   21
            Left            =   4320
            TabIndex        =   43
            Top             =   2040
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Precio"
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
            Index           =   27
            Left            =   240
            TabIndex        =   42
            Top             =   3840
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Dias Acumulados"
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
            Index           =   28
            Left            =   240
            TabIndex        =   40
            Top             =   4200
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Dias al Vencimiento"
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
            Index           =   29
            Left            =   240
            TabIndex        =   38
            Top             =   3480
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fec. Ultimo Pago"
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
            Index           =   30
            Left            =   4320
            TabIndex        =   36
            Top             =   2400
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Prima / Descuento"
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
            Index           =   31
            Left            =   4320
            TabIndex        =   34
            Top             =   2760
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Int. Acumulados"
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
      End
      Begin XtremeSuiteControls.ComboBox cboCupones 
         Height          =   312
         Left            =   -61720
         TabIndex        =   129
         Top             =   396
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
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
         Left            =   -63400
         TabIndex        =   159
         Top             =   396
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
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
         Left            =   -61600
         TabIndex        =   160
         Top             =   396
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   1
         Left            =   -59800
         TabIndex        =   170
         ToolTipText     =   "Exportar a Excel"
         Top             =   3640
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":1419
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   2
         Left            =   -59800
         TabIndex        =   171
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":1CEA
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   3
         Left            =   -59800
         TabIndex        =   172
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":25BB
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   4
         Left            =   -59800
         TabIndex        =   173
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":2E8C
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   5
         Left            =   -59800
         TabIndex        =   174
         ToolTipText     =   "Exportar a Excel"
         Top             =   3960
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":375D
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   252
         Index           =   6
         Left            =   -59800
         TabIndex        =   176
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Visible         =   0   'False
         Width           =   252
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmIVR_Rec_Titulos.frx":402E
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   372
         Index           =   2
         Left            =   -69880
         TabIndex        =   177
         Top             =   360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Asientos Relacionados"
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
      Begin XtremeShortcutBar.ShortcutCaption scAsientos 
         Height          =   372
         Left            =   -69880
         TabIndex        =   162
         Top             =   3960
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Detalle del Asiento"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   117
         Top             =   360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Cupones"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   372
         Left            =   -69880
         TabIndex        =   115
         Top             =   360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Movimientos"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption scHistorial 
         Height          =   372
         Left            =   -69880
         TabIndex        =   113
         Top             =   360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Cierres"
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
      Begin XtremeShortcutBar.ShortcutCaption scFlujosIngresos 
         Height          =   372
         Left            =   -69880
         TabIndex        =   70
         Top             =   3600
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Ingresos"
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
      Begin XtremeShortcutBar.ShortcutCaption scFlujosPyD 
         Height          =   372
         Left            =   -69880
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Amortización de Primas / Descuentos"
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
   Begin XtremeSuiteControls.GroupBox gbToolAccion 
      Height          =   8895
      Left            =   10800
      TabIndex        =   104
      Top             =   720
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   15690
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   106
         Top             =   3720
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Registro de Cupón"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   1
         Left            =   120
         TabIndex        =   107
         Top             =   4200
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Valorizar"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   2
         Left            =   120
         TabIndex        =   108
         Top             =   4680
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Liquidar"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   4
         Left            =   120
         TabIndex        =   109
         Top             =   6000
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aportaciones"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   5
         Left            =   120
         TabIndex        =   110
         Top             =   6480
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Retiros"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   6
         Left            =   120
         TabIndex        =   111
         Top             =   6960
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Actualiza Participación"
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
      Begin XtremeSuiteControls.PushButton btnToolAccion 
         Height          =   372
         Index           =   3
         Left            =   120
         TabIndex        =   118
         Top             =   5160
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cambio de Tasas"
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
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   372
         Index           =   1
         Left            =   120
         TabIndex        =   132
         Top             =   2040
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmIVR_Rec_Titulos.frx":48FF
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   133
         Top             =   1560
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmIVR_Rec_Titulos.frx":5030
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   372
         Index           =   2
         Left            =   120
         TabIndex        =   134
         Top             =   2520
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Activar"
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
         Picture         =   "frmIVR_Rec_Titulos.frx":5662
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnTool 
         Height          =   372
         Index           =   3
         Left            =   120
         TabIndex        =   135
         Top             =   3000
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Boleta"
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
         Picture         =   "frmIVR_Rec_Titulos.frx":5D89
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtValorLibros 
         Height          =   312
         Left            =   120
         TabIndex        =   145
         Top             =   8160
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   15
         Left            =   120
         TabIndex        =   146
         Top             =   7920
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor Libros:"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   1092
         Left            =   -120
         TabIndex        =   105
         Top             =   240
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   1926
         _StockProps     =   14
         Caption         =   "Acciones"
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
         VisualTheme     =   6
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox gbMain 
      Height          =   1332
      Left            =   0
      TabIndex        =   92
      Top             =   720
      Width           =   10812
      _Version        =   1572864
      _ExtentX        =   19071
      _ExtentY        =   2350
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboEmisor 
         Height          =   312
         Left            =   5400
         TabIndex        =   93
         Top             =   360
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8916
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
         Left            =   360
         TabIndex        =   94
         Top             =   960
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.ComboBox cboInstrumento 
         Height          =   312
         Left            =   360
         TabIndex        =   95
         Top             =   360
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8916
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
         Left            =   5400
         TabIndex        =   96
         Top             =   960
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8916
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
         Index           =   1
         Left            =   360
         TabIndex        =   100
         Top             =   720
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Administrador"
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
         Left            =   360
         TabIndex        =   99
         Top             =   120
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Instrumento"
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
         Left            =   5400
         TabIndex        =   98
         Top             =   120
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Emisor"
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
         Left            =   5400
         TabIndex        =   97
         Top             =   720
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Clasificación"
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
   Begin XtremeSuiteControls.FlatEdit txtInversionId 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   120
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCartaIB 
      Height          =   492
      Left            =   10800
      TabIndex        =   3
      Top             =   120
      Width           =   2172
      _Version        =   1572864
      _ExtentX        =   3831
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3840
      TabIndex        =   91
      Top             =   120
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   492
      Left            =   6120
      TabIndex        =   125
      Top             =   120
      Width           =   2772
      _Version        =   1572864
      _ExtentX        =   4890
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   48
      Left            =   4800
      TabIndex        =   136
      Top             =   120
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   26
      Left            =   9120
      TabIndex        =   2
      Top             =   120
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. IB Ref:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Inversión:"
      BackColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
      Height          =   735
      Left            =   0
      TabIndex        =   137
      Top             =   0
      Width           =   13455
      _Version        =   1572864
      _ExtentX        =   23733
      _ExtentY        =   1296
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
      VisualTheme     =   6
      Alignment       =   1
   End
End
Attribute VB_Name = "frmIVR_Rec_Titulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, vScroll As Boolean
Dim itmX As ListViewItem, vFecha As Date, vPeriodicidad As Integer
Dim vDivisaLocaL As String, vToken As String
Dim mTituloIdConExt As Long


Private Sub btnAdquisicion_Click(Index As Integer)

If txtEstado.Tag <> "S" Then Exit Sub


gIVR_Transito.TituloId = txtInversionId.Text
gIVR_Transito.Codigo = txtInversionId.Text
gIVR_Transito.Concepto = "T"
gIVR_Transito.Tipo = "T"

gIVR_Transito.Monto = CCur(txtAd_Requerido.Text)
gIVR_Transito.TipoMov = "C"

gIVR_Transito.Divisa = cboDivisa.ItemData(cboDivisa.ListIndex)
gIVR_Transito.TipoCambio = CCur(txtTipoCambio.Text)

gIVR_Transito.Descripcion = "Id.:" & txtInversionId.Text & ", Operación: " & txtOperacion.Text

Select Case Index
    Case 0 'Nuevo
        
        frmIVR_Rec_Bancos_Registro.Show vbModal
        
        
    Case 2 'Eliminar
        
        Dim i As Integer
        With lswAd.ListItems
            For i = 1 To .Count
                If .Item(i).Checked = True Then
                    strSQL = "delete  IVR_TRANSACCIONES Where TRANSAC_ID = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                End If
            Next i
        End With
        
End Select

Call sbAdquisicion_Load

End Sub

Private Sub btnExport_Click(Index As Integer)

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
    Case 0 'Primas/Descuentos
        Call Excel_Exportar_Lsw(lswFlujos_PyD)
    Case 1 'Flujo Intereses
        Call Excel_Exportar_Lsw(lswFlujos_Ingresos)
    Case 2 'Cierres
        Call Excel_Exportar_Lsw(lswCierres)
    Case 3 'Fondos Mov
        Call Excel_Exportar_Lsw(lswFi_Mov)
    Case 4 'Cupones
        Call Excel_Exportar_Lsw(lswCupones)
    Case 5 'Asiento Detalle
        Call Excel_Exportar_Lsw(lswAsiento)
    Case 6 'Asiento Main
        Call Excel_Exportar_Lsw(lswAsientoMain)
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnTablaPyD_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

If Mid(txtEstado.Text, 1, 1) = "A" Then
    
    strSQL = "exec spIvr_Titulo_PyD_Recalculo_TEA " & txtInversionId.Text & ", " & txtTEA.Text & ", '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)

End If

Me.MousePointer = vbDefault

MsgBox "Flujos de Primas y Descuentos Actualizados!", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnTool_Click(Index As Integer)

Select Case Index
    Case 0 'Nuevo
        Call sbInicializa
    
    Case 1 'Guardar
        If Mid(txtEstado.Text, 1, 1) = "S" Then
            Call sbGuardar
        Else
            MsgBox "Solo se puede Guardar en Estado de Solicitud!", vbExclamation
        End If
    
    Case 2 'Activar
        If Mid(txtEstado.Text, 1, 1) = "S" Then
        
            Call sbAdquisicion_Load
            
            If CCur(txtAd_Pendiente.Text) = 0 Then
                Call sbActivacion
            Else
                
                MsgBox "El monto de adquisición se encuentra pendiente!", vbExclamation
            
            End If
        Else
            MsgBox "Solo se puede Activar en Estado de Solicitud!", vbInformation
        End If
       
    Case 3 'Boleta de Registro
        Call sbBoleta
End Select

End Sub

Private Sub btnToolAccion_Click(Index As Integer)


If Mid(txtEstado.Text, 1, 1) <> "A" Then
    MsgBox "Consulte una Inversión que se encuentre activa!", vbInformation
    Exit Sub
End If


gIVR_Transito.TituloId = txtInversionId.Text
gIVR_Transito.Codigo = txtInversionId.Text
gIVR_Transito.Divisa = cboDivisa.ItemData(cboDivisa.ListIndex)
gIVR_Transito.TipoCambio = CCur(txtTipoCambio.Text)
gIVR_Transito.Descripcion = "Id.:" & txtInversionId.Text & ", Operación: " & txtOperacion.Text



Select Case Index
    Case 0 'Cupones
        
        gIVR_Transito.Concepto = "Cupon"
        gIVR_Transito.Tipo = "T"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "D"
        
        frmIVR_Proc_Cupones.Show vbModal
    
    
    Case 1 'Valorizar
        frmIVR_Proc_Valorizacion.Show
    
    Case 2 'Liquidar
        
        gIVR_Transito.Concepto = "LIQ"
        gIVR_Transito.Tipo = "T"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "D"
        
        frmIVR_Proc_Liquidacion.Show vbModal
    
    
    Case 3 'Cambio de Tasas
        frmIVR_Proc_Cambio_Tasas.Show vbModal
        
    Case 4 'Aportaciones
        
        gIVR_Transito.Concepto = "FI_APO"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        
        gIVR_Transito.TipoMov = "C"
        
        frmIVR_Proc_Fondos_Mov.Show vbModal
    
    Case 5 'Retiros
        gIVR_Transito.Concepto = "FI_RET"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "D"
        
        frmIVR_Proc_Fondos_Mov.Show vbModal
        
    Case 6 'Participaciones
        gIVR_Transito.Concepto = "FI_PART"
        gIVR_Transito.Tipo = "F"
        gIVR_Transito.Monto = 0
        gIVR_Transito.TipoMov = "C"
        
        frmIVR_Proc_Fondos_Participacion.Show vbModal
    
End Select

Call sbConsulta(txtInversionId.Text)

End Sub


Private Sub sbBoleta()
Dim strSQL As String

On Error GoTo vPrintError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowTitle = "SCGI: Boleta de Registro de Inversión"
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    
    .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fxUsuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(3) = "fxFecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT

    .ReportFileName = SIFGlobal.fxPathReportes("IVR_Boleta_Titulo.rpt")
    strSQL = "{vIVR_INVERSIONES.TITULO_ID} = '" & txtInversionId.Text & "'"
        
     .SelectionFormula = strSQL
    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub
 
vPrintError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbActivacion()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spIVR_TITULOS_ACTIVA " & txtInversionId.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbConsulta(txtInversionId.Text)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub


Private Sub cboBaseCalculo_Click()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub cboCupones_Click()
If vPaso Then Exit Sub

Call sbIVR_Cupones_Load(lswCupones, txtInversionId.Text, Mid(cboCupones.Text, 1, 1))

End Sub

Private Sub cboDivisa_Click()
If vPaso Then Exit Sub

If cboDivisa.ItemData(cboDivisa.ListIndex) = vDivisaLocaL Then
    txtTipoCambio.Text = "1"
    txtTipoCambio.Locked = True
Else
    txtTipoCambio.Locked = False
End If

End Sub

Private Sub cboFi_Estado_Click()
If vPaso Then Exit Sub

Call sbIVR_Fondos_Load(lswFi_Mov, txtInversionId.Text, cboFi_Estado.Text, cboFi_Tipo.Text)

End Sub

Private Sub cboFi_Tipo_Click()
If vPaso Then Exit Sub

Call sbIVR_Fondos_Load(lswFi_Mov, txtInversionId.Text, cboFi_Estado.Text, cboFi_Tipo.Text)

End Sub

Private Sub cboInstrumento_Click()
Dim pLeft As Integer, pTop As Integer

If vPaso Then Exit Sub

pLeft = 120
pTop = 1320

gbFondos.BorderStyle = xtpFrameNone
gbAcciones.BorderStyle = xtpFrameNone

gbFondos.Visible = False
gbAcciones.Visible = False

gbFondos.Top = pTop
gbFondos.Left = pLeft

gbAcciones.Top = pTop
gbAcciones.Left = pLeft


btnToolAccion.Item(0).Enabled = False
btnToolAccion.Item(1).Enabled = False
btnToolAccion.Item(2).Enabled = False
btnToolAccion.Item(3).Enabled = False
btnToolAccion.Item(4).Enabled = False
btnToolAccion.Item(5).Enabled = False
btnToolAccion.Item(6).Enabled = False


txtValorFacial.BackColor = RGB(187, 215, 247)
txtTasaNominal.BackColor = RGB(187, 215, 247)
txtPrecio.BackColor = RGB(187, 215, 247)
txtComisionAdm.BackColor = RGB(187, 215, 247)
txtComisionBNV.BackColor = RGB(187, 215, 247)
txtIVA.BackColor = RGB(187, 215, 247)

Select Case Mid(cboInstrumento.ItemData(cboInstrumento.ListIndex), 1, 1)
    Case "T"
        btnToolAccion.Item(0).Enabled = True 'Cupon
        btnToolAccion.Item(1).Enabled = True 'Valorizar
        btnToolAccion.Item(2).Enabled = True 'Liquidar
        btnToolAccion.Item(3).Enabled = True 'Cambio de Tasas
    Case "F"
        gbFondos.Visible = True

        btnToolAccion.Item(4).Enabled = True
        btnToolAccion.Item(5).Enabled = True
        btnToolAccion.Item(6).Enabled = True
    
    
    Case "A"
        gbAcciones.Visible = True

        btnToolAccion.Item(4).Enabled = True
        btnToolAccion.Item(5).Enabled = True
        btnToolAccion.Item(6).Enabled = True

    Case "D"
        btnToolAccion.Item(3).Enabled = True 'Cambio de Tasas

End Select

End Sub

Private Sub cboPeriodicidad_Click()

If vPaso Then Exit Sub

vPeriodicidad = fxIVR_Periodicidad(cboPeriodicidad.ItemData(cboPeriodicidad.ListIndex))

    Call sbTitulo_Cal_Refresh


End Sub

Private Sub dtpFechaCompra_Change()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub dtpFechaVence_Change()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub FlatScrollBar_Change()
On Error GoTo vError

If vScroll Then

    If txtInversionId.Text = "0" And FlatScrollBar.Value = 0 Then
       txtInversionId.Text = "99999999"
    End If

    strSQL = "select Top 1 TITULO_ID from vIVR_INVERSIONES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where TITULO_ID > " & txtInversionId.Text & " order by TITULO_ID asc"
    Else
       strSQL = strSQL & " where TITULO_ID < " & txtInversionId.Text & " order by TITULO_ID desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtInversionId.Text = rs!TITULO_ID
      Call sbConsulta(txtInversionId.Text)
      
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



Private Sub sbCDP_Cal_Refresh()
Dim pBase As Integer

On Error GoTo vError

If Not IsNumeric(txtFi_Inversion.Text) Then
    txtFi_Inversion.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionNo.Text) Then
    txtFi_ParticipacionNo.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionValor.Text) Then
    txtFi_ParticipacionValor.Text = "1"
End If


pBase = cboFi_BaseCalculo.ItemData(cboFi_BaseCalculo.ListIndex)


'Formato

txtFi_ParticipacionNo.Text = Format(CCur(txtFi_Inversion.Text) / CDbl(txtFi_ParticipacionValor), "0.0000000000")
txtFi_ParticipacionValor.Text = Format(CDbl(txtFi_ParticipacionValor), "Standard")
txtFi_Inversion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

txtLiquidacion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

Exit Sub

vError:

End Sub

Private Sub sbAcciones_Cal_Refresh()
Dim pBase As Integer

On Error GoTo vError

If Not IsNumeric(txtFi_Inversion.Text) Then
    txtFi_Inversion.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionNo.Text) Then
    txtFi_ParticipacionNo.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionValor.Text) Then
    txtFi_ParticipacionValor.Text = "1"
End If


pBase = cboFi_BaseCalculo.ItemData(cboFi_BaseCalculo.ListIndex)


'Formato

txtFi_ParticipacionNo.Text = Format(CCur(txtFi_Inversion.Text) / CDbl(txtFi_ParticipacionValor), "0.0000000000")
txtFi_ParticipacionValor.Text = Format(CDbl(txtFi_ParticipacionValor), "Standard")
txtFi_Inversion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

txtLiquidacion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

Exit Sub

vError:

End Sub



Private Sub sbFondos_Cal_Refresh()
Dim pBase As Integer

On Error GoTo vError

If Not IsNumeric(txtFi_Inversion.Text) Then
    txtFi_Inversion.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionNo.Text) Then
    txtFi_ParticipacionNo.Text = "0"
End If

If Not IsNumeric(txtFi_ParticipacionValor.Text) Then
    txtFi_ParticipacionValor.Text = "1"
End If


pBase = cboFi_BaseCalculo.ItemData(cboFi_BaseCalculo.ListIndex)


'Formato

txtFi_ParticipacionNo.Text = Format(CCur(txtFi_Inversion.Text) / CDbl(txtFi_ParticipacionValor), "###,###,###,##0.0000000000")
txtFi_ParticipacionValor.Text = Format(CDbl(txtFi_ParticipacionValor), "Standard")
txtFi_Inversion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

txtLiquidacion.Text = Format(CCur(txtFi_Inversion.Text), "Standard")

Exit Sub

vError:

End Sub

Private Sub sbTitulo_Cal_Refresh()

If vPaso Then Exit Sub

Dim pCuponUlt As Date, pAcumMnt As Double, pTasaNominal As Double
Dim pBase As Integer

On Error GoTo vError

If Not IsNumeric(txtIVA.Text) Then
    txtIVA.Text = "0"
End If

If Not IsNumeric(txtTasaNominal.Text) Then
    txtTasaNominal.Text = "0"
End If

If Not IsNumeric(txtPrecio.Text) Then
    txtPrecio.Text = "100"
End If

If Not IsNumeric(txtValorFacial.Text) Then
    txtValorFacial.Text = "0"
End If

If Not IsNumeric(txtComisionAdm.Text) Then
    txtComisionAdm.Text = "0"
End If
If Not IsNumeric(txtComisionBNV.Text) Then
    txtComisionBNV.Text = "0"
End If

pBase = cboBaseCalculo.ItemData(cboBaseCalculo.ListIndex)

If vPeriodicidad = 99 Then
  txtPrecio.Text = "100"
  txtPrecio.Locked = True
  txtPrecio.BackColor = vbWhite
Else
  txtPrecio.Locked = False
  txtPrecio.BackColor = txtValorFacial.BackColor
End If

If vPeriodicidad = 0 Then
    txtTasaNominal.Text = "0"
    txtTasaNominal.Locked = True
Else
    txtTasaNominal.Locked = False
End If

If vPeriodicidad = 0 Or vPeriodicidad = 99 Then

    If pBase = 360 Then
        txtDiasAcumulados.Text = "0"
        txtDiasVencimiento.Text = fxFi_Days360(dtpFechaCompra.Value, dtpFechaVence.Value)
    Else
        txtDiasAcumulados.Text = "0"
        txtDiasVencimiento.Text = DateDiff("d", dtpFechaCompra.Value, dtpFechaVence.Value)
    End If

Else

    pCuponUlt = fxFi_Cupon_Pago_Ultimo(dtpFechaCompra.Value, dtpFechaVence.Value, vPeriodicidad)
    txtFechaUltPago.Text = Format(pCuponUlt, "yyyy-mm-dd")
    
    
    If pBase = 360 Then
        txtDiasAcumulados.Text = fxFi_Days360(pCuponUlt, dtpFechaCompra.Value)
        txtDiasVencimiento.Text = fxFi_Days360(dtpFechaCompra.Value, dtpFechaVence.Value)
    Else
        txtDiasAcumulados.Text = DateDiff("d", pCuponUlt, dtpFechaCompra.Value)
        txtDiasVencimiento.Text = DateDiff("d", dtpFechaCompra.Value, dtpFechaVence.Value)
    End If

End If

pTasaNominal = CDbl(txtTasaNominal.Text) / 100

pTasaNominal = pTasaNominal / pBase

pAcumMnt = CCur(txtValorFacial.Text) * CLng(txtDiasAcumulados.Text) * pTasaNominal

txtIntAcumulado.Text = Format(pAcumMnt, "Standard")

Dim vTransado As Double

vTransado = CCur(txtValorFacial.Text) * (CDbl(txtPrecio.Text) / 100)

'vTransado = vTransado + CCur(txtComisionAdm.Text) + CCur(txtComisionBNV.Text) + pAcumMnt
vTransado = vTransado + pAcumMnt


If vPeriodicidad = 99 Then
    txtPyD.Text = Format(0, "Standard")
Else
    txtPyD.Text = Format(vTransado - pAcumMnt - CCur(txtValorFacial.Text), "Standard")
End If


txtCostoNeto.Text = Format(vTransado - pAcumMnt, "Standard")


'txtPrecio.Text = Format((CCur(txtValorTransado) / CCur(txtValorFacial.Text)) * 100, "0.000000")


'Formato
txtIVA.Text = Format(CCur(txtIVA.Text), "Standard")

txtValorFacial.Text = Format(CCur(txtValorFacial.Text), "Standard")
txtValorTransado.Text = Format(vTransado, "Standard")
txtLiquidacion.Text = Format(vTransado + CCur(txtComisionAdm.Text) + CCur(txtComisionBNV.Text) + CCur(txtIVA.Text), "Standard")

txtComisionAdm.Text = Format(CCur(txtComisionAdm.Text), "Standard")
txtComisionBNV.Text = Format(CCur(txtComisionBNV.Text), "Standard")


Exit Sub

vError:

End Sub


Public Sub sbConsulta_Externa(pTituloId As Long)

mTituloIdConExt = pTituloId

End Sub


Private Sub Form_Load()

mTituloIdConExt = 0
vFecha = fxFechaServidor

vPaso = True

cboCupones.Clear
cboCupones.AddItem "Registrados"
cboCupones.AddItem "Proyectados"
cboCupones.Text = "Registrados"


vScroll = False
FlatScrollBar.Value = 0
vScroll = True

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

With lswFi_Mov.ColumnHeaders
    .Clear
    .Add , , "Transac Id", 1100
    .Add , , "Est", 300, vbCenter
    .Add , , "Seq Id ", 700, vbCenter
    .Add , , "Tipo Mov", 1100, vbCenter
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Intereses", 1800, vbRightJustify
    .Add , , "Participa. No", 2500, vbRightJustify
    .Add , , "Participa. Valor", 2500, vbRightJustify
End With

With lswCupones.ColumnHeaders
    .Clear
    .Add , , "Fecha", 1200
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Monto", 2100, vbRightJustify
    .Add , , "Int.Acum.", 2100, vbRightJustify
    .Add , , "Int.Pend.", 2100, vbRightJustify
    .Add , , "Documento", 2100, vbCenter
    .Add , , "Transac Id", 1200
    .Add , , "T. Fecha", 1200, vbCenter
    .Add , , "T. Usuario", 2100, vbCenter
    .Add , , "T. Notas", 2100
End With

With lswFlujos_Ingresos.ColumnHeaders
    .Clear
    .Add , , "Corte", 1400
    .Add , , "Dias", 1000, vbCenter
    .Add , , "Int. Cal.", 2100, vbRightJustify
    .Add , , "Int. Apl.", 2100, vbRightJustify
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Cupón Ref", 1800, vbCenter
    
'
'    .Add , , "Corte", 1400
'    .Add , , "Prima/Descuento", 2200, vbRightJustify
'    .Add , , "Saldo Inversión", 2200, vbRightJustify
'    .Add , , "Inversión", 2200, vbRightJustify
'    .Add , , "Bancos", 2200, vbRightJustify
'    .Add , , "Int. Acumulado", 2200, vbRightJustify
'    .Add , , "Int. Periodo", 2200, vbRightJustify


End With

With lswFlujos_PyD.ColumnHeaders
    .Clear
    .Add , , "Corte", 1400
    .Add , , "Tipo", 2100, vbCenter
    .Add , , "Dias", 1200, vbCenter
    .Add , , "Monto", 2200, vbRightJustify
End With


With lswAd.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Monto", 2400, vbRightJustify
    .Add , , "Cuenta", 2100, vbCenter
    .Add , , "Tipo Doc.", 2200
    .Add , , "Detalle", 2100, vbCenter
    .Add , , "B. Id", 1200, vbCenter
    .Add , , "B. Estado", 1200, vbCenter
End With

With lswCom.ColumnHeaders
    .Clear
    .Add , , "Código", 1200, vbCenter
    .Add , , "Descripción", 3200
    .Add , , "Monto", 2400, vbRightJustify
    .Add , , "Cuenta", 2100, vbCenter
End With


With lswCierres.ColumnHeaders
    .Clear
    .Add , , "Periodo", 1200
    .Add , , "Valor", 2400, vbRightJustify
End With

cboBaseCalculo.Clear
cboBaseCalculo.AddItem "Base Actual"
cboBaseCalculo.ItemData(cboBaseCalculo.ListCount - 1) = CStr(365)
cboBaseCalculo.AddItem "Base 360"
cboBaseCalculo.ItemData(cboBaseCalculo.ListCount - 1) = CStr(360)
cboBaseCalculo.Text = "Base 360"


Call sbCbo_Copia(cboBaseCalculo, cboFi_BaseCalculo)
Call sbCbo_Copia(cboBaseCalculo, cboAc_BaseCalculo)

vPaso = False


End Sub


Private Sub sbAdquisicion_Load()
Call sbIVR_Transac_Load(lswAd, txtInversionId.Text, "T", "T")

txtAd_Requerido.Text = txtLiquidacion.Text

Dim i As Integer, pMonto As Currency

With lswAd.ListItems

pMonto = 0
For i = 1 To .Count
    pMonto = pMonto + CCur(.Item(i).SubItems(3))
Next i


pMonto = pMonto / fxSys_Tipo_Cambio_Apl(CCur(txtTipoCambio.Text))

txtAd_Registrado.Text = Format(pMonto, "Standard")

txtAd_Pendiente.Text = Format(CCur(txtAd_Requerido.Text) - pMonto, "Standard")

End With
 
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

Private Sub lswFi_Mov_DblClick()
If vPaso Then Exit Sub


Dim pTransacId As Long

gIVR_Transito.TituloId = txtInversionId.Text
gIVR_Transito.Codigo = txtInversionId.Text
gIVR_Transito.Divisa = cboDivisa.ItemData(cboDivisa.ListIndex)
gIVR_Transito.TipoCambio = CCur(txtTipoCambio.Text)
gIVR_Transito.Descripcion = "Id.:" & txtInversionId.Text & ", Operación: " & txtOperacion.Text

pTransacId = lswFi_Mov.SelectedItem.Text

    Dim frm As Form
    
    Call sbFormsCall("frmIVR_Proc_Fondos_Mov", 0, 0, 0, False, Me, False)
    Call sbFormActivo("frmIVR_Proc_Fondos_Mov", frm)
    
    Call frm.sbConsultaExterna(pTransacId)
    
    
    
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index > 0 Then
  If CLng(txtInversionId.Text) = 0 Then
        MsgBox "Consulte una inversión!", vbInformation
        tcMain.Item(0).Selected = True
        Exit Sub
  End If
End If

Dim i As Integer, pMonto As Currency
Dim pDebito As Currency, pCredito As Currency


Select Case Item.Index
    Case 0 'General
    Case 1 'Transac
        Call sbAdquisicion_Load
    Case 2 'Comisiones
        Call sbIVR_Comisiones_Load(lswCom, txtInversionId.Text)
        
        'Totales
        With lswCom.ListItems
        
        pMonto = 0
        For i = 1 To .Count
            pMonto = pMonto + CCur(.Item(i).SubItems(2))
        Next i
        End With
        
        txtComisionTotal.Text = Format(pMonto, "Standard")
                
        
    Case 3 'Flujos
        Call sbIVR_Flujos_Load(lswFlujos_PyD, txtInversionId.Text, "PYD")
        Call sbIVR_Flujos_Load(lswFlujos_Ingresos, txtInversionId.Text, "INT")
    
    Case 4 'Cierres
        Call sbIVR_Cierres_Load(lswCierres, txtInversionId.Text)
    
    Case 5 'Fondos
        Call sbIVR_Fondos_Load(lswFi_Mov, txtInversionId.Text, cboFi_Estado.Text, cboFi_Tipo.Text)
   
    Case 6 'Cupones
        Call sbIVR_Cupones_Load(lswCupones, txtInversionId.Text, Mid(cboCupones.Text, 1, 1))
    
    Case 7 'Asiento
        Call sbAsientos_Load
    
        Call sbIVR_Asiento_Load(lswAsiento, txtInversionId.Text)

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

End Select

End Sub

Private Sub sbAsientos_Load()
Dim pInicio As String, pCorte As String, pFiltro As String, pDetalle As String

pInicio = Format(dtpFechaCompra.Value, "yyyy/mm/dd") & " 00:00:00"
pCorte = Format(fxFechaServidor, "yyyy/mm/dd") & " 23:59:59"
        
pFiltro = ""
pDetalle = "Inversion Id: " & txtInversionId.Text
        
vPaso = True
Call sbIVR_Asientos_Main(lswAsientoMain, pInicio, pCorte, pFiltro, pDetalle)
        
vPaso = False
        
scAsientos.Caption = "Seleccione un asiento"
lswAsiento.ListItems.Clear
        
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

vPaso = True

vToken = ""

strSQL = "select COD_DIVISA  From vSys_Divisas  Where DIVISA_LOCAL = 1"
Call OpenRecordSet(rs, strSQL)
   vDivisaLocaL = Trim(rs!Cod_Divisa)
rs.Close


strSQL = "select  Tipo + '_' + rtrim(COD_INSTRUMENTO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & "  From IVR_INSTRUMENTOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboInstrumento, strSQL, False, True)

strSQL = "select  rtrim(COD_RECURSO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_FUENTE_RECURSOS" _
       & " Where ACTIVA = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboRecurso, strSQL, False, True)

strSQL = "select  rtrim(COD_EMISOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_EMISORES" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboEmisor, strSQL, False, True)


strSQL = "select  rtrim(COD_ADMINISTRADOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_ADMINISTRADOR" _
       & " Where ESTADO = 'A'" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboAdministrador, strSQL, False, True)

strSQL = "select  rtrim(COD_PORTAFOLIO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PORTAFOLIOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboPortafolio, strSQL, False, True)


strSQL = "select  rtrim(COD_CATEGORIA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_CATEGORIA_TIPOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)


strSQL = "select  rtrim(COD_RESERVA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_RESERVAS" _
       & " Where ACTIVA = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboReserva, strSQL, False, True)


strSQL = "select  rtrim(COD_PERIODICIDAD) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PERIODICIDAD" _
       & " Where ACTIVA = 1" _
       & " order by dias"
Call sbCbo_Llena_New(cboPeriodicidad, strSQL, False, True)


strSQL = "select rtrim(COD_DIVISA) AS 'Idx', rtrim(DESCRIPCION) as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

vPaso = False

Call sbInicializa

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(ByVal pTituloId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

mTituloIdConExt = 0

strSQL = "select * from vIVR_INVERSIONES" _
       & " Where Titulo_ID = " & pTituloId
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    txtInversionId.Text = rs!TITULO_ID
    
   
    txtEstado.Text = rs!Estado_Desc
    txtEstado.Tag = rs!Estado
    txtCartaIB.Text = rs!IB_Documento
    
    If rs!Estado = "S" Then
        txtCartaIB.Locked = False
    Else
        txtCartaIB.Locked = True
    End If
    
    tcMain.Item(0).Selected = True
    
    Call sbCboAsignaDato(cboInstrumento, rs!Instrumento_Desc, True, rs!Instrumento_Idx)
    Call sbCboAsignaDato(cboEmisor, rs!Emisor_Desc, True, rs!COD_EMISOR)
    Call sbCboAsignaDato(cboAdministrador, rs!Administrador_Desc, True, rs!Cod_Administrador)
    Call sbCboAsignaDato(cboClasificacion, rs!Categoria_Desc, True, rs!cod_Categoria)
    
    
    Call sbCboAsignaDato(cboRecurso, rs!Recurso_Desc, True, rs!cod_Recurso)
    Call sbCboAsignaDato(cboPortafolio, rs!Portafolio_Desc, True, rs!Cod_Portafolio)
    Call sbCboAsignaDato(cboReserva, rs!Reserva_Desc, True, rs!cod_Reserva)
    
    
    
    
    Call sbCboAsignaDato(cboDivisa, rs!Divisa_Desc, True, rs!Cod_Divisa)
    Call sbCboAsignaDato(cboPeriodicidad, rs!Periodicidad_Desc, True, rs!cod_Periodicidad)
    
    txtNotas.Text = rs!Detalle
    txtEjecutivo.Text = rs!Ejecutivo_Cuenta
    
    txtOperacion.Text = Trim(rs!Operacion)
    txtSerie.Text = Trim(rs!Serie)
    
    If rs!Base_Intereses = 360 Then
        cboBaseCalculo.Text = "Base 360"
    Else
        cboBaseCalculo.Text = "Base Actual"
    End If
    cboFi_BaseCalculo.Text = cboBaseCalculo.Text
    cboAc_BaseCalculo.Text = cboBaseCalculo.Text
    
    
    chkFi_CapiltalizaInt.Value = rs!Calcula_Intereses
    
    
    txtTipoCambio.Text = Format(rs!Tipo_Cambio, "###,##0.000000")
    txtValorFacial.Text = Format(rs!Valor_Facial, "Standard")
    txtValorTransado.Text = Format(rs!Valor_Transado, "Standard")
    txtCostoNeto.Text = Format(rs!Costo_Neto, "Standard")
    
    txtComisionAdm.Text = Format(rs!Comision_Puesto, "Standard")
    txtComisionBNV.Text = Format(rs!Comision_BNV, "Standard")
    
    txtIVA.Text = Format(rs!iva, "Standard")
    
    
    txtTasaNominal.Text = Format(rs!Tasa_Inicial, "###,##0.000000")
    
    
    txtRendNeto.Text = Format(rs!Rend_Neto, "###,##0.000000")
    txtRendNominal.Text = Format(rs!Rend_Nominal, "###,##0.000000")
    
    txtTEA.Text = Format(rs!Tasa_Efectiva, "###,##0.000000000")
    txtTIR.Text = Format(rs!TIR, "###,##0.00000000")
    
    If Not IsNull(rs!NDias_Inversion) Then
        txtDiasVencimiento.Text = rs!NDias_Inversion
    
    Else
        txtDiasVencimiento.Text = "999"
    End If
    txtDiasAcumulados.Text = IIf(IsNull(rs!Interes_Acum_Dias), 0, rs!Interes_Acum_Dias)
    txtIntAcumulado.Text = Format(rs!Interes_Acum_Monto, "Standard")
    
    If Not IsNull(rs!Fecha_pago) Then
        txtFechaUltPago.Text = Format(rs!Fecha_pago, "dd-mm-yyyy")
    Else
        txtFechaUltPago.Text = ""
    End If
    
    
    txtPyD.Text = Format(rs!PyD_BASE, "Standard")
    txtPrecio.Text = Format(rs!Precio, "###,##0.000000")
    
    txtFi_Inversion.Text = Format(rs!Valor_Facial, "Standard")
    txtFi_ParticipacionNo.Text = rs!Participacion_Numero
    txtFi_ParticipacionValor.Text = rs!Participacion_Valor
    
    txtAc_NoAcciones.Text = rs!Participacion_Numero
    txtAc_ValorAccion.Text = rs!Participacion_Valor
    txtAc_ValorTransado.Text = Format(rs!Monto_Principal, "Standard")
    
    
    dtpFechaCompra.Value = rs!Fecha_Compra
    
    
    If Not IsNull(rs!Fecha_Vencimiento) Then
        dtpFechaVence.Value = rs!Fecha_Vencimiento
    End If
    
    dtpFi_FechaCompra.Value = rs!Fecha_Compra
    dtpAc_FechaCompra.Value = rs!Fecha_Compra
    
    
    txtISRMonto.Text = Format(rs!ISR_MONTO, "Standard")
    
    txtValorLibros.Text = Format(rs!Valor_Libros, "Standard")
    txtValorLibros.ToolTipText = Format(rs!Fecha_Corte & "", "yyyy-mm-dd")
    
    'Si es un titulo refresca los calculos
    Select Case Mid(rs!Instrumento_Idx, 1, 1)
        Case "T"
            Call sbTitulo_Cal_Refresh
        Case "F"
            Call sbFondos_Cal_Refresh
        Case "A"
            Call sbAcciones_Cal_Refresh
        Case "C"
            Call sbCDP_Cal_Refresh
    End Select
'    gbToolAccion.Enabled = True
    
    btnToolAccion.Item(0).Visible = True
    btnToolAccion.Item(1).Visible = True
    btnToolAccion.Item(2).Visible = True
    btnToolAccion.Item(3).Visible = True
    btnToolAccion.Item(4).Visible = True
    btnToolAccion.Item(5).Visible = True
    btnToolAccion.Item(6).Visible = True
    
Else
  Me.MousePointer = vbDefault
  MsgBox "No se Localizó el registro!", vbExclamation
  Call sbInicializa
End If
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInicializa()

vPaso = True

txtInversionId.Text = "0"
txtEstado.Text = "Solicitud"
txtEstado.Tag = "S"

txtCartaIB.Text = ""
txtCartaIB.Locked = False

btnToolAccion.Item(0).Visible = False
btnToolAccion.Item(1).Visible = False
btnToolAccion.Item(2).Visible = False
btnToolAccion.Item(3).Visible = False
btnToolAccion.Item(4).Visible = False
btnToolAccion.Item(5).Visible = False
btnToolAccion.Item(6).Visible = False



txtCartaIB.Text = ""

vToken = Mid(glogon.Usuario, 1, 3) & Format(vFecha, "mmdd") & Format(Time, "mmss")

tcMain.Item(0).Selected = True

txtNotas.Text = ""
txtEjecutivo.Text = ""

txtOperacion.Text = ""
txtSerie.Text = ""
txtISIN.Text = ""
txtCupIp.Text = ""



txtTipoCambio.Text = "1"
txtValorFacial.Text = Format(0, "Standard")
txtValorTransado.Text = Format(0, "Standard")

txtComisionAdm.Text = Format(0, "Standard")
txtComisionBNV.Text = Format(0, "Standard")

txtTasaNominal.Text = "0"
txtTEA.Text = "0"
txtTIR.Text = "0"

txtRendNeto.Text = "0"
txtRendNominal.Text = "0"

txtLiquidacion.Text = "0"
txtIVA.Text = "0"
txtISRMonto.Text = "0"
txtDiasAcumulados.Text = "0"
txtDiasVencimiento.Text = "0"

txtPyD.Text = Format(0, "Standard")
txtPrecio.Text = Format(100, "Standard")
txtIntAcumulado.Text = Format(0, "Standard")

txtFi_Inversion.Text = Format(0, "Standard")
txtFi_ParticipacionNo.Text = Format(0, "Standard")
txtFi_ParticipacionValor.Text = Format(0, "Standard")

chkFi_CapiltalizaInt.Value = xtpChecked


txtAc_NoAcciones.Text = "0"
txtAc_ValorAccion.Text = Format(0, "Standard")
txtAc_ValorTransado.Text = Format(0, "Standard")
txtAc_ValorActual.Text = Format(0, "Standard")

dtpFechaCompra.Value = vFecha
dtpFechaVence.Value = vFecha

dtpFi_FechaCompra.Value = vFecha
dtpAc_FechaCompra.Value = vFecha

cboBaseCalculo.Text = "Base 360"

vPaso = False

Call cboPeriodicidad_Click
Call cboInstrumento_Click

If mTituloIdConExt > 0 Then
   Call sbConsulta(mTituloIdConExt)
End If

End Sub


Private Sub sbGuardar_Titulos()
On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pInstrumento As String

pInstrumento = Mid(cboInstrumento.ItemData(cboInstrumento.ListIndex), 3, 10)

strSQL = "exec spIVR_TITULOS_T_REGISTRA " & txtInversionId.Text _
       & ", '" & pInstrumento _
       & "','" & cboClasificacion.ItemData(cboClasificacion.ListIndex) _
       & "','" & cboEmisor.ItemData(cboEmisor.ListIndex) _
       & "','" & cboAdministrador.ItemData(cboAdministrador.ListIndex) _
       & "','" & cboPortafolio.ItemData(cboPortafolio.ListIndex) _
       & "','" & cboRecurso.ItemData(cboRecurso.ListIndex) _
       & "','" & cboReserva.ItemData(cboReserva.ListIndex) _
       & "','" & cboDivisa.ItemData(cboDivisa.ListIndex) _
       & "','" & cboPeriodicidad.ItemData(cboPeriodicidad.ListIndex) _
       & "','" & cboBaseCalculo.ItemData(cboBaseCalculo.ListIndex) _
       & "','" & txtOperacion.Text & "','" & txtSerie.Text _
       & "','" & txtISIN.Text & "','" & txtCupIp.Text _
       & "','" & txtEjecutivo.Text & "','" & txtNotas.Text _
       & "','" & Format(dtpFechaCompra.Value, "yyyy-mm-dd") _
       & "','" & Format(dtpFechaVence.Value, "yyyy-mm-dd") _
       & "', " & CCur(txtTipoCambio.Text) & " , " & CDbl(txtTasaNominal.Text) & " , " & CDbl(txtPrecio.Text) _
       & " , " & CCur(txtValorFacial.Text) & " , " & CCur(txtValorTransado.Text) _
       & " , " & CCur(txtComisionAdm.Text) & " , " & CCur(txtComisionBNV.Text) _
       & " , " & CCur(txtIVA.Text) & " ,'" & txtCartaIB.Text & "','" & glogon.Usuario & "'" _
       & " , 0, 0, 0" _
       & " , " & CDbl(txtRendNeto.Text) & ", " & CDbl(txtRendNominal.Text) _
       & " , " & CDbl(txtTIR.Text) & ", " & CDbl(txtTEA.Text) _
       & " , " & CCur(txtISRMonto.Text)
       
Call OpenRecordSet(rs, strSQL)

If rs!TITULO_ID > 0 Then
   Me.MousePointer = vbDefault
   Call sbConsulta(rs!TITULO_ID)
Else
   Me.MousePointer = vbDefault
   MsgBox rs!Mensaje, vbExclamation
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuardar_Fondos()
On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pInstrumento As String

pInstrumento = Mid(cboInstrumento.ItemData(cboInstrumento.ListIndex), 3, 10)


strSQL = "exec spIVR_TITULOS_F_REGISTRA " & txtInversionId.Text _
       & ", '" & pInstrumento _
       & "','" & cboClasificacion.ItemData(cboClasificacion.ListIndex) _
       & "','" & cboEmisor.ItemData(cboEmisor.ListIndex) _
       & "','" & cboAdministrador.ItemData(cboAdministrador.ListIndex) _
       & "','" & cboPortafolio.ItemData(cboPortafolio.ListIndex) _
       & "','" & cboRecurso.ItemData(cboRecurso.ListIndex) _
       & "','" & cboReserva.ItemData(cboReserva.ListIndex) _
       & "','" & cboDivisa.ItemData(cboDivisa.ListIndex) _
       & "','" & txtOperacion.Text & "','" & txtSerie.Text _
       & "','" & txtISIN.Text & "','" & txtCupIp.Text _
       & "','" & txtEjecutivo.Text & "','" & txtNotas.Text _
       & "','" & Format(dtpFi_FechaCompra.Value, "yyyy-mm-dd") _
       & "', " & CCur(txtTipoCambio.Text) & " , " & CCur(txtFi_Inversion.Text) _
       & " , " & CDbl(txtFi_ParticipacionValor.Text) & " , " & CDbl(txtFi_ParticipacionNo.Text) _
       & " , " & chkFi_CapiltalizaInt.Value & " ,'" & txtCartaIB.Text & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)



If rs!TITULO_ID > 0 Then
   Me.MousePointer = vbDefault
   Call sbConsulta(rs!TITULO_ID)
Else
   Me.MousePointer = vbDefault
   MsgBox rs!Mensaje, vbExclamation
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuardar_Acciones()
On Error GoTo vError

Me.MousePointer = vbHourglass




Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuardar()
On Error GoTo vError


'-Validacion
Dim vMensaje As String

vMensaje = ""

If cboDivisa.ItemData(cboDivisa.ListIndex) <> vDivisaLocaL _
    And CCur(txtTipoCambio.Text) = 1 Then
    vMensaje = vMensaje & vbCrLf & " - Revisar el Tipo de Cambio, la Transacción es en divisa foránea!"
End If

If Len(Trim(txtOperacion.Text)) = 0 Then
    vMensaje = vMensaje & vbCrLf & " - No se ha indicado un número de operación válido!"
End If

If Len(Trim(txtSerie.Text)) = 0 Then
    vMensaje = vMensaje & vbCrLf & " - No se ha indicado un número de serie válido!"
End If

If Len(vMensaje) > 0 Then
    MsgBox vMensaje, vbExclamation
End If


Me.MousePointer = vbHourglass


Select Case Mid(cboInstrumento.ItemData(cboInstrumento.ListIndex), 1, 1)
    Case "T" 'Titulos Valores
        Call sbGuardar_Titulos
    Case "F" 'Fondos de Inversion
        Call sbGuardar_Fondos
    Case "A" 'Acciones
        Call sbGuardar_Acciones
End Select



Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtComisionAdm_GotFocus()
On Error GoTo vError
    txtComisionAdm.Text = CCur(txtComisionAdm.Text)
vError:
End Sub

Private Sub txtComisionAdm_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub txtComisionBNV_GotFocus()
On Error GoTo vError
    txtComisionBNV.Text = CCur(txtComisionBNV.Text)
vError:
End Sub

Private Sub txtComisionBNV_LostFocus()
    Call sbTitulo_Cal_Refresh

End Sub


Private Sub txtFi_Inversion_LostFocus()
Call sbFondos_Cal_Refresh
End Sub

Private Sub txtFi_ParticipacionValor_LostFocus()
Call sbFondos_Cal_Refresh
End Sub

Private Sub txtInversionId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Inversión Id"
    gBusquedas.Col2Name = "Operación"
    gBusquedas.Col3Name = "Serie"
    gBusquedas.Consulta = "Select Titulo_Id, Operacion, Serie, Estado_Desc, Instrumento_Desc, Administrador_Desc, Recurso_Desc" _
                        & " from vIVR_INVERSIONES"
    gBusquedas.Columna = "Titulo_Id"
    gBusquedas.Orden = "Titulo_Id"

    frmBusquedas.Show vbModal
    
    If IsNumeric(gBusquedas.Resultado) Then
       txtInversionId.Text = gBusquedas.Resultado
       Call sbConsulta(txtInversionId)
    End If
    
End If
End Sub



Private Sub txtIVA_GotFocus()
On Error GoTo vError
    txtIVA.Text = CCur(txtIVA.Text)
vError:
End Sub

Private Sub txtIVA_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub txtPrecio_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub txtTasaNominal_GotFocus()
On Error GoTo vError
    txtTasaNominal.Text = CDbl(txtTasaNominal.Text)
vError:
End Sub

Private Sub txtTasaNominal_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub


Private Sub txtValorFacial_GotFocus()
On Error GoTo vError
    txtValorFacial.Text = CCur(txtValorFacial.Text)
vError:
End Sub

Private Sub txtValorFacial_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub

Private Sub txtValorTransado_GotFocus()
On Error GoTo vError
    txtValorTransado.Text = CCur(txtValorTransado.Text)
vError:

End Sub

Private Sub txtValorTransado_LostFocus()
    Call sbTitulo_Cal_Refresh
End Sub
