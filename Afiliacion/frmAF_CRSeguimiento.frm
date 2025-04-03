VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_CRSeguimiento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control/Seguimiento de Renuncias"
   ClientHeight    =   11595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11595
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox gbChecks 
      Height          =   1935
      Left            =   120
      TabIndex        =   77
      Top             =   8640
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   3413
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkC_Mortalidad 
         Height          =   375
         Left            =   480
         TabIndex        =   78
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Mortalidad ?"
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
         Value           =   2
      End
      Begin XtremeSuiteControls.CheckBox chkC_Reingreso 
         Height          =   375
         Left            =   480
         TabIndex        =   79
         Top             =   720
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reingreso Automático ?"
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
         Value           =   2
      End
      Begin XtremeSuiteControls.CheckBox chkC_Volver 
         Height          =   375
         Left            =   480
         TabIndex        =   80
         Top             =   1080
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Volver a Afiliarse a Futuro?"
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
         Value           =   2
      End
      Begin XtremeSuiteControls.CheckBox chkC_AumentoTasas 
         Height          =   375
         Left            =   480
         TabIndex        =   81
         Top             =   1440
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "No Aumento de Tasas?"
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
         Value           =   2
      End
      Begin XtremeSuiteControls.CheckBox chkC_Apl 
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   0
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Aplica Filtros Especiales?"
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
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbDetalle 
      Height          =   4332
      Left            =   3360
      TabIndex        =   30
      Top             =   4200
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   7641
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   3360
         TabIndex        =   31
         Top             =   240
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9123
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1440
         TabIndex        =   32
         Top             =   240
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
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   3492
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   10452
         _Version        =   1441793
         _ExtentX        =   18436
         _ExtentY        =   6159
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
         Item(0).Caption =   "Renuncia"
         Item(0).ControlCount=   18
         Item(0).Control(0)=   "Label2(0)"
         Item(0).Control(1)=   "Label2(1)"
         Item(0).Control(2)=   "Label2(2)"
         Item(0).Control(3)=   "Label2(3)"
         Item(0).Control(4)=   "Label2(4)"
         Item(0).Control(5)=   "txtEstado"
         Item(0).Control(6)=   "txtVencimiento"
         Item(0).Control(7)=   "txtPromotorDesc"
         Item(0).Control(8)=   "txtPromotorCod"
         Item(0).Control(9)=   "txtNotas"
         Item(0).Control(10)=   "chkMortalidad"
         Item(0).Control(11)=   "chkReingreso"
         Item(0).Control(12)=   "Label2(8)"
         Item(0).Control(13)=   "txtCausa"
         Item(0).Control(14)=   "txtTipoRenuncia"
         Item(0).Control(15)=   "lswMotivos"
         Item(0).Control(16)=   "Label2(9)"
         Item(0).Control(17)=   "chkVolver"
         Item(1).Caption =   "Control"
         Item(1).ControlCount=   13
         Item(1).Control(0)=   "FlatEdit1"
         Item(1).Control(1)=   "txtRegUser"
         Item(1).Control(2)=   "FlatEdit3"
         Item(1).Control(3)=   "txtResUser"
         Item(1).Control(4)=   "FlatEdit5"
         Item(1).Control(5)=   "txtEstadoControl"
         Item(1).Control(6)=   "FlatEdit7"
         Item(1).Control(7)=   "txtRegFecha"
         Item(1).Control(8)=   "FlatEdit9"
         Item(1).Control(9)=   "txtResFecha"
         Item(1).Control(10)=   "FlatEdit11"
         Item(1).Control(11)=   "txtVence"
         Item(1).Control(12)=   "lsw"
         Item(2).Caption =   "Gestión"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "gbGestion"
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   1692
            Left            =   -69880
            TabIndex        =   34
            Top             =   1800
            Visible         =   0   'False
            Width           =   10332
            _Version        =   1441793
            _ExtentX        =   18224
            _ExtentY        =   2984
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVence 
            Height          =   312
            Left            =   -62920
            TabIndex        =   35
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtResFecha 
            Height          =   312
            Left            =   -62920
            TabIndex        =   36
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRegFecha 
            Height          =   312
            Left            =   -62920
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstadoControl 
            Height          =   312
            Left            =   -67240
            TabIndex        =   38
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtResUser 
            Height          =   312
            Left            =   -67240
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRegUser 
            Height          =   312
            Left            =   -67240
            TabIndex        =   40
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkMortalidad 
            Height          =   372
            Left            =   2280
            TabIndex        =   41
            Top             =   3120
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Renuncia por Mortalidad"
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorCod 
            Height          =   312
            Left            =   2280
            TabIndex        =   42
            Top             =   2160
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
         Begin XtremeSuiteControls.FlatEdit txtEstado 
            Height          =   312
            Left            =   2280
            TabIndex        =   43
            Top             =   600
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVencimiento 
            Height          =   312
            Left            =   6720
            TabIndex        =   44
            Top             =   600
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorDesc 
            Height          =   312
            Left            =   3960
            TabIndex        =   45
            Top             =   2160
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   552
            Left            =   2280
            TabIndex        =   46
            Top             =   2520
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
            _ExtentY        =   974
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkReingreso 
            Height          =   372
            Left            =   5520
            TabIndex        =   47
            Top             =   3120
            Width           =   5652
            _Version        =   1441793
            _ExtentX        =   9970
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Aplica para Re-Ingreso Automático"
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   312
            Left            =   -68560
            TabIndex        =   48
            Top             =   600
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
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
            Text            =   "Registro.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   312
            Left            =   -68560
            TabIndex        =   49
            Top             =   960
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
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
            Text            =   "Resuelto.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit5 
            Height          =   312
            Left            =   -68560
            TabIndex        =   50
            Top             =   1320
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
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
            Text            =   "Estado.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit7 
            Height          =   312
            Left            =   -64360
            TabIndex        =   51
            Top             =   600
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
            Text            =   "Fecha.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit9 
            Height          =   312
            Left            =   -64360
            TabIndex        =   52
            Top             =   960
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
            Text            =   "Fecha.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit11 
            Height          =   312
            Left            =   -64360
            TabIndex        =   53
            Top             =   1320
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
            Text            =   "Vencimiento.: "
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox gbGestion 
            Height          =   2652
            Left            =   -69880
            TabIndex        =   59
            Top             =   480
            Visible         =   0   'False
            Width           =   9492
            _Version        =   1441793
            _ExtentX        =   16743
            _ExtentY        =   4678
            _StockProps     =   79
            Caption         =   "Gestión"
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
            BorderStyle     =   2
            Begin XtremeSuiteControls.ComboBox cboGestion 
               Height          =   312
               Left            =   1080
               TabIndex        =   60
               Top             =   480
               Width           =   5532
               _Version        =   1441793
               _ExtentX        =   9763
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
            Begin XtremeSuiteControls.ComboBox cboResolucion 
               Height          =   312
               Left            =   6600
               TabIndex        =   61
               Top             =   480
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4895
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
            Begin XtremeSuiteControls.FlatEdit txtGestionNota 
               Height          =   792
               Left            =   1080
               TabIndex        =   62
               Top             =   1080
               Width           =   8292
               _Version        =   1441793
               _ExtentX        =   14626
               _ExtentY        =   1397
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
            Begin XtremeSuiteControls.PushButton cmdGuardar 
               Height          =   492
               Left            =   7800
               TabIndex        =   0
               Top             =   2160
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Guardar"
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
               Picture         =   "frmAF_CRSeguimiento.frx":0000
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   5
               Left            =   1080
               TabIndex        =   65
               Top             =   840
               Width           =   1092
               _Version        =   1441793
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Notas: "
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   6
               Left            =   6600
               TabIndex        =   64
               Top             =   240
               Width           =   1932
               _Version        =   1441793
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Resolución:"
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   7
               Left            =   1080
               TabIndex        =   63
               Top             =   240
               Width           =   1092
               _Version        =   1441793
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Gestión: "
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
               WordWrap        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtCausa 
            Height          =   312
            Left            =   2280
            TabIndex        =   1
            Top             =   960
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTipoRenuncia 
            Height          =   312
            Left            =   4440
            TabIndex        =   2
            Top             =   600
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswMotivos 
            Height          =   732
            Left            =   2280
            TabIndex        =   74
            Top             =   1320
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
            _ExtentY        =   1291
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkVolver 
            Height          =   852
            Left            =   8880
            TabIndex        =   76
            Top             =   1320
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   1503
            _StockProps     =   79
            Caption         =   "Esta dispuesto a volver afiliarse a futuro ?"
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
            Enabled         =   0   'False
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   492
            Index           =   9
            Left            =   720
            TabIndex        =   75
            Top             =   1320
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Motivos Específicos"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   8
            Left            =   6840
            TabIndex        =   67
            Top             =   360
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Vencimiento"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   0
            Left            =   2280
            TabIndex        =   58
            Top             =   360
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Estado"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   1
            Left            =   720
            TabIndex        =   57
            Top             =   960
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Causa"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   2
            Left            =   4440
            TabIndex        =   56
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo Renuncia:"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   3
            Left            =   720
            TabIndex        =   55
            Top             =   2160
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Ejecutivo"
            BackColor       =   -2147483633
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   4
            Left            =   720
            TabIndex        =   54
            Top             =   2520
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas"
            BackColor       =   -2147483633
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBoleta 
         Height          =   312
         Left            =   8520
         TabIndex        =   72
         Top             =   240
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Boleta"
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
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   312
         Left            =   9240
         TabIndex        =   73
         Top             =   240
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Appearance      =   16
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro_RenunciaId 
      Height          =   312
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Text            =   "1"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro_RenunciaId 
      Height          =   312
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Text            =   "999999999"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConCedula 
      Height          =   312
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConNombre 
      Height          =   312
      Left            =   1440
      TabIndex        =   10
      Top             =   1440
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConUsuario 
      Height          =   312
      Left            =   1440
      TabIndex        =   12
      Top             =   1920
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConEjecutivo 
      Height          =   312
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox cboConCausa 
      Height          =   312
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   3012
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
   Begin XtremeSuiteControls.ComboBox cboConInstitucion 
      Height          =   312
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   3012
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
   Begin XtremeSuiteControls.ComboBox cboConProvincia 
      Height          =   312
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   3012
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
   Begin XtremeSuiteControls.ComboBox cboConTipoRen 
      Height          =   312
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   3012
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
   Begin XtremeSuiteControls.ComboBox cboConTipoFecha 
      Height          =   312
      Left            =   960
      TabIndex        =   24
      Top             =   6480
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.DateTimePicker dtpConInicio 
      Height          =   312
      Left            =   1680
      TabIndex        =   27
      Top             =   6960
      Width           =   1452
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
   Begin XtremeSuiteControls.DateTimePicker dtpConCorte 
      Height          =   312
      Left            =   1680
      TabIndex        =   28
      Top             =   7320
      Width           =   1452
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   1680
      TabIndex        =   29
      Top             =   7920
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmAF_CRSeguimiento.frx":0731
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4215
      Left            =   3360
      TabIndex        =   3
      Top             =   30
      Width           =   9855
      _Version        =   524288
      _ExtentX        =   17383
      _ExtentY        =   7435
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
      MaxCols         =   13
      SpreadDesigner  =   "frmAF_CRSeguimiento.frx":0E31
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboConZona 
      Height          =   312
      Left            =   120
      TabIndex        =   68
      Top             =   3000
      Width           =   3012
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
   Begin XtremeSuiteControls.ComboBox cboConEstado 
      Height          =   312
      Left            =   120
      TabIndex        =   70
      Top             =   6000
      Width           =   3012
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   71
      Top             =   5760
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   12
      Left            =   120
      TabIndex        =   69
      Top             =   2760
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Zonas:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   11
      Left            =   960
      TabIndex        =   26
      Top             =   7320
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corte:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   10
      Left            =   960
      TabIndex        =   25
      Top             =   6960
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Inicio:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Renuncia:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Provincia:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Entidad:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Causa:"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   372
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Ejecutivo"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   372
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Usuario"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   372
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nombre"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Identificación"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Renuncia Id"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPanel 
      Height          =   11190
      Left            =   0
      Picture         =   "frmAF_CRSeguimiento.frx":1771
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3330
   End
End
Attribute VB_Name = "frmAF_CRSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbRenuncia_Boleta()

On Error GoTo vError

Me.MousePointer = vbHourglass


 With frmContenedor.Crt
     .Reset

     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "Reportes del Módulo de Personas"
     .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(5) = "fxCodigoBarras = '*" & txtCodigo.Text & "*'"
     
     .Connect = glogon.ConectRPT
     
     .ReportFileName = SIFGlobal.fxPathReportes("Personas_CrBoletaRenuncias.rpt")
      strSQL = "{vAFI_Renuncia_Boleta.cod_renuncia} = " & txtCodigo.Text & ""
     
     .SelectionFormula = strSQL
     .PrintReport
 End With
 
Me.MousePointer = vbDefault
 
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnBoleta_Click()
 Call sbRenuncia_Boleta
End Sub

Private Sub btnBuscar_Click()
 Call sbBusca_Lista
End Sub

Private Sub sbGestion_Guardar()
Dim vIdGestion As Long

On Error GoTo vError

Me.MousePointer = vbHourglass
'
'vIdGestion = fxSegGestion(txtCodigo.Text, cboGestion.ItemData(cboGestion.ListIndex))
'
'If Mid(cboResolucion, 1, 1) <> "T" Then
'        strSQL = "Update afi_cr_renuncias set estado = '" & Mid(cboResolucion, 1, 1) & "'," _
'                & "resuelto_fecha = dbo.MyGetdate(),resuelto_user = '" & glogon.Usuario & "'" _
'                & " where cod_renuncia = " & txtCodigo.Text & " and cedula = '" & txtCedula.Text & "'"
'        Call ConectionExecute(strSQL)
'
'        Call sbgAFIBitacora("06", "Modifica Seguimiento Renuncia Persona:  " & Trim(txtNombre.Text) & "Ced: " & txtCedula.Text, Trim(txtCedula))
'
'End If
'
''Registra el Seguimiento
'strSQL = "insert afi_cr_seguimiento(id,cod_renuncia,cod_gestion,estado,fecha,usuario,notas)" _
'        & " values(" & vIdGestion & ",'" & txtCodigo.Text & "','" & cboGestion.ItemData(cboGestion.ListIndex) & "'," _
'        & "'" & Mid(cboResolucion, 1, 1) & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtGestionNota.Text & "')"
'

'spAFI_Renuncia_CambioEstado
'      @RenunciaId  int = NULL, @Estado varchar(1) = NULL, @Gestion varchar(10) = '01'
'    , @Notas   varchar(500) = NULL,     @Usuario  varchar(30) = NULL,   @Equipo varchar(50) = NULL
'    , @Version varchar(50) = NULL
strSQL = "exec spAFI_Renuncia_CambioEstado " & txtCodigo.Text & ", '" & Mid(cboResolucion, 1, 1) & "', '" & cboGestion.ItemData(cboGestion.ListIndex) _
       & "', '" & txtGestionNota.Text & "', '" & glogon.Usuario & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
Call ConectionExecute(strSQL)




txtGestionNota.Text = ""

Me.MousePointer = vbDefault
MsgBox "Información Registrada satisfactoriamente", vbInformation

Call sbConsulta_Renuncia(txtCodigo.Text)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub


Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 13
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "No. Renuncia"
    vHeaders.Headers(3) = "Estado"
    vHeaders.Headers(4) = "Identifiación"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Tipo Renuncia"
    vHeaders.Headers(7) = "Vencimiento"
    vHeaders.Headers(8) = "Causa"
    vHeaders.Headers(9) = "Ejecutivo"
    vHeaders.Headers(10) = "Registro [Usuario]"
    vHeaders.Headers(11) = "Registro [Fecha]"
    vHeaders.Headers(12) = "Resuelve [Usuario]"
    vHeaders.Headers(13) = "Resuelve [Fecha]"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Renuncias_Listado")
End Sub

Private Sub chkC_Apl_Click()

chkC_Mortalidad.Value = xtpGrayed
chkC_Reingreso.Value = xtpGrayed
chkC_Volver.Value = xtpGrayed
chkC_AumentoTasas.Value = xtpGrayed

If chkC_Apl.Value = xtpChecked Then
    chkC_Mortalidad.Enabled = True
    chkC_Reingreso.Enabled = True
    chkC_Volver.Enabled = True
    chkC_AumentoTasas.Enabled = True
Else
    chkC_Mortalidad.Enabled = False
    chkC_Reingreso.Enabled = False
    chkC_Volver.Enabled = False
    chkC_AumentoTasas.Enabled = False
End If

End Sub

Private Sub cmdGuardar_Click()
    
If IsNumeric(txtCodigo.Text) Then
    Call sbGestion_Guardar
End If

End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 1

gbDetalle.BorderStyle = xtpFrameNone

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Fecha", 1800
    .Add , , "Usuario", 1800
    .Add , , "Estado", 1400
    .Add , , "Notas", 3270
End With

Call chkC_Apl_Click

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBusca_Lista()
 
On Error GoTo vError
  
Me.MousePointer = vbHourglass
  

vGrid.MaxRows = 0
vPaso = True

strSQL = "select '' as 'Btn', R.cod_renuncia" _
        & ", R.Estado_Desc, R.cedula, R.Nombre, R.Tipo_Renuncia" _
        & ", R.vencimiento,  R.Causa_Desc, R.Ejecutivo_Desc" _
        & ", R.registro_user, R.registro_Fecha, R.Resuelto_User, R.Resuelto_Fecha" _
        & " from vAFI_Renuncias R" _
        & " Where R.Cod_Renuncia between " & txtFiltro_RenunciaId(0).Text & " and " & txtFiltro_RenunciaId(1).Text
        
If cboConEstado.Text <> "TODOS" Then
      strSQL = strSQL & " and R.Estado = '" & cboConEstado.ItemData(cboConEstado.ListIndex) & "'"
End If
  
If cboConTipoRen.Text <> "TODAS" Then
      strSQL = strSQL & " and R.Tipo = '" & Mid(cboConTipoRen.Text, 1, 1) & "'"
End If
  
If Len(Trim(txtConCedula.Text)) > 0 Then
      strSQL = strSQL & " and R.cedula like '%" & txtConCedula.Text & "%'"
End If

If Len(Trim(txtConNombre.Text)) > 0 Then
      strSQL = strSQL & " and R.Nombre like '%" & txtConNombre.Text & "%'"
End If

If Len(Trim(txtConUsuario.Text)) > 0 Then
      strSQL = strSQL & " and R.registro_user like '%" & txtConUsuario.Text & "%'"
End If

If Len(Trim(txtConEjecutivo.Text)) > 0 Then
      strSQL = strSQL & " and R.Ejecutivo_Desc like '%" & txtConEjecutivo.Text & "%'"
End If

If cboConCausa.Text <> "TODOS" Then
      strSQL = strSQL & " and R.Id_Causa = " & cboConCausa.ItemData(cboConCausa.ListIndex)
End If

If cboConInstitucion.Text <> "TODOS" Then
      strSQL = strSQL & " and R.cod_Institucion = " & cboConInstitucion.ItemData(cboConInstitucion.ListIndex)
End If

If cboConProvincia.Text <> "TODOS" Then
      strSQL = strSQL & " and R.Provincia = '" & cboConProvincia.ItemData(cboConProvincia.ListIndex) & "'"
End If

If cboConZona.Text <> "TODOS" Then
     strSQL = strSQL & " and dbo.fxAfi_Zonas_Aplica('" & cboConZona.ItemData(cboConZona.ListIndex) _
            & "','" & glogon.Usuario & "', R.cod_Institucion, R.UP) = 1"
End If

Select Case Mid(cboConTipoFecha.Text, 1, 3)
    Case "Reg"
    strSQL = strSQL & " and R.registro_Fecha between '" & Format(dtpConInicio.Value, "yyyy/MM/dd") & " 00:00:00'" _
           & " and '" & Format(dtpConCorte.Value, "yyyy/MM/dd") & " 23:59:59'"
    Case "Ven"
    strSQL = strSQL & " and R.Vencimiento between '" & Format(dtpConInicio.Value, "yyyy/MM/dd") & " 00:00:00'" _
           & " and '" & Format(dtpConCorte.Value, "yyyy/MM/dd") & " 23:59:59'"

    Case "Res"
    strSQL = strSQL & " and R.Resuelto_Fecha between '" & Format(dtpConInicio.Value, "yyyy/MM/dd") & " 00:00:00'" _
           & " and '" & Format(dtpConCorte.Value, "yyyy/MM/dd") & " 23:59:59'"

End Select

If chkC_Apl.Value = xtpChecked Then
   If chkC_Mortalidad.Value <> xtpGrayed Then
        strSQL = strSQL & " and R.Mortalidad = " & chkC_Mortalidad.Value
   End If

   If chkC_Reingreso.Value <> xtpGrayed Then
        strSQL = strSQL & " and R.APLICA_REINGRESO = " & chkC_Reingreso.Value
   End If

   If chkC_Volver.Value <> xtpGrayed Then
        strSQL = strSQL & " and R.Volver = " & chkC_Volver.Value
   End If

   If chkC_AumentoTasas.Value <> xtpGrayed Then
        strSQL = strSQL & " and R.Aumenta_Puntos = " & chkC_AumentoTasas.Value
   End If

End If



Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)
'Limpia Ultima linea en Blanco
vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbConsulta_Renuncia(pRenuncia As Long)
Dim itmX As ListViewItem

On Error GoTo vError
 
Me.MousePointer = vbHourglass


gbGestion.Enabled = False

strSQL = "Select R.*,rTrim(C.Descripcion) as 'CausaX',S.nombre" _
       & ",isnull(P.id_promotor,0) as 'Id_Promotor',isnull(P.nombre,'AFILIACION UNIVERSAL') as PromotorX" _
       & " from afi_cr_renuncias R inner join causas_renuncias C on R.id_causa = C.id_causa" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join Promotores P on R.id_Promotor = P.id_Promotor" _
       & " where R.cod_renuncia = " & pRenuncia
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 
 tcMain.Item(0).Selected = True
 
 txtCodigo.Text = rs!Cod_Renuncia
 txtNombre.Text = rs!Nombre
 txtCedula.Text = rs!Cedula
 txtVencimiento.Text = rs!Vencimiento
 
 txtCausa.Text = rs!CausaX

 Select Case rs!Tipo
    Case "P"
       txtTipoRenuncia.Text = "PATRONAL"
    Case "A"
       txtTipoRenuncia.Text = "ASOCIACION"
  End Select
  
  Select Case rs!Estado
   Case "T"
    txtEstado.Text = "Transito"
   Case "R"
    txtEstado.Text = "Rescatada"
   Case "P"
    txtEstado.Text = "Perdida"
   Case "V"
    txtEstado.Text = "Vencida"
  End Select
  
  txtPromotorDesc.Text = rs!PromotorX
  txtPromotorCod.Text = rs!ID_PROMOTOR
  
  chkReingreso.Value = rs!Aplica_Reingreso
  chkMortalidad.Value = rs!Mortalidad
  chkVolver.Value = rs!Volver
  
  txtNotas.Text = rs!Notas
  
  'Control
     txtRegUser.Text = IIf(Not IsNull(rs!registro_user), rs!registro_user, "")
     txtRegFecha.Text = IIf(Not IsNull(rs!Registro_Fecha), Format(rs!Registro_Fecha, "mm/dd/yyyy"), "")
     txtResFecha.Text = IIf(Not IsNull(rs!resuelto_fecha), Format(rs!resuelto_fecha, "mm/dd/yyyy"), "")
     txtResUser.Text = IIf(Not IsNull(rs!resuelto_user), rs!resuelto_user, "")
     Select Case rs!Estado
         Case "T"
           txtEstadoControl.Text = "Transito"
           gbGestion.Enabled = True
         Case "P"
           txtEstadoControl.Text = "Perdida"
        Case "R"
           txtEstadoControl.Text = "Rescatada"
        Case "V"
           txtEstadoControl.Text = "Vencida"
     End Select
     
     txtVencimiento.Text = IIf(Not IsNull(rs!Vencimiento), Format(rs!Vencimiento, "mm/dd/yyyy"), "")
     
     cboResolucion.Text = txtEstadoControl.Text

End If
rs.Close

'Consulta Motivos
lswMotivos.ListItems.Clear
With lswMotivos.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6200
End With

  
strSQL = "exec spAFI_CR_Motivos_Consulta " & pRenuncia & ",1"
Call OpenRecordSet(rs, strSQL)

vPaso = True
With lswMotivos.ListItems
   .Clear
   Do While Not rs.EOF
    Set itmX = .Add(, , rs!Descripcion)
        itmX.Tag = rs!Cod_Motivo
        
        itmX.Checked = IIf((rs!asignado = 1), True, False)
        
    rs.MoveNext
   Loop
   rs.Close
End With
vPaso = False


'Lista de Seguimiento
strSQL = "select * from afi_cr_seguimiento where cod_renuncia = " & pRenuncia & ""

lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Id)
   
   itmX.SubItems(1) = IIf(Not IsNull(rs!fecha), Format(rs!fecha, "mm/dd/yyyy"), "")
   itmX.SubItems(2) = IIf(Not IsNull(rs!Usuario), rs!Usuario, "")
   itmX.SubItems(4) = IIf(Not IsNull(rs!Notas), rs!Notas, "")
   Select Case rs!Estado
         Case "T"
           itmX.SubItems(3) = "Transito"
         Case "P"
           itmX.SubItems(3) = "Perdida"
        Case "R"
           itmX.SubItems(3) = "Rescatada"
        Case "V"
           itmX.SubItems(3) = "Vencida"
   End Select
   
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Resize()
Dim pMinHeight As Long, pMinWidth As Long
Dim pHeight As Long, pWidth As Long

On Error Resume Next

pMinHeight = 9036
pMinWidth = 13512

If Me.Height < pMinHeight Then
   pHeight = pMinHeight
Else
   pHeight = Me.Height
End If

If Me.Width < pMinWidth Then
   pWidth = pMinWidth
Else
   pWidth = Me.Width
End If


imgPanel.Height = pHeight

vGrid.Height = pHeight - (gbDetalle.Height + 650)
vGrid.Width = pWidth - (vGrid.Left + 150)


gbDetalle.top = vGrid.Height + 50
gbDetalle.Left = vGrid.Left + ((vGrid.Width - gbDetalle.Width) / 2)


End Sub


Private Function fxSegGestion(vRenuncia As Long, vCodg As String) As Integer

strSQL = "Select isnull(Max(id),0) as consecutivo from afi_cr_seguimiento" _
        & " where cod_renuncia = " & vRenuncia & " and cod_gestion = '" & vCodg & "'"
        
Call OpenRecordSet(rs, strSQL)
  fxSegGestion = rs!Consecutivo + 1
rs.Close

End Function



Private Sub sbInicializa()
 
On Error GoTo vError
 
Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

tcMain.Item(0).Selected = True

cboConTipoRen.Clear
cboConTipoRen.AddItem "TODAS"
cboConTipoRen.AddItem "Asociación"
cboConTipoRen.AddItem "Patronal"
cboConTipoRen.Text = "TODAS"

cboConTipoFecha.Clear
cboConTipoFecha.AddItem "Registro"
cboConTipoFecha.AddItem "Vencimiento"
cboConTipoFecha.AddItem "Resolución"
cboConTipoFecha.Text = "Registro"

cboResolucion.Clear
cboResolucion.AddItem "Transito"
cboResolucion.AddItem "Rescatada"
cboResolucion.AddItem "Perdida"
cboResolucion.Text = "Transito"

cboConEstado.Clear
cboConEstado.AddItem "TODOS"
cboConEstado.AddItem "Transito"
cboConEstado.ItemData(cboConEstado.ListCount - 1) = "T"
cboConEstado.AddItem "Rescatada"
cboConEstado.ItemData(cboConEstado.ListCount - 1) = "R"
cboConEstado.AddItem "Perdida"
cboConEstado.ItemData(cboConEstado.ListCount - 1) = "P"
cboConEstado.AddItem "Vencida"
cboConEstado.ItemData(cboConEstado.ListCount - 1) = "V"
cboConEstado.AddItem "Pendiente"
cboConEstado.ItemData(cboConEstado.ListCount - 1) = "E"
cboConEstado.Text = "Transito"



strSQL = "select rtrim(cod_gestion) as 'IdX' , rtrim(descripcion) as 'ItmX' from afi_cr_gestiones"
Call sbCbo_Llena_New(cboGestion, strSQL, False, True)

'Busquedas
strSQL = "select id_causa as IdX, rtrim(descripcion) as itmX from causas_renuncias WHERE ACTIVO = 1"
Call sbCbo_Llena_New(cboConCausa, strSQL, True, True)

strSQL = "select cod_institucion as 'IdX' , rtrim(descripcion) as 'ItmX' from Instituciones order by descripcion"
Call sbCbo_Llena_New(cboConInstitucion, strSQL, True, True)

strSQL = "select Provincia as 'IdX' , rtrim(descripcion) as 'ItmX' from Provincias order by Provincia"
Call sbCbo_Llena_New(cboConProvincia, strSQL, True, True)


strSQL = "select COD_ZONA as 'IdX' , rtrim(descripcion) as 'ItmX' from AFI_ZONAS order by descripcion"
Call sbCbo_Llena_New(cboConZona, strSQL, True, True)


dtpConCorte.Value = fxFechaServidor
dtpConInicio.Value = DateAdd("d", -45, dtpConCorte.Value)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswMotivos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAFI_CR_Motivos_Registra " & txtCodigo.Text & ",'" & Item.Tag & "','" _
       & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub TimerX_Timer()

Timerx.Interval = 0
Timerx.Enabled = False

Call sbInicializa

End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

vGrid.Row = Row
vGrid.col = 2

Call sbConsulta_Renuncia(vGrid.Text)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
